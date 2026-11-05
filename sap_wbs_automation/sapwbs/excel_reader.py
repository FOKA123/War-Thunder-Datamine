"""Citirea devizului: detecția dinamică a antetelor și extragerea normelor.

Nicio poziție de coloană sau rând nu este fixă. Pentru fiecare foaie:

1. se scanează primele ``header_scan_rows`` rânduri;
2. fiecare celulă este comparată (normalizat, fără diacritice) cu alias-urile
   din configurație; potrivirea exactă valorează mai mult decât cea parțială;
3. rândul care conține toate coloanele obligatorii cu scorul maxim devine antet;
4. coloana de cod este validată și după conținut (regex pe valori), cu
   fallback pe coloana care conține cele mai multe coduri valide.

Foile care nu au coloanele obligatorii sunt ignorate (și apar în log).
"""
from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from .config import AppConfig, FieldSource
from .model import (ERROR, INFO, WARNING, HeaderMap, Issue, PtInfo, ReadResult,
                    SourceRow, WorkbookContext)
from .textutil import clean_spaces, first_line, norm, parse_number

log = logging.getLogger(__name__)

EXACT, PREFIX, WORD = 100, 60, 30


# --------------------------------------------------------------------------- #
#  Grilă în memorie
# --------------------------------------------------------------------------- #
class SheetGrid:
    """Valorile unei foi, încărcate o singură dată (indecși 0-based)."""

    def __init__(self, title: str, rows: Sequence[Sequence[Any]]):
        self.title = title
        self.rows: List[Tuple[Any, ...]] = [tuple(r) for r in rows]

    def __len__(self) -> int:
        return len(self.rows)

    def value(self, row: int, col: int) -> Any:
        if 0 <= row < len(self.rows):
            data = self.rows[row]
            if 0 <= col < len(data):
                return data[col]
        return None

    def iter_cells(self) -> Iterable[Tuple[int, int, Any]]:
        for r, data in enumerate(self.rows):
            for c, val in enumerate(data):
                if val is not None and str(val).strip() != "":
                    yield r, c, val


def to_text(value: Any) -> str:
    """Valoare de celulă -> text (12.0 -> "12")."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return clean_spaces(value)


# --------------------------------------------------------------------------- #
#  Detecția antetului (Rule-Based Engine)
# --------------------------------------------------------------------------- #
def score_header(header_text: Any, aliases: Sequence[str]) -> int:
    """Scorul potrivirii dintre textul unui antet și lista de alias-uri."""
    header = norm(header_text)
    if not header:
        return 0
    best = 0
    for alias in aliases:
        alias_n = norm(alias)
        if not alias_n:
            continue
        if header == alias_n:
            return EXACT
        if len(alias_n) <= 2:  # alias-urile foarte scurte ("E") doar exact
            continue
        if header.startswith(alias_n + " "):
            best = max(best, PREFIX)
        elif f" {alias_n} " in f" {header} ":
            best = max(best, WORD)
    return best


def match_columns(header_values: Sequence[Any], columns: Dict[str, List[str]],
                  priority: Sequence[str] = ()) -> Dict[str, Tuple[int, int]]:
    """Asociază coloanele logice cu indecși de coloană (greedy, după scor).

    Întoarce {coloană_logică: (index_coloană, scor)}; o coloană fizică este
    folosită o singură dată.
    """
    order = {name: i for i, name in enumerate(priority)}
    candidates = []
    for logical, aliases in columns.items():
        for col, value in enumerate(header_values):
            score = score_header(value, aliases)
            if score:
                candidates.append((score, -order.get(logical, len(order)), logical, col))
    candidates.sort(key=lambda item: (item[0], item[1]), reverse=True)
    assigned: Dict[str, Tuple[int, int]] = {}
    used = set()
    for score, _prio, logical, col in candidates:
        if logical in assigned or col in used:
            continue
        assigned[logical] = (col, score)
        used.add(col)
    return assigned


def match_pt_columns(header_values: Sequence[Any], pt_regex: str) -> Dict[int, int]:
    pattern = re.compile(pt_regex, re.IGNORECASE)
    result: Dict[int, int] = {}
    for col, value in enumerate(header_values):
        match = pattern.match(first_line(value))
        if match:
            result.setdefault(int(match.group(1)), col)
    return result


def detect_header(grid: SheetGrid, cfg: AppConfig) -> Optional[HeaderMap]:
    """Găsește rândul de antet cu scorul maxim care conține coloanele obligatorii."""
    ecfg = cfg.excel
    priority = list(ecfg.required_columns) + [c for c in ecfg.columns if c not in ecfg.required_columns]
    best: Optional[HeaderMap] = None
    for r in range(min(len(grid), ecfg.header_scan_rows)):
        values = grid.rows[r]
        assigned = match_columns(values, ecfg.columns, priority)
        if not all(req in assigned for req in ecfg.required_columns):
            continue
        pts = match_pt_columns(values, ecfg.pt_column_regex)
        score = sum(s for _, s in assigned.values()) + 5 * len(pts)
        if best is None or score > best.score:
            best = HeaderMap(
                sheet=grid.title,
                header_row=r + 1,
                columns={k: c for k, (c, _) in assigned.items()},
                pt_columns=pts,
                score=score,
            )
    return best


def validate_code_column(grid: SheetGrid, header: HeaderMap, code_regex: str,
                         sample_rows: int = 500) -> Optional[str]:
    """Verifică după conținut coloana de cod; o corectează dacă e cazul.

    Întoarce un mesaj de avertizare când coloana a fost schimbată.
    """
    pattern = re.compile(code_regex)
    start = header.header_row  # index 0-based al primului rând de date
    end = min(len(grid), start + sample_rows)

    def matches(col: int) -> int:
        return sum(1 for r in range(start, end) if pattern.match(to_text(grid.value(r, col)).upper()))

    current = header.columns["code"]
    current_hits = matches(current)
    if current_hits >= 3:
        return None
    used = set(header.columns.values()) | set(header.pt_columns.values())
    width = max((len(grid.rows[r]) for r in range(start, end)), default=0)
    best_col, best_hits = current, current_hits
    for col in range(width):
        if col in used:
            continue
        hits = matches(col)
        if hits > best_hits:
            best_col, best_hits = col, hits
    if best_col != current:
        header.columns["code"] = best_col
        return (f"Coloana de cod din antet nu conține coduri valide; "
                f"se folosește coloana {header.letter('code')} ({best_hits} coduri găsite)")
    return None


# --------------------------------------------------------------------------- #
#  Căutare după etichetă
# --------------------------------------------------------------------------- #
def find_label(grid: SheetGrid, labels: Sequence[str]) -> Optional[Tuple[int, int]]:
    """Prima celulă (parcurgere pe rânduri) al cărei text corespunde unei etichete."""
    wanted = [norm(lbl) for lbl in labels if norm(lbl)]
    for r, c, value in grid.iter_cells():
        text = norm(value)
        if any(text == w or text.startswith(w + " ") for w in wanted) and len(text) <= 60:
            return r, c
    return None


def value_right_of(grid: SheetGrid, r: int, c: int, max_right: int = 3) -> str:
    """Prima valoare ne-goală din dreapta etichetei (cel mult ``max_right`` celule)."""
    for offset in range(1, max_right + 1):
        text = to_text(grid.value(r, c + offset))
        if text:
            return text
    return ""


def find_label_value(grid: SheetGrid, labels: Sequence[str], max_right: int = 3) -> Tuple[str, str]:
    pos = find_label(grid, labels)
    if pos is None:
        return "", ""
    from openpyxl.utils import get_column_letter

    value = value_right_of(grid, pos[0], pos[1], max_right)
    return value, f"{grid.title}!{get_column_letter(pos[1] + 1)}{pos[0] + 1}"


# --------------------------------------------------------------------------- #
#  Cititorul principal
# --------------------------------------------------------------------------- #
class WorkbookReader:
    """Citește devizul și produce rândurile sursă + contextul proiectului."""

    def __init__(self, path: Path, cfg: AppConfig):
        self.path = Path(path)
        self.cfg = cfg
        self.issues: List[Issue] = []
        self._grids: Dict[str, SheetGrid] = {}
        self._wb = None

    # -- infrastructură ---------------------------------------------------- #
    def _open(self):
        import warnings

        import openpyxl

        # openpyxl avertizează pentru extensii Excel necunoscute (validări de date etc.)
        warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
        log.info("Deschid devizul: %s", self.path)
        # data_only=True: se citesc valorile calculate (salvate de Excel), nu formulele
        self._wb = openpyxl.load_workbook(self.path, data_only=True, read_only=True)
        return self._wb

    def _grid(self, title: str) -> Optional[SheetGrid]:
        if title in self._grids:
            return self._grids[title]
        if self._wb is None or title not in self._wb.sheetnames:
            return None
        ws = self._wb[title]
        rows = ws.iter_rows(max_col=self.cfg.excel.max_scan_col, values_only=True)
        grid = SheetGrid(title, list(rows))
        self._grids[title] = grid
        return grid

    def _sheet_selected(self, title: str, state: str) -> bool:
        ecfg = self.cfg.excel
        if ecfg.skip_hidden_sheets and state != "visible":
            return False
        if ecfg.include_sheets and not any(re.search(p, title, re.I) for p in ecfg.include_sheets):
            return False
        return not any(re.search(p, title, re.I) for p in ecfg.exclude_sheets)

    def _issue(self, severity: str, message: str, sheet: str = "", row: Optional[int] = None,
               code: str = "") -> None:
        self.issues.append(Issue(severity, message, sheet, row, code))
        level = {INFO: logging.INFO, WARNING: logging.WARNING, ERROR: logging.ERROR}[severity]
        where = f"{sheet}!{row}" if sheet and row else sheet
        log.log(level, "%s%s", f"[{where}] " if where else "", message)

    # -- API ---------------------------------------------------------------- #
    def read(self) -> ReadResult:
        wb = self._open()
        try:
            headers: List[HeaderMap] = []
            rows: List[SourceRow] = []
            for ws in wb.worksheets:
                state = getattr(ws, "sheet_state", "visible")
                if not self._sheet_selected(ws.title, state):
                    log.debug("Foaie ignorată (filtru/ascunsă): %s", ws.title)
                    continue
                grid = self._grid(ws.title)
                header = detect_header(grid, self.cfg) if grid else None
                if header is None:
                    log.debug("Foaie fără coloanele obligatorii, ignorată: %s", ws.title)
                    continue
                warning = validate_code_column(grid, header, self.cfg.excel.code_regex)
                if warning:
                    self._issue(WARNING, warning, ws.title)
                headers.append(header)
                log.info(
                    "Foaia '%s': antet pe rândul %d | cod=%s cant=%s retea=%s activitate=%s partea=%s | PT: %s",
                    ws.title, header.header_row, header.letter("code"), header.letter("quantity"),
                    header.letter("network"), header.letter("activity"), header.letter("part"),
                    ",".join(f"PT{n}" for n in sorted(header.pt_columns)) or "-",
                )
                rows.extend(self._extract_rows(grid, header))
            if not headers:
                self._issue(ERROR, "Nu am găsit nicio foaie cu coloanele obligatorii "
                                   f"{self.cfg.excel.required_columns}. Verificați alias-urile din config.")
            context = self._read_context()
            return ReadResult(rows=rows, headers=headers, context=context, issues=self.issues)
        finally:
            wb.close()

    # -- extragere rânduri -------------------------------------------------- #
    def _extract_rows(self, grid: SheetGrid, header: HeaderMap) -> List[SourceRow]:
        ecfg = self.cfg.excel
        code_re = re.compile(ecfg.code_regex)
        cols = header.columns

        def cell(r: int, logical: str) -> Any:
            col = cols.get(logical)
            return grid.value(r, col) if col is not None else None

        result: List[SourceRow] = []
        empty_streak = 0
        missing_cache = 0
        numeric_seen = 0
        for r in range(header.header_row, len(grid)):
            code = to_text(cell(r, "code")).upper()
            if not code:
                empty_streak += 1
                if empty_streak >= ecfg.stop_after_empty_rows:
                    break
                continue
            empty_streak = 0
            if not code_re.match(code):
                continue  # subtotaluri, titluri de capitol etc.
            raw_qty = cell(r, "quantity")
            qty = parse_number(raw_qty)
            if raw_qty is None:
                missing_cache += 1
            elif qty is not None:
                numeric_seen += 1
            if qty is None or abs(qty) < 1e-9:
                continue
            pts: Dict[int, float] = {}
            for pt_no, col in header.pt_columns.items():
                val = parse_number(grid.value(r, col))
                if val is not None and abs(val) > 1e-9:
                    pts[pt_no] = val
            result.append(
                SourceRow(
                    sheet=grid.title,
                    row=r + 1,
                    code=code,
                    quantity=qty,
                    pt_quantities=pts,
                    network_raw=to_text(cell(r, "network")),
                    activity_raw=to_text(cell(r, "activity")),
                    part=to_text(cell(r, "part")).upper(),
                    description=to_text(cell(r, "description")),
                    install_type=to_text(cell(r, "install_type")).upper(),
                    override=to_text(cell(r, "override")),
                )
            )
        if numeric_seen == 0 and missing_cache > 0:
            self._issue(
                ERROR,
                "Coloana de cantitate nu are valori calculate (formule nesalvate de Excel). "
                "Deschideți și salvați fișierul în Excel sau rulați cu --recalculeaza.",
                grid.title,
            )
        log.info("Foaia '%s': %d rânduri cu cantitate nenulă", grid.title, len(result))
        return result

    # -- context ------------------------------------------------------------- #
    def _read_context(self) -> WorkbookContext:
        rcfg = self.cfg.rules
        ctx = WorkbookContext()
        grid = self._grid(rcfg.context_sheet)
        if grid is None:
            self._issue(WARNING, f"Foaia de context '{rcfg.context_sheet}' nu există; "
                                 "se folosesc doar regulile din config.")
        else:
            ctx.voltage, origin_v = find_label_value(grid, rcfg.voltage_labels)
            ctx.solution, origin_s = find_label_value(grid, rcfg.solution_labels)
            log.info("Context: nivel tensiune='%s' (%s), soluție='%s' (%s)",
                     ctx.voltage, origin_v or "negăsit", ctx.solution, origin_s or "negăsit")
            ctx.level3_by_network = self._read_level3_table(grid)
            ctx.pts = self._read_pt_blocks(grid)
        wcfg = self.cfg.wbs
        for name, source in (("project_definition", wcfg.project_definition),
                             ("budget_code", wcfg.budget_code),
                             ("location", wcfg.location)):
            value, origin = self.resolve_field(source)
            ctx.fields[name] = value
            ctx.field_origin[name] = origin
            log.info("Câmp proiect '%s' = '%s' (sursa: %s)", name, value, origin or "negăsit")
        return ctx

    def _read_level3_table(self, grid: SheetGrid) -> Dict[str, str]:
        """Tabelul "Activitate SAP | Nivel 3" (căutat după antet, oriunde în foaie)."""
        net_hdr, n3_hdr = (norm(h) for h in self.cfg.rules.level3_table_headers)
        net_re = re.compile(self.cfg.rules.network_regex)
        for r, c, value in grid.iter_cells():
            if norm(value) != net_hdr:
                continue
            row_vals = grid.rows[r]
            n3_col = next((cc for cc in range(c + 1, min(len(row_vals), c + 6))
                           if norm(row_vals[cc]) == n3_hdr), None)
            if n3_col is None:
                continue
            mapping: Dict[str, str] = {}
            for rr in range(r + 1, len(grid)):
                net = to_text(grid.value(rr, c)).upper()
                if not net_re.match(net):
                    break
                n3 = to_text(grid.value(rr, n3_col)).upper()
                if n3:
                    mapping[net] = n3
            if mapping:
                log.info("Tabel Nivel 3 găsit în %s: %s", grid.title,
                         ", ".join(f"{k}->{v}" for k, v in mapping.items()))
                return mapping
        log.info("Tabelul 'Activitate SAP / Nivel 3' nu a fost găsit; se folosește config.rules.networks")
        return {}

    def _read_pt_blocks(self, grid: SheetGrid) -> Dict[int, PtInfo]:
        """Blocurile PT1, PT2 ... din foaia de date (ex. Cod Post / Cod Linie)."""
        pt_re = re.compile(self.cfg.excel.pt_column_regex, re.IGNORECASE)
        labels = {field: [norm(x) for x in names]
                  for field, names in self.cfg.wbs.pt_assignment_labels.items()}
        result: Dict[int, PtInfo] = {}
        for r, c, value in grid.iter_cells():
            match = pt_re.match(to_text(value))
            if not match:
                continue
            number = int(match.group(1))
            info = PtInfo(number)
            found_label = False
            for rr in range(r, min(len(grid), r + 4)):
                for cc in range(c, c + 12):
                    text = norm(grid.value(rr, cc))
                    for field, wanted in labels.items():
                        if text and text in wanted:
                            found_label = True
                            info.values.setdefault(field, to_text(grid.value(rr, cc + 1)))
            if found_label and number not in result:
                result[number] = info
        return result

    def resolve_field(self, source: FieldSource) -> Tuple[str, str]:
        """Valoarea unui câmp de proiect: nume definit > adresă > etichetă > implicit."""
        wb = self._wb
        for name in source.defined_names:
            value, origin = self._defined_name_value(wb, name)
            if value:
                return value, origin
        if source.cell and "!" in source.cell:
            sheet, coord = source.cell.split("!", 1)
            value = self._cell_value(sheet.strip("'"), coord)
            if value:
                return value, source.cell
        sheets = source.sheets or (wb.sheetnames if wb else [])
        for title in sheets:
            grid = self._grid(title)
            if grid is None or not source.labels:
                continue
            value, origin = find_label_value(grid, source.labels)
            if value:
                return value, origin
        if source.default:
            return str(source.default), "config (implicit)"
        return "", ""

    def _cell_value(self, sheet: str, coord: str) -> str:
        from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

        grid = self._grid(sheet)
        if grid is None:
            return ""
        col_letter, row = coordinate_from_string(coord.replace("$", "").split(":")[0])
        return to_text(grid.value(row - 1, column_index_from_string(col_letter) - 1))

    def _defined_name_value(self, wb, name: str) -> Tuple[str, str]:
        if wb is None:
            return "", ""
        candidates = []
        names = getattr(wb, "defined_names", None)
        try:
            if names is not None and name in names:
                candidates.append(names[name])
        except TypeError:  # versiuni vechi openpyxl
            dn = names.get(name) if hasattr(names, "get") else None
            if dn is not None:
                candidates.append(dn)
        for ws in wb.worksheets:
            local = getattr(ws, "defined_names", None)
            if local and name in local:
                candidates.append(local[name])
        for dn in candidates:
            try:
                for sheet, coord in dn.destinations:
                    value = self._cell_value(sheet, coord)
                    if value:
                        return value, f"nume definit {name} ({sheet}!{coord})"
            except Exception:  # noqa: BLE001 - nume definite invalide sunt ignorate
                continue
        return "", ""
