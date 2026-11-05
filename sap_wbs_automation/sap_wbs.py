#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAP WBS + NORME - automatizare (un singur fișier)
=================================================

Citește devizul general (.xlsm), deduce structura WBS (nivel 3 și 4), activitățile
din rețelele standard și normele (serviciile) cu cantitățile lor, apoi le introduce
în SAP prin SAP GUI Scripting (win32com.client), în tranzacția de proiect (CJ02).

Utilizare (Windows, SAP Logon deschis și autentificat):
    python sap_wbs.py deviz.xlsm --dry-run                 # doar citire + raport (fără SAP)
    python sap_wbs.py deviz.xlsm --max-n4 1 --pas-cu-pas   # primul test în SAP
    python sap_wbs.py deviz.xlsm                           # rulare completă
    python sap_wbs.py -h                                   # toate opțiunile

Dependențe: pip install openpyxl pandas pywin32   (PyYAML doar pentru config.yaml opțional)

Configurarea: valorile implicite sunt în secțiunea CONFIGURARE de mai jos (clasele
*Config). Opțional, un fișier config.yaml lângă script (sau --config) le suprascrie.
ID-urile SAP GUI sunt toate în clasa SapIds.

Cuprins:
    01. UTILITARE TEXT ȘI NUMERE
    02. CONFIGURARE (valorile implicite; pot fi suprascrise din config.yaml)
    03. STRUCTURI DE DATE
    04. CITIREA DEVIZULUI (Rule-Based Engine pentru antete)
    05. REGULI DE DEDUCȚIE (rețea, activitate, sufix, denumiri WBS)
    06. PLANUL IERARHIC (Nivel 3 -> Nivel 4 -> Activitate -> Servicii)
    07. RAPOARTE EXCEL
    08. ID-URI SAP GUI (din înregistrarea Script1.vbs)
    09. CONEXIUNEA SAP, BARA DE STARE, POP-UP-URI, TABELE
    10. ACȚIUNI ÎN TRANZACȚIA DE PROIECT (CJ02)
    11. EXECUȚIA PLANULUI ÎN SAP + TRATAREA ERORILOR
    12. RECALCULARE FORMULE PRIN EXCEL (opțional)
    13. LOGGING (fișier + consolă)
    14. LINIE DE COMANDĂ
"""
from __future__ import annotations

import argparse
import dataclasses
import logging
import math
import re
import sys
import tempfile
import time
import unicodedata
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple

__version__ = "1.0.0"

log = logging.getLogger("sapwbs")


# =============================================================================
#  01. UTILITARE TEXT ȘI NUMERE
# =============================================================================
#
#  Funcții utilitare pentru text și numere (fără dependențe externe).
# =============================================================================

_TOKEN_RE = re.compile(r"[a-z0-9]+")


def strip_diacritics(text: str) -> str:
    """Elimină diacriticele (ă, â, î, ș, ş, ț, ţ ...) păstrând litera de bază."""
    decomposed = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def norm(value: Any) -> str:
    """Normalizează un text pentru comparații robuste.

    "Cod normă" -> "cod norma", "Cant." -> "cant", "T.I." -> "t i".
    """
    if value is None:
        return ""
    text = strip_diacritics(str(value).lower())
    return " ".join(_TOKEN_RE.findall(text))


def first_line(value: Any) -> str:
    """Prima linie (ne-goală) dintr-o celulă cu text pe mai multe rânduri."""
    if value is None:
        return ""
    for line in str(value).splitlines():
        if line.strip():
            return line.strip()
    return ""


def clean_spaces(value: Any) -> str:
    """Colapsează spațiile multiple și elimină spațiile de la capete."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def parse_number(value: Any) -> Optional[float]:
    """Convertește o valoare din Excel în float; întoarce None dacă nu e număr.

    Acceptă numere native și texte de tip "1.234,56", "1,5", "2.75", "-".
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    text = str(value).strip().replace(" ", "").replace(" ", "")
    if not text or not re.search(r"\d", text):
        return None
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def format_quantity(value: float, decimals: int = 3, decimal_separator: str = ",") -> str:
    """Formatează o cantitate pentru câmpurile SAP (ex. 348.6 -> "348,6", 40.0 -> "40").

    Zecimalele inutile sunt eliminate; nu se folosește separator de mii
    (SAP acceptă valoarea fără el, indiferent de setările utilizatorului).
    """
    rounded = round(float(value), decimals)
    if rounded == 0:
        rounded = 0.0  # evită "-0"
    text = f"{rounded:.{decimals}f}"
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text.replace(".", decimal_separator)


def truncate(text: str, max_len: int) -> str:
    """Taie textul la lungimea maximă acceptată de câmpul SAP."""
    return text if len(text) <= max_len else text[:max_len].rstrip()


# =============================================================================
#  02. CONFIGURARE (valorile implicite; pot fi suprascrise din config.yaml)
# =============================================================================
#
#  Configurația aplicației.
#
#  Valorile implicite sunt definite aici; fișierul ``config.yaml`` le suprascrie.
#  Orice cheie necunoscută din YAML produce o eroare, ca să nu treacă neobservate
#  greșelile de scriere.
# =============================================================================

class ConfigError(Exception):
    """Configurație invalidă."""


# --------------------------------------------------------------------------- #
#  Secțiuni
# --------------------------------------------------------------------------- #
def _default_columns() -> Dict[str, List[str]]:
    return {
        "code": ["Cod Norma", "Cod normă", "Cod norma", "Cod articol", "Cod serviciu", "E"],
        "quantity": [
            "Cant.", "Cant", "Cantitate", "Cantitate totala", "Cant. Totala",
            "Cantitate totală", "Norma", "Norme_ore", "Norme ore",
        ],
        "network": ["Activitate SAP", "Retea standard", "Retea SAP", "Rețea standard"],
        "activity": ["Denumire activitate SAP", "Activitate SAP denumire", "Denumire activitate"],
        "part": ["Partea", "Parte"],
        "description": ["Descriere", "Descriere normă", "Descriere norma", "Denumire"],
        "install_type": ["T.I.", "Instalatia", "Instalația", "Tip instalatie"],
        "override": ["Activitate SAP ales", "Activitate SAP aleasa", "Activitate aleasa"],
    }


@dataclass
class ExcelConfig:
    header_scan_rows: int = 40
    max_scan_col: int = 120
    skip_hidden_sheets: bool = True
    include_sheets: List[str] = field(default_factory=list)  # regex; gol = toate
    exclude_sheets: List[str] = field(default_factory=lambda: ["^BackupFormule$", "^Anexa", "vechi"])
    stop_after_empty_rows: int = 60
    columns: Dict[str, List[str]] = field(default_factory=_default_columns)
    required_columns: List[str] = field(default_factory=lambda: ["code", "quantity", "network"])
    pt_column_regex: str = r"^PT\s*(\d+)$"
    code_regex: str = r"^[A-Z]\d{5}$"


@dataclass
class NetworkInfo:
    level3: str = ""
    voltage: Optional[str] = None   # MT / JT / None (oricare)
    solution: Optional[str] = None  # LEA / LES / None (oricare)
    desc: str = ""


def _default_networks() -> Dict[str, NetworkInfo]:
    return {
        "A412": NetworkInfo(level3="M1", desc="APARATAJ"),
        "A414": NetworkInfo(level3="M1", desc="TRANSFORMATOR"),
        "A441": NetworkInfo(level3="M1", desc="CLADIRE"),
        "A451": NetworkInfo(level3="M1", voltage="MT", solution="LEA", desc="LEA MT COND NEIZOLAT"),
        "A452": NetworkInfo(level3="M1", voltage="MT", solution="LEA", desc="LEA MT COND IZOLAT"),
        "A453": NetworkInfo(level3="M1", voltage="MT", solution="LES", desc="RETEA LES MT"),
        "A462": NetworkInfo(level3="B1", voltage="JT", solution="LEA", desc="RETEA LEA JT"),
        "A463": NetworkInfo(level3="B1", voltage="JT", solution="LES", desc="RETEA LES JT"),
        "A471": NetworkInfo(level3="B1", voltage="JT", desc="BRANSAMENTE JT"),
        "A61F": NetworkInfo(level3="T1", desc="TLC"),
    }


def _default_keyword_rules() -> List[Dict[str, str]]:
    return [
        {"description_regex": r"^(demont|demol|dezafect|desfiint)", "prefer_regex": r"\bdemont|\bdem\b"},
        {"description_regex": r"\binloc", "prefer_regex": r"\binloc"},
        {"description_regex": r".*", "prefer_regex": r"\bmont|\bfurniz|\bpozare|\binstal"},
    ]


@dataclass
class RulesConfig:
    allowed_parts: List[str] = field(default_factory=lambda: ["L", "M", "S", "E"])
    part_order: List[str] = field(default_factory=lambda: ["L", "M", "S", "E"])
    activity_suffix_strip: List[str] = field(default_factory=lambda: ["NUM"])
    activity_key_regex: str = r"^([A-Z0-9]{3})"
    activity_split_regex: str = r"\s+/\s+(?=[A-Z0-9]{3})"
    network_regex: str = r"^A[0-9A-Z]{3}$"
    network_split_regex: str = r"\s*/\s*"
    activity_name_max_len: int = 40
    ambiguity_fallback: str = "first"  # first | skip
    networks: Dict[str, NetworkInfo] = field(default_factory=_default_networks)
    activity_keyword_rules: List[Dict[str, str]] = field(default_factory=_default_keyword_rules)
    context_sheet: str = "Date"
    voltage_labels: List[str] = field(default_factory=lambda: ["Nivel Tensiune"])
    solution_labels: List[str] = field(default_factory=lambda: ["Solutie proiectata"])
    level3_table_headers: List[str] = field(default_factory=lambda: ["Activitate SAP", "Nivel 3"])


@dataclass
class FieldSource:
    """De unde se citește o valoare de proiect (nume definit > etichetă > adresă > implicit)."""

    defined_names: List[str] = field(default_factory=list)
    labels: List[str] = field(default_factory=list)
    sheets: List[str] = field(default_factory=list)  # gol = toate foile
    cell: str = ""  # adresă explicită, ex. "Date!C5"
    default: str = ""


@dataclass
class WbsConfig:
    level3_code: str = "+{n3}"
    level3_desc: Dict[str, str] = field(
        default_factory=lambda: {"M1": "MEDIE TENSIUNE", "B1": "JOASA TENSIUNE", "T1": "TELECONTROL"}
    )
    level4_code: str = "+{net}R{idx}"
    level4_desc_pt: str = "{base} PT{pt}"
    level4_desc_common: str = "{base}"
    level4_desc_common_with_pt: str = "{base} COMUN"
    desc_max_len: int = 40
    project_definition: FieldSource = field(
        default_factory=lambda: FieldSource(
            defined_names=["COD_PROIECT"],
            labels=["Cod proiect SAP", "Definitie proiect", "Proiect SAP", "Cod proiect"],
            sheets=["Date"],
        )
    )
    budget_code: FieldSource = field(
        default_factory=lambda: FieldSource(
            defined_names=["COD_BUGET"],
            labels=["Cod buget", "Cod bugetar", "Cod buget SAP"],
            sheets=["Date"],
        )
    )
    location: FieldSource = field(
        default_factory=lambda: FieldSource(labels=["Locatie SAP", "Locatie"], sheets=["Date"])
    )
    assignment_field_by_network: Dict[str, str] = field(
        default_factory=lambda: {"A412": "EQUNR", "A414": "EQUNR", "A441": "EQUNR"}
    )
    assignment_default_field: str = "TPLNR"
    pt_assignment_labels: Dict[str, List[str]] = field(
        default_factory=lambda: {"EQUNR": ["Cod Post"], "TPLNR": ["Cod Linie"]}
    )
    default_tplnr: str = ""
    default_equnr: str = ""
    level3_tplnr: Dict[str, str] = field(default_factory=dict)
    budget_on_levels: List[int] = field(default_factory=lambda: [4])
    custom_fields_by_network: Dict[str, Dict[str, str]] = field(
        default_factory=lambda: {"A441": {"ZZPARTE": "1"}}
    )


@dataclass
class SapConfig:
    transaction: str = "CJ02"
    system_name: str = ""
    client: str = ""
    user: str = ""
    std_network_alternative: str = "QS"
    network_popup_values: List[str] = field(default_factory=lambda: ["AMC", "MBT1"])
    purchasing_group: str = "AMC"   # AFVGD-EKGRP
    requisitioner: str = "AMC"      # AFVGD-AFNAM
    set_unlimited_overfulfillment: bool = True  # ESLL-UEBTK
    price_one_regex: str = r"^E00\d{3}$"        # servicii valorice -> preț brut 1
    decimal_separator: str = ","
    quantity_decimals: int = 3
    busy_timeout_s: float = 120.0
    max_popups: int = 10
    auto_confirm_info_popups: bool = True


@dataclass
class OutputConfig:
    log_dir: str = "logs"
    report_dir: str = "rapoarte"
    log_level: str = "INFO"


@dataclass
class AppConfig:
    excel: ExcelConfig = field(default_factory=ExcelConfig)
    rules: RulesConfig = field(default_factory=RulesConfig)
    wbs: WbsConfig = field(default_factory=WbsConfig)
    sap: SapConfig = field(default_factory=SapConfig)
    output: OutputConfig = field(default_factory=OutputConfig)
    base_dir: Path = field(default_factory=Path.cwd)

    def resolve_path(self, value: str) -> Path:
        path = Path(value)
        return path if path.is_absolute() else self.base_dir / path


# --------------------------------------------------------------------------- #
#  Încărcare YAML
# --------------------------------------------------------------------------- #
def _apply(obj: Any, data: Dict[str, Any], path: str) -> Any:
    """Aplică valorile din YAML peste un dataclass.

    Dicționarele se îmbină (cheile din YAML le suprascriu pe cele implicite),
    listele și valorile simple se înlocuiesc.
    """
    if not isinstance(data, dict):
        raise ConfigError(f"'{path}' trebuie să fie un dicționar")
    known = {f.name for f in dataclasses.fields(obj)}
    for key, value in data.items():
        if key not in known or key == "base_dir":
            raise ConfigError(f"Cheie necunoscută în configurație: '{path}.{key}'")
        current = getattr(obj, key)
        if dataclasses.is_dataclass(current):
            _apply(current, value or {}, f"{path}.{key}")
        elif isinstance(obj, RulesConfig) and key == "networks":
            if not isinstance(value, dict):
                raise ConfigError(f"'{path}.networks' trebuie să fie un dicționar")
            for net, spec in value.items():
                net_code = str(net).upper()
                base = current.get(net_code, NetworkInfo())
                current[net_code] = _apply(base, spec or {}, f"{path}.networks.{net}")
        elif isinstance(current, dict):
            if not isinstance(value, dict):
                raise ConfigError(f"'{path}.{key}' trebuie să fie un dicționar")
            current.update(value)
        else:
            setattr(obj, key, value)
    return obj


def load_config(path: Optional[Path]) -> AppConfig:
    """Încarcă configurația; fără fișier se folosesc valorile implicite."""
    cfg = AppConfig()
    if path is None:
        return cfg
    path = Path(path)
    if not path.exists():
        raise ConfigError(f"Fișierul de configurare nu există: {path}")
    try:
        import yaml  # import local: PyYAML e necesar doar când există fișier
    except ImportError as exc:  # pragma: no cover
        raise ConfigError("Lipsește pachetul PyYAML (pip install pyyaml)") from exc
    with path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}
    _apply(cfg, data, "config")
    cfg.base_dir = path.resolve().parent
    if cfg.rules.ambiguity_fallback not in ("first", "skip"):
        raise ConfigError("rules.ambiguity_fallback trebuie să fie 'first' sau 'skip'")
    return cfg


# =============================================================================
#  03. STRUCTURI DE DATE
# =============================================================================
#
#  Structurile de date comune: rânduri sursă, probleme, planul ierarhic.
# =============================================================================

INFO = "INFO"
WARNING = "WARNING"
ERROR = "ERROR"


@dataclass
class Issue:
    """O problemă găsită la citire/planificare (apare în raportul pre-flight)."""

    severity: str
    message: str
    sheet: str = ""
    row: Optional[int] = None
    code: str = ""

    def location(self) -> str:
        if self.sheet and self.row:
            return f"{self.sheet}!{self.row}"
        return self.sheet


# --------------------------------------------------------------------------- #
#  Excel
# --------------------------------------------------------------------------- #
@dataclass
class HeaderMap:
    """Rezultatul detecției antetului într-o foaie."""

    sheet: str
    header_row: int                      # 1-based (ca în Excel)
    columns: Dict[str, int]              # coloană logică -> index 0-based
    pt_columns: Dict[int, int]           # număr PT -> index 0-based
    score: float = 0.0

    def letter(self, logical: str) -> str:
        from openpyxl.utils import get_column_letter

        idx = self.columns.get(logical)
        return get_column_letter(idx + 1) if idx is not None else "-"


@dataclass
class SourceRow:
    """Un rând de normă/material cu cantitate nenulă, citit din deviz."""

    sheet: str
    row: int
    code: str
    quantity: float
    pt_quantities: Dict[int, float]
    network_raw: str
    activity_raw: str
    part: str = ""
    description: str = ""
    install_type: str = ""
    override: str = ""

    @property
    def ref(self) -> str:
        return f"{self.sheet}!{self.row}"

    @property
    def common_quantity(self) -> float:
        """Cantitatea nealocată pe PT-uri (coloana totală minus suma PT)."""
        return self.quantity - sum(self.pt_quantities.values())


@dataclass
class PtInfo:
    number: int
    values: Dict[str, str] = field(default_factory=dict)  # ex. {"EQUNR": "...", "TPLNR": "..."}


@dataclass
class WorkbookContext:
    """Informații de proiect citite din foaia de date (context pentru reguli)."""

    voltage: str = ""
    solution: str = ""
    level3_by_network: Dict[str, str] = field(default_factory=dict)
    pts: Dict[int, PtInfo] = field(default_factory=dict)
    fields: Dict[str, str] = field(default_factory=dict)        # project_definition, budget_code, location
    field_origin: Dict[str, str] = field(default_factory=dict)  # de unde a fost citită valoarea


@dataclass
class ReadResult:
    rows: List[SourceRow]
    headers: List[HeaderMap]
    context: WorkbookContext
    issues: List[Issue]


# --------------------------------------------------------------------------- #
#  Plan
# --------------------------------------------------------------------------- #
@dataclass
class ServiceLine:
    code: str
    quantity: float
    description: str = ""
    sources: List[Tuple[str, int]] = field(default_factory=list)

    @property
    def source_refs(self) -> str:
        return ", ".join(f"{sheet}!{row}" for sheet, row in self.sources)


@dataclass
class ActivityPlan:
    key: str            # ex. "XL4"
    part: str           # L / M / S / E
    base_name: str      # ex. "XL4 Pozare LES MT KM"
    text: str           # ex. "XL4 Pozare LES MT KM (L)" (textul final în SAP)
    services: Dict[str, ServiceLine] = field(default_factory=dict)

    def add(self, code: str, quantity: float, description: str, source: Tuple[str, int]) -> None:
        line = self.services.get(code)
        if line is None:
            line = self.services[code] = ServiceLine(code=code, quantity=0.0, description=description)
        line.quantity += quantity
        line.sources.append(source)


@dataclass
class Wbs4Plan:
    code: str                    # ex. "+453R1"
    description: str
    network: str                 # ex. "A453"
    level3: str                  # ex. "M1"
    pt: Optional[int]            # None = lucrări comune (nealocate pe PT)
    assignments: Dict[str, str] = field(default_factory=dict)   # STORT/TPLNR/EQUNR
    custom_fields: Dict[str, str] = field(default_factory=dict)  # ZZCOD_BUGET, ZZPARTE ...
    activities: Dict[str, ActivityPlan] = field(default_factory=dict)

    def service_count(self) -> int:
        return sum(len(a.services) for a in self.activities.values())


@dataclass
class Wbs3Plan:
    n3: str                      # ex. "M1"
    code: str                    # ex. "+M1"
    description: str
    assignments: Dict[str, str] = field(default_factory=dict)
    custom_fields: Dict[str, str] = field(default_factory=dict)
    children: Dict[str, Wbs4Plan] = field(default_factory=dict)


@dataclass
class ProjectPlan:
    project_definition: str
    budget_code: str
    level3: Dict[str, Wbs3Plan]
    issues: List[Issue]
    stats: Dict[str, float] = field(default_factory=dict)

    def iter_wbs4(self):
        for n3 in self.level3.values():
            for n4 in n3.children.values():
                yield n3, n4

    def has_errors(self) -> bool:
        return any(i.severity == ERROR for i in self.issues)


# =============================================================================
#  04. CITIREA DEVIZULUI (Rule-Based Engine pentru antete)
# =============================================================================
#
#  Citirea devizului: detecția dinamică a antetelor și extragerea normelor.
#
#  Nicio poziție de coloană sau rând nu este fixă. Pentru fiecare foaie:
#
#  1. se scanează primele ``header_scan_rows`` rânduri;
#  2. fiecare celulă este comparată (normalizat, fără diacritice) cu alias-urile
#     din configurație; potrivirea exactă valorează mai mult decât cea parțială;
#  3. rândul care conține toate coloanele obligatorii cu scorul maxim devine antet;
#  4. coloana de cod este validată și după conținut (regex pe valori), cu
#     fallback pe coloana care conține cele mai multe coduri valide.
#
#  Foile care nu au coloanele obligatorii sunt ignorate (și apar în log).
# =============================================================================

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
                    for field_name, wanted in labels.items():
                        if text and text in wanted:
                            found_label = True
                            info.values.setdefault(field_name, to_text(grid.value(rr, cc + 1)))
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


# =============================================================================
#  05. REGULI DE DEDUCȚIE (rețea, activitate, sufix, denumiri WBS)
# =============================================================================
#
#  Regulile de deducție: rețea standard, activitate, sufix, denumiri.
#
#  Legătura normă -> activitate vine din coloanele "Activitate SAP" (rețeaua
#  standard, ex. A453) și "Denumire activitate SAP" (ex. "XL4 Pozare LES MT KM").
#  Când există mai multe variante ("A451/A452", "YL1 ... / YM1 ...") se aplică,
#  în ordine:
#
#  1. coloana opțională de forțare ("Activitate SAP ales");
#  2. perechi poziționale (rețeaua i <-> activitatea i, dacă listele au aceeași lungime);
#  3. contextul din foaia Date (nivel tensiune MT/JT, soluție LEA/LES);
#  4. rețelele deja folosite de rânduri neambigue din același deviz;
#  5. reguli de cuvinte cheie pe descrierea normei (demontare / înlocuire / montare);
#  6. politica de rezervă din config (prima variantă cu avertizare, sau sări rândul).
# =============================================================================

@dataclass
class Resolution:
    network: Optional[str] = None
    activity_base: Optional[str] = None
    key: Optional[str] = None
    part: str = ""
    skip_reason: str = ""
    warnings: List[str] = field(default_factory=list)
    ambiguous_network: bool = False

    @property
    def ok(self) -> bool:
        return not self.skip_reason


# --------------------------------------------------------------------------- #
#  Funcții pure (ușor de testat)
# --------------------------------------------------------------------------- #
def split_networks(raw: str, cfg: AppConfig) -> List[str]:
    net_re = re.compile(cfg.rules.network_regex)
    parts = [p.strip().upper() for p in re.split(cfg.rules.network_split_regex, raw or "") if p.strip()]
    return [p for p in parts if net_re.match(p)]


def split_activities(raw: str, cfg: AppConfig) -> List[str]:
    key_re = re.compile(cfg.rules.activity_key_regex)
    parts = [clean_spaces(p) for p in re.split(cfg.rules.activity_split_regex, raw or "") if p.strip()]
    return [p for p in parts if key_re.match(p)]


def clean_activity_name(raw: str, strip_suffixes: Sequence[str]) -> str:
    """Elimină unitatea de măsură de la final (ex. "... NUM")."""
    text = clean_spaces(raw)
    for suffix in strip_suffixes:
        text = re.sub(rf"\s+{re.escape(suffix)}$", "", text)
    return text


def activity_key(name: str, key_regex: str) -> Optional[str]:
    match = re.match(key_regex, (name or "").strip().upper())
    return match.group(1) if match else None


def build_activity_text(base: str, part: str, max_len: int) -> str:
    """Textul final al activității în SAP, ex. "XL4 Pozare LES MT KM (L)"."""
    suffix = f" ({part})"
    return truncate(base, max_len - len(suffix)) + suffix


def part_of(row: SourceRow) -> str:
    """Partea (L/M/S/E): coloana "Partea" sau, lipsind, prima literă a codului."""
    return (row.part or row.code[:1]).upper()


def network_matches_context(info: Optional[NetworkInfo], ctx: WorkbookContext) -> bool:
    if info is None:
        return True
    voltage = norm(ctx.voltage).upper().replace(" ", "+")
    solution = norm(ctx.solution).upper().replace(" ", "+")
    if info.voltage and voltage and info.voltage.upper() not in voltage.split("+"):
        return False
    if info.solution and solution and info.solution.upper() not in solution.split("+"):
        return False
    return True


def choose_by_keywords(description: str, options: Sequence[str], rules: Iterable[dict]) -> Optional[str]:
    desc = norm(description)
    for rule in rules:
        if not re.search(rule.get("description_regex", ".*"), desc):
            continue
        prefer = re.compile(rule.get("prefer_regex", ""), re.IGNORECASE)
        for opt in options:
            if prefer.search(norm(opt) + " "):
                return opt
    return None


# --------------------------------------------------------------------------- #
#  Motorul de reguli
# --------------------------------------------------------------------------- #
class RuleEngine:
    def __init__(self, cfg: AppConfig, context: WorkbookContext):
        self.cfg = cfg
        self.ctx = context
        self.used_networks: Set[str] = set()

    def learn_used_networks(self, rows: Iterable[SourceRow]) -> None:
        """Rețelele care apar neambiguu în deviz (folosite la dezambiguizare)."""
        for row in rows:
            nets = split_networks(row.network_raw, self.cfg)
            if len(nets) == 1 and part_of(row) in self.cfg.rules.allowed_parts:
                self.used_networks.add(nets[0])

    def resolve(self, row: SourceRow) -> Resolution:
        rcfg = self.cfg.rules
        res = Resolution(part=part_of(row))
        if res.part not in rcfg.allowed_parts:
            res.skip_reason = f"partea '{res.part}' nu este procesată"
            return res

        nets = split_networks(row.network_raw, self.cfg)
        acts = split_activities(row.activity_raw, self.cfg)
        override_net, override_key = self._parse_override(row.override)

        if override_net:
            if nets and override_net not in nets:
                res.warnings.append(f"rețeaua forțată {override_net} nu e printre variantele {nets}")
            nets_for_pairing = nets
            chosen_net = override_net
        else:
            nets_for_pairing = nets
            chosen_net = None

        if not nets and not chosen_net:
            res.skip_reason = (f"fără rețea standard validă în 'Activitate SAP' "
                               f"('{row.network_raw or 'gol'}')")
            return res
        if not acts:
            res.skip_reason = (f"fără activitate validă în 'Denumire activitate SAP' "
                               f"('{row.activity_raw or 'gol'}')")
            return res

        if chosen_net is None:
            chosen_net = self._choose_network(nets, res)
            if chosen_net is None:
                return res

        res.network = chosen_net
        chosen_act = self._choose_activity(row, chosen_net, nets_for_pairing, acts, override_key, res)
        if chosen_act is None:
            return res

        base = clean_activity_name(chosen_act, rcfg.activity_suffix_strip)
        res.activity_base = base
        res.key = activity_key(base, rcfg.activity_key_regex)
        if not res.key:
            res.skip_reason = f"nu pot extrage codul activității din '{base}'"
        return res

    # -- detalii ----------------------------------------------------------- #
    def _parse_override(self, raw: str):
        if not raw:
            return None, None
        tokens = re.split(r"[\s;|,/]+", raw.strip().upper())
        net_re = re.compile(self.cfg.rules.network_regex)
        net = next((t for t in tokens if net_re.match(t)), None)
        key = next((t for t in tokens if t != net and re.match(r"^[A-Z0-9]{3}$", t)), None)
        return net, key

    def _choose_network(self, nets: List[str], res: Resolution) -> Optional[str]:
        if len(nets) == 1:
            return nets[0]
        rcfg = self.cfg.rules
        candidates = list(nets)
        by_ctx = [n for n in candidates if network_matches_context(rcfg.networks.get(n), self.ctx)]
        if by_ctx:
            candidates = by_ctx
        if len(candidates) > 1:
            by_use = [n for n in candidates if n in self.used_networks]
            if by_use:
                candidates = by_use
        if len(candidates) == 1:
            res.warnings.append(f"rețea aleasă {candidates[0]} din {'/'.join(nets)} (context/utilizare)")
            return candidates[0]
        res.ambiguous_network = True
        if rcfg.ambiguity_fallback == "skip":
            res.skip_reason = (f"rețea ambiguă {'/'.join(nets)}; completați coloana "
                               f"'Activitate SAP ales'")
            return None
        res.warnings.append(f"rețea ambiguă {'/'.join(nets)} -> aleasă prima variantă "
                            f"compatibilă: {candidates[0]}")
        return candidates[0]

    def _choose_activity(self, row: SourceRow, chosen_net: str, nets: List[str], acts: List[str],
                         override_key: Optional[str], res: Resolution) -> Optional[str]:
        rcfg = self.cfg.rules
        if override_key:
            for act in acts:
                if activity_key(act, rcfg.activity_key_regex) == override_key:
                    return act
            res.warnings.append(f"activitatea forțată {override_key} nu e printre variante")
        if len(acts) == 1:
            return acts[0]
        if len(nets) == len(acts) and chosen_net in nets:
            return acts[nets.index(chosen_net)]
        picked = choose_by_keywords(row.description, acts, rcfg.activity_keyword_rules)
        if picked:
            if picked != acts[0]:
                res.warnings.append(f"activitate aleasă după descriere: '{picked}'")
            else:
                res.warnings.append(f"activitate cu variante; aleasă '{picked}' după descriere")
            return picked
        if rcfg.ambiguity_fallback == "skip":
            res.skip_reason = f"activitate ambiguă: {' / '.join(acts)}"
            return None
        res.warnings.append(f"activitate ambiguă -> aleasă prima variantă: '{acts[0]}'")
        return acts[0]


# --------------------------------------------------------------------------- #
#  Denumiri WBS
# --------------------------------------------------------------------------- #
def network_short(network: str) -> str:
    """A453 -> 453, A61F -> 61F (partea folosită în codul WBS de nivel 4)."""
    return network[1:] if network[:1].upper() == "A" else network


def level4_code(cfg: AppConfig, network: str, index: int) -> str:
    return cfg.wbs.level4_code.format(net=network_short(network), idx=index, network=network)


def level4_description(cfg: AppConfig, network: str, pt: Optional[int], network_has_pt: bool) -> str:
    info = cfg.rules.networks.get(network)
    base = (info.desc if info and info.desc else network)
    if pt is not None:
        text = cfg.wbs.level4_desc_pt.format(base=base, pt=pt, network=network)
    elif network_has_pt:
        text = cfg.wbs.level4_desc_common_with_pt.format(base=base, network=network)
    else:
        text = cfg.wbs.level4_desc_common.format(base=base, network=network)
    return truncate(text, cfg.wbs.desc_max_len)


def level3_code(cfg: AppConfig, n3: str) -> str:
    return cfg.wbs.level3_code.format(n3=n3)


def level3_description(cfg: AppConfig, n3: str) -> str:
    return truncate(cfg.wbs.level3_desc.get(n3, n3), cfg.wbs.desc_max_len)


# =============================================================================
#  06. PLANUL IERARHIC (Nivel 3 -> Nivel 4 -> Activitate -> Servicii)
# =============================================================================
#
#  Construirea planului ierarhic: Nivel 3 -> Nivel 4 -> Activitate -> Servicii.
#
#  Reguli de alocare pe nivelul 4 (stabilite de noi):
#
#  * fiecare rețea standard (ex. A453) are propriile elemente de nivel 4;
#  * cantitatea din coloana PTk merge pe nivelul 4 al PT-ului k: ``+{rețea}R{k}``
#    (ex. ``+441R2`` = clădirea PT2);
#  * restul (cantitatea totală minus suma PT) = lucrări comune. Dacă rețeaua nu are
#    nimic alocat pe PT, lucrările comune primesc indicele 1 (``+452R1``, ca în
#    înregistrarea SAP); altfel primesc primul indice liber după ultimul PT din deviz
#    (cu 3 PT-uri: ``+453R4``).
# =============================================================================

EPS = 1e-9

Allocation = Tuple[SourceRow, Resolution, str, Optional[int], float]


class Planner:
    def __init__(self, cfg: AppConfig):
        self.cfg = cfg

    def build(self, read: ReadResult, project_override: str = "", budget_override: str = "") -> ProjectPlan:
        cfg = self.cfg
        ctx = read.context
        issues: List[Issue] = list(read.issues)

        def add_issue(severity: str, message: str, row: Optional[SourceRow] = None) -> None:
            issue = Issue(severity, message, row.sheet if row else "", row.row if row else None,
                          row.code if row else "")
            issues.append(issue)
            level = {INFO: logging.INFO, WARNING: logging.WARNING, ERROR: logging.ERROR}[severity]
            log.log(level, "%s%s", f"[{issue.location()} {issue.code}] " if row else "", message)

        engine = RuleEngine(cfg, ctx)
        engine.learn_used_networks(read.rows)

        allocations: List[Allocation] = []
        ignored_parts: Counter = Counter()
        qty_in = 0.0
        for row in read.rows:
            res = engine.resolve(row)
            if not res.ok:
                if res.part not in cfg.rules.allowed_parts:
                    ignored_parts[res.part or "?"] += 1
                else:
                    add_issue(WARNING, f"Rând sărit: {res.skip_reason}", row)
                continue
            for warning in res.warnings:  # alegeri euristice -> de verificat de utilizator
                add_issue(WARNING, warning, row)
            net_info = cfg.rules.networks.get(res.network)
            level3 = ctx.level3_by_network.get(res.network) or (net_info.level3 if net_info else "")
            if not level3:
                add_issue(WARNING, f"Rând sărit: nu știu nivelul 3 pentru rețeaua {res.network} "
                                   "(adăugați-o în config.rules.networks)", row)
                continue
            qty_in += row.quantity
            for pt, qty in sorted(row.pt_quantities.items()):
                allocations.append((row, res, level3, pt, qty))
            common = row.common_quantity
            if common < -EPS:
                add_issue(WARNING, f"Suma pe PT-uri depășește cantitatea totală cu {-common:g}; "
                                   "se folosesc doar valorile PT", row)
            elif common > EPS:
                allocations.append((row, res, level3, None, common))

        for part, count in sorted(ignored_parts.items()):
            add_issue(INFO, f"{count} rânduri cu partea '{part}' ignorate (conform cerinței)")

        level3_nodes = self._build_tree(allocations, ctx)

        project = project_override or ctx.fields.get("project_definition", "")
        budget = budget_override or ctx.fields.get("budget_code", "")
        if not project:
            add_issue(ERROR, "Lipsește definiția proiectului SAP: completați celula de lângă eticheta "
                             "'Cod proiect SAP' (foaia Date), numele definit COD_PROIECT sau --proiect")
        if not budget:
            add_issue(ERROR, "Lipsește codul de buget: completați celula de lângă eticheta 'Cod buget' "
                             "(foaia Date), numele definit COD_BUGET sau --cod-buget")
        self._apply_project_fields(level3_nodes, budget, ctx)

        qty_out = sum(line.quantity for n3 in level3_nodes.values() for n4 in n3.children.values()
                      for act in n4.activities.values() for line in act.services.values())
        if abs(qty_in - qty_out) > 1e-6 * max(1.0, abs(qty_in)):
            add_issue(WARNING, f"Control sumă: cantitate citită {qty_in:g} != cantitate planificată {qty_out:g}")
        else:
            log.info("Control sumă OK: %g unități citite = %g planificate", qty_in, qty_out)

        stats = {
            "randuri_citite": len(read.rows),
            "randuri_planificate": len({(a[0].sheet, a[0].row) for a in allocations}),
            "nivel3": len(level3_nodes),
            "nivel4": sum(len(n.children) for n in level3_nodes.values()),
            "activitati": sum(len(c.activities) for n in level3_nodes.values() for c in n.children.values()),
            "servicii": sum(c.service_count() for n in level3_nodes.values() for c in n.children.values()),
            "cantitate_citita": qty_in,
            "cantitate_planificata": qty_out,
        }
        return ProjectPlan(project_definition=project, budget_code=budget, level3=level3_nodes,
                           issues=issues, stats=stats)

    # ------------------------------------------------------------------ #
    def _build_tree(self, allocations: List[Allocation], ctx) -> Dict[str, Wbs3Plan]:
        cfg = self.cfg
        networks_with_pt = {res.network for _, res, _, pt, _ in allocations if pt is not None}
        max_pt = max((pt for _, _, _, pt, _ in allocations if pt is not None), default=0)
        common_index_with_pt = max_pt + 1

        nodes: Dict[str, Wbs3Plan] = {}
        for row, res, n3, pt, qty in allocations:
            node3 = nodes.get(n3)
            if node3 is None:
                node3 = nodes[n3] = Wbs3Plan(n3=n3, code=level3_code(cfg, n3),
                                             description=level3_description(cfg, n3))
            has_pt = res.network in networks_with_pt
            index = pt if pt is not None else (common_index_with_pt if has_pt else 1)
            code4 = level4_code(cfg, res.network, index)
            node4 = node3.children.get(code4)
            if node4 is None:
                node4 = node3.children[code4] = Wbs4Plan(
                    code=code4,
                    description=level4_description(cfg, res.network, pt, has_pt),
                    network=res.network,
                    level3=n3,
                    pt=pt,
                )
            text = build_activity_text(res.activity_base, res.part, cfg.rules.activity_name_max_len)
            activity = node4.activities.get(text)
            if activity is None:
                activity = node4.activities[text] = ActivityPlan(
                    key=res.key, part=res.part, base_name=res.activity_base, text=text)
            activity.add(row.code, qty, row.description, (row.sheet, row.row))
        return self._sorted(nodes)

    def _sorted(self, nodes: Dict[str, Wbs3Plan]) -> Dict[str, Wbs3Plan]:
        order3 = list(self.cfg.wbs.level3_desc.keys())
        part_order = {p: i for i, p in enumerate(self.cfg.rules.part_order)}

        def n3_key(n3: str):
            return (order3.index(n3) if n3 in order3 else len(order3), n3)

        result: Dict[str, Wbs3Plan] = {}
        for n3 in sorted(nodes, key=n3_key):
            node3 = nodes[n3]
            children = sorted(node3.children.values(),
                              key=lambda n: (n.network, n.pt is None, n.pt or 0))
            node3.children = {c.code: c for c in children}
            for child in children:
                acts = sorted(child.activities.values(),
                              key=lambda a: (a.key, part_order.get(a.part, 99), a.text))
                child.activities = {a.text: a for a in acts}
                for act in acts:
                    act.services = dict(sorted(act.services.items()))
            result[n3] = node3
        return result

    def _apply_project_fields(self, nodes: Dict[str, Wbs3Plan], budget: str, ctx) -> None:
        wcfg = self.cfg.wbs
        location = ctx.fields.get("location", "")
        for node3 in nodes.values():
            if location:
                node3.assignments["STORT"] = location
            tplnr = wcfg.level3_tplnr.get(node3.n3) or wcfg.default_tplnr
            if tplnr:
                node3.assignments["TPLNR"] = tplnr
            if 3 in wcfg.budget_on_levels and budget:
                node3.custom_fields["ZZCOD_BUGET"] = budget
            for node4 in node3.children.values():
                if location:
                    node4.assignments["STORT"] = location
                field_name = wcfg.assignment_field_by_network.get(node4.network, wcfg.assignment_default_field)
                value = ""
                if node4.pt is not None and node4.pt in ctx.pts:
                    value = ctx.pts[node4.pt].values.get(field_name, "")
                if not value:
                    value = wcfg.default_equnr if field_name == "EQUNR" else wcfg.default_tplnr
                if value:
                    node4.assignments[field_name] = value
                if 4 in wcfg.budget_on_levels and budget:
                    node4.custom_fields["ZZCOD_BUGET"] = budget
                for name, val in wcfg.custom_fields_by_network.get(node4.network, {}).items():
                    if val not in (None, ""):
                        node4.custom_fields[name] = str(val)


def describe_plan(plan: ProjectPlan, decimals: int = 3, decimal_separator: str = ",") -> str:
    """Reprezentare text a arborelui (pentru log / consolă)."""

    lines = [f"Proiect: {plan.project_definition or '<lipsă>'} | Cod buget: {plan.budget_code or '<lipsă>'}"]
    for node3 in plan.level3.values():
        lines.append(f"{node3.code}  {node3.description}")
        for node4 in node3.children.values():
            extra = ", ".join(f"{k}={v}" for k, v in {**node4.assignments, **node4.custom_fields}.items())
            lines.append(f"  {node4.code}  {node4.description}  [{node4.network}]"
                         + (f"  ({extra})" if extra else ""))
            for act in node4.activities.values():
                services = "; ".join(
                    f"{line.code}={format_quantity(line.quantity, decimals, decimal_separator)}"
                    for line in act.services.values())
                lines.append(f"      {act.text}: {services}")
    return "\n".join(lines)


# =============================================================================
#  07. RAPOARTE EXCEL
# =============================================================================
#
#  Rapoartele Excel: planul (pre-flight) și rezultatele execuției.
# =============================================================================

def _autosize(writer, sheet_name: str, frame) -> None:
    ws = writer.sheets[sheet_name]
    for idx, column in enumerate(frame.columns, start=1):
        values = [str(column)] + [str(v) for v in frame[column].head(500).tolist()]
        width = min(80, max(len(v) for v in values) + 2)
        ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = width
    ws.freeze_panes = "A2"


def plan_frames(plan: ProjectPlan, read: Optional[ReadResult] = None):
    import pandas as pd

    plan_rows = []
    for n3, n4 in plan.iter_wbs4():
        for act in n4.activities.values():
            for line in act.services.values():
                plan_rows.append({
                    "Nivel 3": n3.code, "Denumire N3": n3.description,
                    "Nivel 4": n4.code, "Denumire N4": n4.description,
                    "Retea standard": n4.network, "PT": n4.pt if n4.pt is not None else "comun",
                    "Activitate": act.text, "Partea": act.part,
                    "Cod serviciu": line.code, "Cantitate": round(line.quantity, 6),
                    "Descriere": line.description, "Surse Excel": line.source_refs,
                    "Atribuiri": ", ".join(f"{k}={v}" for k, v in n4.assignments.items()),
                    "Campuri client": ", ".join(f"{k}={v}" for k, v in n4.custom_fields.items()),
                })
    issues = [{"Severitate": i.severity, "Locatie": i.location(), "Cod": i.code, "Mesaj": i.message}
              for i in plan.issues]
    summary = [{"Indicator": k, "Valoare": v} for k, v in plan.stats.items()]
    summary.insert(0, {"Indicator": "Definitie proiect", "Valoare": plan.project_definition or "<lipsa>"})
    summary.insert(1, {"Indicator": "Cod buget", "Valoare": plan.budget_code or "<lipsa>"})
    frames = {
        "Plan": pd.DataFrame(plan_rows),
        "Probleme": pd.DataFrame(issues, columns=["Severitate", "Locatie", "Cod", "Mesaj"]),
        "Sumar": pd.DataFrame(summary),
    }
    if read is not None:
        frames["Coloane detectate"] = pd.DataFrame([
            {"Foaie": h.sheet, "Rand antet": h.header_row,
             **{f"col {k}": h.letter(k) for k in h.columns},
             "PT": ",".join(f"PT{n}" for n in sorted(h.pt_columns))}
            for h in read.headers])
    return frames


def write_plan_report(plan: ProjectPlan, path: Path, read: Optional[ReadResult] = None) -> Path:
    import pandas as pd

    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, frame in plan_frames(plan, read).items():
            frame.to_excel(writer, sheet_name=name, index=False)
            _autosize(writer, name, frame)
    log.info("Raport pre-flight scris: %s", path)
    return path


def write_results_report(results, path: Path) -> Path:
    import pandas as pd

    path.parent.mkdir(parents=True, exist_ok=True)
    frame = pd.DataFrame([{
        "Nivel 3": r.level3, "WBS": r.wbs, "Retea": r.network, "Activitate": r.activity,
        "Cod serviciu": r.code, "Cantitate": r.quantity, "Surse Excel": r.sources,
        "Status": r.status, "Mesaj SAP": r.message,
    } for r in results.rows.values()])
    events = pd.DataFrame(results.events, columns=["WBS", "Status", "Mesaj"])
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="Rezultate", index=False)
        _autosize(writer, "Rezultate", frame)
        events.to_excel(writer, sheet_name="Evenimente WBS", index=False)
        _autosize(writer, "Evenimente WBS", events)
    log.info("Raport rezultate scris: %s", path)
    return path


# =============================================================================
#  08. ID-URI SAP GUI (din înregistrarea Script1.vbs)
# =============================================================================
#
#  ID-urile elementelor SAP GUI, preluate din înregistrarea VBS (Script1.vbs).
#
#  Toate ID-urile sunt centralizate aici: dacă SAP-ul (sau layout-ul) se schimbă,
#  se modifică doar acest fișier. Elementele marcate ``NEVERIFICAT`` nu apar în
#  înregistrare și trebuie confirmate cu o înregistrare scurtă (vezi README).
# =============================================================================

class SapIds:
    """Toate ID-urile elementelor SAP GUI; se modifică doar aici."""

    # --- generale --------------------------------------------------------------- #
    MAIN_WINDOW = "wnd[0]"
    STATUS_BAR = "wnd[0]/sbar"
    OKCODE = "wnd[0]/tbar[0]/okcd"
    BTN_BACK = "wnd[0]/tbar[0]/btn[3]"      # F3
    BTN_SAVE = "wnd[0]/tbar[0]/btn[11]"     # Ctrl+S

    # --- ecranul inițial CJ02 ------------------------------------------------------ #
    INITIAL_PROJECT_FIELD = "wnd[0]/usr/ctxt*PROJ-PSPID"   # NEVERIFICAT (câmpul "Definiție proiect")
    INITIAL_WBS_FIELD = "wnd[0]/usr/ctxt*PRPS-POSID"       # NEVERIFICAT (se golește)
    BTN_WBS_OVERVIEW = "wnd[0]/tbar[1]/btn[5]"             # din înregistrare: după salvare -> lista WBS

    # --- lista elementelor WBS -------------------------------------------------- #
    WBS_TABSTRIP = "wnd[0]/usr/tabsTABCJLE"
    TAB_WBS_GENERAL = WBS_TABSTRIP + "/tabpLEGR"
    TAB_WBS_LETR = WBS_TABSTRIP + "/tabpLETR"
    TAB_WBS_ASSIGN = WBS_TABSTRIP + "/tabpLECO"
    WBS_TABLE = TAB_WBS_GENERAL + "/ssubLISTE:SAPLCJWB:0902/tblSAPLCJWBTAB_902"
    WBS_ASSIGN_TABLE = TAB_WBS_ASSIGN + "/ssubLISTE:SAPLCJWB:0903/tblSAPLCJWBTAB_903"

    # câmp -> (nume câmp SAP, coloana din înregistrare folosită ca rezervă)
    WBS_COL_LEVEL = ("PRPS-STUFE", 1)
    WBS_COL_IDENT = ("RCWBS-IDENT", 2)
    WBS_COL_DESC = ("PRPS-POST1", 3)
    ASSIGN_COLS = {
        "STORT": ("PRPS-STORT", 7),
        "EQUNR": ("PRPS-EQUNR", 14),
        "TPLNR": ("PRPS-TPLNR", 15),
    }

    MENU_WBS_DETAIL = "wnd[0]/mbar/menu[4]/menu[3]"
    BTN_WBS_CUSTOMER_TAB = "wnd[0]/tbar[1]/btn[19]"
    CUSTOM_FIELDS = {
        "ZZCOD_BUGET": "wnd[0]/usr/ssubCUSTSCR1:SAPLXCN1:0700/ctxtPRPS-ZZCOD_BUGET",
        "ZZPARTE": "wnd[0]/usr/ssubCUSTSCR1:SAPLXCN1:0700/txtPRPS-ZZPARTE",
    }

    # --- activități (rețea) ------------------------------------------------------- #
    TAB_ACTIVITIES = "wnd[0]/usr/tabsTABSTR199/tabpVGTS"
    MENU_COPY_STD_NETWORK = "wnd[0]/mbar/menu[0]/menu[5]"
    STD_NETWORK_NR = "wnd[1]/usr/ctxtRCN10-STDNR"
    STD_NETWORK_ALT = "wnd[1]/usr/txtRCN10-STDAL"
    POPUP_OK = "wnd[1]/tbar[0]/btn[0]"

    TAB_ACT_EXTERNAL = "wnd[0]/usr/tabsTABSTRIP_2000/tabpFRML"
    ACT_TABLE = TAB_ACT_EXTERNAL + "/ssubSUBSCR_2000:SAPLCOVG:2050/tblSAPLCOVGTCTRL_2050"
    ACT_DELETE_BTN = TAB_ACT_EXTERNAL + "/ssubSUBSCR_2000:SAPLCOVG:2050/btnBUTTON_DELETE"
    ACT_COL_VORNR = ("AFVGD-VORNR", 0)
    ACT_COL_TEXT = ("AFVGD-LTXA1", 2)

    # --- detaliu activitate (fereastra modală wnd[1]) ------------------------------- #
    ACT_DETAIL_WINDOW = "wnd[1]"
    ACT_DETAIL_TEXT = "wnd[1]/usr/subIDENTIFICATION:SAPLCONW:0100/txtAFVGM-LTXA1"
    ACT_DETAIL_TAB_EXTERNAL = "wnd[1]/usr/tabsTABSTRIP_1000/tabpFRMD"
    ACT_DETAIL_EKGRP = ACT_DETAIL_TAB_EXTERNAL + "/ssubSUBSCR_1000:SAPLCONW:1510/ctxtAFVGD-EKGRP"
    ACT_DETAIL_AFNAM = ACT_DETAIL_TAB_EXTERNAL + "/ssubSUBSCR_1000:SAPLCONW:1510/txtAFVGD-AFNAM"
    ACT_DETAIL_SERVICES_BTN = "wnd[1]/tbar[0]/btn[6]"
    ACT_DETAIL_CLOSE_BTN = "wnd[1]/tbar[0]/btn[5]"

    # --- specificație servicii ------------------------------------------------------ #
    SRV_TABLE = "wnd[0]/usr/subSERVICE:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW"
    SRV_COL_CODE = ("ESLL-SRVPOS", 2)
    SRV_COL_QTY = ("ESLL-MENGE", 4)
    SRV_COL_PRICE = ("ESLL-TBTWR", 6)
    SRV_COL_UNLIMITED = ("ESLL-UEBTK", 8)

    # --- pop-up-uri standard ------------------------------------------------------- #
    POPUP_VALUE_PREFIX = "usr/sub:SAPLSPO4:0300"     # cerere de valori (ex. AMC / MBT1)
    POPUP_YES = "usr/btnSPOP-VAROPTION1"
    POPUP_NO = "usr/btnSPOP-VAROPTION2"
    POPUP_OPTION_YES = "usr/btnSPOP-OPTION1"
    POPUP_OPTION_NO = "usr/btnSPOP-OPTION2"
    POPUP_ERRORLOG_EDIT = "usr/btnDY_VAROPTION2"      # salvare cu erori -> înapoi la editare
    POPUP_ERRORLOG_OPT1 = "usr/btnDY_VAROPTION1"


ids = SapIds


# =============================================================================
#  09. CONEXIUNEA SAP, BARA DE STARE, POP-UP-URI, TABELE
# =============================================================================
#
#  Conexiunea la SAP GUI Scripting și operațiile generice pe ecrane.
#
#  * ``SapSession.attach`` -> GetObject("SAPGUI") -> GetScriptingEngine -> sesiunea potrivită;
#  * ``status()`` citește bara de stare (``wnd[0]/sbar``) după fiecare acțiune critică;
#  * ``handle_popups()`` tratează ferestrele modale neașteptate (informări,
#    confirmări, cereri de valori), cu log pentru fiecare;
#  * ``SapTable`` citește/scrie tabelele SAP (table control) cu derulare automată.
#
#  Modulul nu importă ``win32com`` decât la conectare, deci restul aplicației
#  (citire Excel, plan, teste) funcționează și fără SAP / Windows.
# =============================================================================

# --------------------------------------------------------------------------- #
#  Excepții
# --------------------------------------------------------------------------- #
class SapError(Exception):
    """Eroare generică SAP."""


class SapConnectionError(SapError):
    """Nu se poate realiza conexiunea la SAP GUI."""


class SapActionError(SapError):
    """O acțiune în SAP a eșuat (mesaj de eroare în bara de stare / pop-up)."""

    def __init__(self, message: str, status: Optional["StatusMessage"] = None):
        super().__init__(message)
        self.status = status


class SapLineError(SapActionError):
    """Eroare la un singur rând de serviciu (normă); rândul a fost anulat, se poate continua."""


@dataclass
class StatusMessage:
    type: str = ""     # S=succes, E=eroare, W=avertizare, I=informare, A=abandon
    text: str = ""
    msg_id: str = ""
    number: str = ""

    @property
    def is_error(self) -> bool:
        return self.type in ("E", "A")

    @property
    def is_warning(self) -> bool:
        return self.type == "W"

    def __str__(self) -> str:
        if not self.text:
            return "(fără mesaj)"
        code = f" [{self.msg_id}{self.number}]" if self.msg_id else ""
        return f"{self.type}: {self.text}{code}"


# --------------------------------------------------------------------------- #
#  Sesiunea
# --------------------------------------------------------------------------- #
class SapSession:
    def __init__(self, session: Any, cfg: SapConfig):
        self.session = session
        self.cfg = cfg

    # -- conectare ---------------------------------------------------------- #
    @classmethod
    def attach(cls, cfg: SapConfig) -> "SapSession":
        """Se atașează la o sesiune SAP GUI deja autentificată și maximizează fereastra."""
        try:
            import win32com.client  # type: ignore
        except ImportError as exc:  # pragma: no cover - doar pe Windows
            raise SapConnectionError("Lipsește pywin32 (pip install pywin32); rulează doar pe Windows") from exc
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
        except Exception as exc:  # noqa: BLE001
            raise SapConnectionError("SAP GUI nu rulează. Porniți SAP Logon și autentificați-vă.") from exc
        application = sap_gui_auto.GetScriptingEngine
        if application is None:
            raise SapConnectionError("Nu pot obține motorul de scripting SAP GUI")

        candidates: List[Tuple[Any, Dict[str, str]]] = []
        for i in range(application.Children.Count):
            connection = application.Children(i)
            if getattr(connection, "DisabledByServer", False):
                log.warning("Conexiunea %d are scripting dezactivat pe server (RZ11 sapgui/user_scripting)", i)
                continue
            for j in range(connection.Children.Count):
                sess = connection.Children(j)
                info = sess.Info
                meta = {
                    "id": sess.Id,
                    "system": str(info.SystemName),
                    "client": str(info.Client),
                    "user": str(info.User),
                    "transaction": str(info.Transaction),
                    "busy": str(bool(sess.Busy)),
                }
                candidates.append((sess, meta))
        if not candidates:
            raise SapConnectionError("Nu există nicio sesiune SAP deschisă (autentificați-vă în SAP Logon)")
        for _, meta in candidates:
            log.info("Sesiune disponibilă: %s", meta)

        def wanted(meta: Dict[str, str]) -> bool:
            return all(not expected or meta[key].upper() == expected.upper()
                       for key, expected in (("system", cfg.system_name), ("client", cfg.client),
                                             ("user", cfg.user)))

        matching = [(s, m) for s, m in candidates if wanted(m) and m["busy"] == "False"]
        if not matching:
            raise SapConnectionError(
                f"Nicio sesiune liberă nu corespunde filtrului (sistem='{cfg.system_name}', "
                f"mandant='{cfg.client}', utilizator='{cfg.user}')")
        # preferăm o sesiune aflată în meniul principal, ca să nu deranjăm lucrul în curs
        idle = [c for c in matching if c[1]["transaction"] in ("SESSION_MANAGER", "SMEN", "S000")]
        session, meta = (idle or matching)[0]
        log.info("Mă atașez la sesiunea %s (sistem %s, mandant %s, utilizator %s)",
                 meta["id"], meta["system"], meta["client"], meta["user"])
        sap = cls(session, cfg)
        sap.find(ids.MAIN_WINDOW).maximize()
        return sap

    # -- primitive ---------------------------------------------------------- #
    def wait(self) -> None:
        """Așteaptă cât sesiunea este ocupată (SAP GUI e sincron, dar e o plasă de siguranță)."""
        deadline = time.monotonic() + self.cfg.busy_timeout_s
        while self.session.Busy:
            if time.monotonic() > deadline:
                raise SapActionError(f"SAP nu a răspuns în {self.cfg.busy_timeout_s} s")
            time.sleep(0.2)

    def find(self, element_id: str, required: bool = True) -> Any:
        try:
            element = self.session.findById(element_id, False)
        except Exception:  # noqa: BLE001
            element = None
        if element is None and required:
            raise SapActionError(f"Elementul SAP nu a fost găsit: {element_id}")
        return element

    def exists(self, element_id: str) -> bool:
        return self.find(element_id, required=False) is not None

    def set_text(self, element_id: str, value: str) -> None:
        self.find(element_id).Text = value

    def press(self, element_id: str) -> None:
        self.find(element_id).press()
        self.wait()

    def select(self, element_id: str) -> None:
        self.find(element_id).select()
        self.wait()

    def send_vkey(self, key: int, window: int = 0) -> None:
        self.find(f"wnd[{window}]").sendVKey(key)
        self.wait()

    # -- bara de stare ------------------------------------------------------ #
    def status(self) -> StatusMessage:
        """Citește mesajul curent din bara de stare SAP (wnd[0]/sbar)."""
        bar = self.find(ids.STATUS_BAR, required=False)
        if bar is None:
            return StatusMessage()
        return StatusMessage(
            type=str(getattr(bar, "MessageType", "") or ""),
            text=str(getattr(bar, "Text", "") or ""),
            msg_id=str(getattr(bar, "MessageId", "") or "").strip(),
            number=str(getattr(bar, "MessageNumber", "") or "").strip(),
        )

    def enter(self, accept_warnings: bool = True) -> StatusMessage:
        """Enter pe fereastra principală; avertizările (W) se confirmă cu încă un Enter."""
        self.send_vkey(0)
        status = self.status()
        attempts = 0
        while accept_warnings and status.is_warning and attempts < 3:
            log.info("Avertizare SAP confirmată: %s", status.text)
            self.send_vkey(0)
            status = self.status()
            attempts += 1
        return status

    def raise_on_error(self, context: str, status: Optional[StatusMessage] = None) -> StatusMessage:
        status = status or self.status()
        if status.is_error:
            raise SapActionError(f"{context}: {status.text}", status)
        return status

    def start_transaction(self, tcode: str) -> None:
        """Deschide tranzacția prin câmpul de comandă (/n închide tranzacția curentă)."""
        self.close_all_popups()
        self.set_text(ids.OKCODE, f"/n{tcode}")
        self.send_vkey(0)
        self.raise_on_error(f"Deschiderea tranzacției {tcode}")
        log.info("Tranzacția %s deschisă", tcode)

    # -- pop-up-uri ---------------------------------------------------------- #
    def top_window(self) -> int:
        level = 0
        for i in range(1, 6):
            if self.exists(f"wnd[{i}]"):
                level = i
        return level

    def popup_text(self, level: int) -> str:
        window = self.find(f"wnd[{level}]", required=False)
        if window is None:
            return ""
        texts: List[str] = []
        try:
            title = str(window.Text or "").strip()
            if title:
                texts.append(title)
            usr = self.find(f"wnd[{level}]/usr", required=False)
            stack = [usr] if usr is not None else []
            while stack and len(texts) < 15:
                node = stack.pop(0)
                for idx in range(node.Children.Count):
                    child = node.Children(idx)
                    ctype = str(child.Type)
                    if ctype in ("GuiLabel", "GuiTextField", "GuiTextedit") and str(child.Text).strip():
                        texts.append(str(child.Text).strip())
                    elif ctype in ("GuiSimpleContainer", "GuiUserArea", "GuiScrollContainer",
                                   "GuiTab", "GuiTabStrip", "GuiBox"):
                        stack.append(child)
        except Exception:  # noqa: BLE001 - textul e doar informativ
            pass
        return " | ".join(texts)

    def _find_value_field(self, level: int) -> Any:
        container = self.find(f"wnd[{level}]/{ids.POPUP_VALUE_PREFIX}", required=False)
        if container is None:
            return None
        stack = [container]
        while stack:
            node = stack.pop(0)
            for idx in range(node.Children.Count):
                child = node.Children(idx)
                if "SVALD-VALUE" in str(child.Id) and bool(getattr(child, "Changeable", True)):
                    return child
                if getattr(child, "Children", None) is not None and child.Children.Count:
                    stack.append(child)
        return None

    def handle_popups(self, base_level: int = 0, values: Optional[Sequence[str]] = None,
                      confirm: bool = True, context: str = "") -> List[str]:
        """Tratează ferestrele modale de deasupra nivelului ``base_level``.

        * cerere de valori (SAPLSPO4) -> se completează următoarea valoare din ``values``;
        * întrebare Da/Nu -> Da dacă ``confirm`` altfel Nu;
        * jurnal de erori la salvare -> "înapoi la editare" și excepție;
        * informare -> Enter (dacă ``auto_confirm_info_popups``), altfel excepție.
        Întoarce textele pop-up-urilor tratate.
        """
        queue = list(values or [])
        handled: List[str] = []
        previous, repeats = None, 0
        for _ in range(self.cfg.max_popups):
            level = self.top_window()
            if level <= base_level:
                return handled
            wnd = f"wnd[{level}]"
            text = self.popup_text(level)
            repeats = repeats + 1 if (level, text) == previous else 0
            if repeats >= 2:
                raise SapActionError(f"Pop-up-ul nu se închide ({context}): {text}")
            previous = (level, text)
            handled.append(text)
            log.info("Pop-up SAP%s: %s", f" ({context})" if context else "", text or "(fără text)")
            value_field = self._find_value_field(level)
            if value_field is not None:
                if not queue:
                    raise SapActionError(f"SAP cere o valoare suplimentară ({text}), dar lista "
                                         "sap.network_popup_values este epuizată")
                value = queue.pop(0)
                value_field.Text = value
                log.info("Pop-up: completez valoarea '%s'", value)
                self.press(f"{wnd}/tbar[0]/btn[0]")
            elif self.exists(f"{wnd}/{ids.POPUP_ERRORLOG_EDIT}"):
                self.press(f"{wnd}/{ids.POPUP_ERRORLOG_EDIT}")
                raise SapActionError(f"SAP a raportat erori ({context}): {text}")
            elif self.exists(f"{wnd}/{ids.POPUP_YES}"):
                self.press(f"{wnd}/{ids.POPUP_YES if confirm else ids.POPUP_NO}")
            elif self.exists(f"{wnd}/{ids.POPUP_OPTION_YES}"):
                self.press(f"{wnd}/{ids.POPUP_OPTION_YES if confirm else ids.POPUP_OPTION_NO}")
            elif self.cfg.auto_confirm_info_popups:
                if self.exists(f"{wnd}/tbar[0]/btn[0]"):
                    self.press(f"{wnd}/tbar[0]/btn[0]")
                else:
                    self.send_vkey(0, level)
            else:
                raise SapActionError(f"Pop-up neașteptat ({context}): {text}")
        raise SapActionError(f"Prea multe pop-up-uri consecutive ({context})")

    def close_all_popups(self) -> None:
        """Închide toate ferestrele modale (folosit la recuperarea după erori)."""
        for _ in range(6):
            level = self.top_window()
            if level == 0:
                return
            log.warning("Închid fereastra wnd[%d]: %s", level, self.popup_text(level))
            try:
                self.find(f"wnd[{level}]").close()
                self.wait()
            except Exception:  # noqa: BLE001
                try:
                    self.send_vkey(12, level)  # F12 = Anulare
                except Exception:  # noqa: BLE001
                    return
            # o confirmare de tip "Datele se pierd?" poate apărea
            new_level = self.top_window()
            if new_level >= level and self.exists(f"wnd[{new_level}]/{ids.POPUP_YES}"):
                self.press(f"wnd[{new_level}]/{ids.POPUP_YES}")


# --------------------------------------------------------------------------- #
#  Tabele SAP (GuiTableControl)
# --------------------------------------------------------------------------- #
class SapTable:
    """Acces la un table control SAP, cu derulare automată și coloane găsite după nume.

    Indexul de rând folosit de metode este ABSOLUT (0 = primul rând al tabelului).
    """

    def __init__(self, sap: SapSession, table_id: str):
        self.sap = sap
        self.table_id = table_id
        self._col_cache: Dict[str, int] = {}

    def obj(self) -> Any:
        # obiectul se recitește după fiecare interacțiune: SAP îl reconstruiește la fiecare PAI
        return self.sap.find(self.table_id)

    def column(self, spec: Tuple[str, int]) -> int:
        """Indexul coloanei după numele câmpului (ex. "ESLL-SRVPOS"); rezervă: indexul înregistrat."""
        name, fallback = spec
        if name in self._col_cache:
            return self._col_cache[name]
        index = fallback
        try:
            table = self.obj()
            for col in range(table.Columns.Count):
                try:
                    if str(table.GetCell(0, col).Name).upper() == name.upper():
                        index = col
                        break
                except Exception:  # noqa: BLE001
                    continue
        except Exception:  # noqa: BLE001
            pass
        self._col_cache[name] = index
        return index

    def scroll_to(self, row: int) -> None:
        table = self.obj()
        visible = int(table.VisibleRowCount)
        position = int(table.VerticalScrollbar.Position)
        if position <= row < position + visible:
            return
        maximum = int(table.VerticalScrollbar.Maximum)
        table.VerticalScrollbar.Position = max(0, min(row, maximum))
        self.sap.wait()

    def cell(self, row: int, spec: Tuple[str, int]) -> Any:
        col = self.column(spec)
        self.scroll_to(row)
        table = self.obj()
        return table.GetCell(row - int(table.VerticalScrollbar.Position), col)

    def text(self, row: int, spec: Tuple[str, int]) -> str:
        try:
            return str(self.cell(row, spec).Text).strip()
        except Exception:  # noqa: BLE001 - rând inexistent
            return ""

    def set_text(self, row: int, spec: Tuple[str, int], value: str) -> None:
        self.cell(row, spec).Text = value

    def set_checkbox(self, row: int, spec: Tuple[str, int], value: bool = True) -> None:
        self.cell(row, spec).Selected = value

    def select_row(self, row: int, selected: bool = True) -> None:
        self.scroll_to(row)
        self.obj().GetAbsoluteRow(row).Selected = selected

    def focus(self, row: int, spec: Tuple[str, int]) -> None:
        self.cell(row, spec).SetFocus()

    def read_rows(self, specs: Sequence[Tuple[str, int]], key: Tuple[str, int],
                  max_rows: int = 2000) -> List[List[str]]:
        """Citește rândurile până la primul rând cu cheia goală. Întoarce [[text_col1, ...], ...]."""
        cols = [self.column(spec) for spec in specs]
        key_col = self.column(key)
        result: List[List[str]] = []
        row = 0
        table = self.obj()
        total = int(table.RowCount)
        while row < min(total, max_rows):
            self.scroll_to(row)
            table = self.obj()
            position = int(table.VerticalScrollbar.Position)
            visible = int(table.VisibleRowCount)
            if visible <= 0 or not position <= row < position + visible:
                break  # nu se mai poate derula (sfârșitul tabelului)
            for vis in range(row - position, visible):
                try:
                    key_text = str(table.GetCell(vis, key_col).Text).strip()
                except Exception:  # noqa: BLE001
                    return result
                if not key_text:
                    return result
                values = []
                for col in cols:
                    try:
                        values.append(str(table.GetCell(vis, col).Text).strip())
                    except Exception:  # noqa: BLE001
                        values.append("")
                result.append(values)
                row += 1
        return result

    def first_empty_row(self, key: Tuple[str, int]) -> int:
        return len(self.read_rows([key], key))

    def visible_rows(self) -> Iterable[int]:
        table = self.obj()
        position = int(table.VerticalScrollbar.Position)
        return range(position, position + int(table.VisibleRowCount))


# =============================================================================
#  10. ACȚIUNI ÎN TRANZACȚIA DE PROIECT (CJ02)
# =============================================================================
#
#  Acțiunile din tranzacția de proiect (CJ02), traduse din înregistrarea VBS.
#
#  Fiecare metodă publică este o operație parametrizată, verificată prin bara de
#  stare și cu tratarea pop-up-urilor. Funcțiile pure de la final (fără SAP) sunt
#  testate unitar.
# =============================================================================

@dataclass
class WbsRow:
    row: int          # index absolut în tabel
    level: int
    ident: str
    description: str


@dataclass
class ActivityRow:
    row: int
    vornr: str
    text: str


# --------------------------------------------------------------------------- #
#  Funcții pure
# --------------------------------------------------------------------------- #
def ident_matches(existing: str, code: str) -> bool:
    """Compară un cod WBS afișat în SAP cu codul planificat ("+453R1").

    SAP poate afișa forma scurtă (+453R1) sau codul complet (C-0001-453R1);
    potrivirea se face pe sufix, la limită de separator.
    """
    shown = re.sub(r"\s+", "", existing or "").upper().lstrip("+")
    wanted = re.sub(r"\s+", "", code or "").upper().lstrip("+")
    if not shown or not wanted:
        return False
    if shown == wanted:
        return True
    if shown.endswith(wanted):
        return not shown[-len(wanted) - 1].isalnum()
    return False


def find_wbs(rows: Iterable[WbsRow], code: str) -> Optional[WbsRow]:
    return next((r for r in rows if ident_matches(r.ident, code)), None)


def parent_of_new_row(rows: Sequence[WbsRow], level: int) -> Optional[WbsRow]:
    """Părintele pe care îl primește un rând nou adăugat la finalul listei."""
    for row in reversed(rows):
        if row.level < level:
            return row
    return None


def plan_copy_rounds(activities: Sequence) -> List[List]:
    """Grupează activitățile în runde de copiere a rețelei standard.

    O copiere aduce câte o activitate din fiecare tip (XL4, XN4 ...); dacă aceeași
    activitate e necesară de mai multe ori (L, M, E), sunt necesare mai multe copieri.
    """
    rounds: List[List] = []
    for act in activities:
        for current in rounds:
            if all(other.key != act.key for other in current):
                current.append(act)
                break
        else:
            rounds.append([act])
    return rounds


def choose_rows_to_keep(new_rows: Sequence[ActivityRow], keys: Sequence[str],
                        key_regex: str) -> Tuple[Dict[str, ActivityRow], List[ActivityRow], List[str]]:
    """Din activitățile aduse de rețeaua standard, păstrează câte una pentru fiecare cheie.

    Întoarce (păstrate {cheie: rând}, de șters [rânduri], chei negăsite).
    """
    keep: Dict[str, ActivityRow] = {}
    for row in new_rows:
        match = re.match(key_regex, row.text.strip().upper())
        key = match.group(1) if match else None
        if key in keys and key not in keep:
            keep[key] = row
    kept_rows = {id(r) for r in keep.values()}
    delete = [r for r in new_rows if id(r) not in kept_rows]
    missing = [k for k in keys if k not in keep]
    return keep, delete, missing


# --------------------------------------------------------------------------- #
#  Operații SAP
# --------------------------------------------------------------------------- #
class ProjectBuilder:
    def __init__(self, sap: SapSession, cfg: AppConfig):
        self.sap = sap
        self.cfg = cfg

    # -- tranzacție / proiect ----------------------------------------------- #
    def open_project(self, project: str) -> None:
        """Deschide tranzacția și lista elementelor WBS ale proiectului."""
        sap = self.sap
        sap.start_transaction(self.cfg.sap.transaction)
        if not sap.exists(ids.INITIAL_PROJECT_FIELD):
            raise SapActionError(
                "Câmpul 'Definiție proiect' nu a fost găsit pe ecranul inițial "
                f"(verificați sap_ids.INITIAL_PROJECT_FIELD = {ids.INITIAL_PROJECT_FIELD})")
        sap.set_text(ids.INITIAL_PROJECT_FIELD, project)
        if sap.exists(ids.INITIAL_WBS_FIELD):
            sap.set_text(ids.INITIAL_WBS_FIELD, "")
        if sap.exists(ids.BTN_WBS_OVERVIEW):
            sap.press(ids.BTN_WBS_OVERVIEW)
        else:
            sap.send_vkey(0)
        sap.handle_popups(context=f"deschidere proiect {project}")
        sap.raise_on_error(f"Deschiderea proiectului {project}")
        if not sap.exists(ids.WBS_TABLE) and sap.exists(ids.TAB_WBS_GENERAL):
            sap.select(ids.TAB_WBS_GENERAL)
        if not sap.exists(ids.WBS_TABLE):
            raise SapActionError(f"Lista elementelor WBS nu s-a deschis pentru proiectul {project}")
        log.info("Proiectul %s deschis (lista WBS)", project)

    def _wbs_table(self) -> SapTable:
        if not self.sap.exists(ids.WBS_TABLE) and self.sap.exists(ids.TAB_WBS_GENERAL):
            self.sap.select(ids.TAB_WBS_GENERAL)
        return SapTable(self.sap, ids.WBS_TABLE)

    def read_wbs_rows(self) -> List[WbsRow]:
        table = self._wbs_table()
        rows = []
        for index, (level, ident, desc) in enumerate(
                table.read_rows([ids.WBS_COL_LEVEL, ids.WBS_COL_IDENT, ids.WBS_COL_DESC],
                                key=ids.WBS_COL_IDENT)):
            try:
                level_no = int(level)
            except ValueError:
                level_no = 99
            rows.append(WbsRow(index, level_no, ident, desc))
        return rows

    def create_wbs_element(self, parent_code: Optional[str], code: str, level: int, description: str) -> int:
        """Creează un element WBS (nivel 3 sau 4) la finalul listei; întoarce rândul lui.

        Ierarhia în lista WBS e dată de ordine + nivel, deci se verifică întâi că
        noul element ajunge sub părintele corect.
        """
        rows = self.read_wbs_rows()
        existing = find_wbs(rows, code)
        if existing:
            log.info("WBS %s există deja (rândul %d)", code, existing.row + 1)
            return existing.row
        if parent_code:
            parent = parent_of_new_row(rows, level)
            if parent is None or not ident_matches(parent.ident, parent_code):
                raise SapActionError(
                    f"WBS {code} ar ajunge sub '{parent.ident if parent else '-'}', nu sub {parent_code}; "
                    "inserarea în mijlocul listei nu este suportată")
        row = len(rows)
        table = self._wbs_table()
        table.set_text(row, ids.WBS_COL_LEVEL, str(level))
        table.set_text(row, ids.WBS_COL_IDENT, code)
        table.set_text(row, ids.WBS_COL_DESC, description)
        status = self.sap.enter()
        self.sap.handle_popups(context=f"creare WBS {code}")
        self.sap.raise_on_error(f"Crearea WBS {code}", status if status.is_error else None)
        if not find_wbs(self.read_wbs_rows(), code):
            raise SapActionError(f"WBS {code} nu apare în listă după creare ({status})")
        log.info("WBS %s (nivel %d, '%s') adăugat. SAP: %s", code, level, description, status)
        return row

    def set_assignments(self, row: int, values: Dict[str, str]) -> None:
        """Completează atribuirile (locație STORT, loc funcțional TPLNR, echipament EQUNR)."""
        if not values:
            return
        self.sap.select(ids.TAB_WBS_ASSIGN)
        table = SapTable(self.sap, ids.WBS_ASSIGN_TABLE)
        for field_name, value in values.items():
            spec = ids.ASSIGN_COLS.get(field_name)
            if spec is None:
                raise SapActionError(f"Câmp de atribuire necunoscut: {field_name}")
            table.set_text(row, spec, value)
        status = self.sap.enter()
        self.sap.handle_popups(context="atribuiri WBS")
        self.sap.raise_on_error("Atribuirile WBS", status if status.is_error else None)
        self.sap.select(ids.TAB_WBS_GENERAL)

    def set_custom_fields(self, row: int, fields: Dict[str, str]) -> None:
        """Câmpurile client din detaliul WBS (ex. ZZCOD_BUGET = cod buget)."""
        if not fields:
            return
        self.sap.select(ids.TAB_WBS_ASSIGN)
        SapTable(self.sap, ids.WBS_ASSIGN_TABLE).select_row(row, True)
        self.sap.select(ids.MENU_WBS_DETAIL)
        self.sap.press(ids.BTN_WBS_CUSTOMER_TAB)
        for name, value in fields.items():
            element = ids.CUSTOM_FIELDS.get(name)
            if element is None:
                raise SapActionError(f"Câmp client necunoscut în sap_ids.CUSTOM_FIELDS: {name}")
            self.sap.set_text(element, value)
        status = self.sap.enter()
        self.sap.handle_popups(context="câmpuri client WBS")
        self.sap.raise_on_error("Câmpurile client WBS", status if status.is_error else None)

    def save(self, context: str) -> StatusMessage:
        self.sap.press(ids.BTN_SAVE)
        self.sap.handle_popups(confirm=True, context=f"salvare {context}")
        status = self.sap.raise_on_error(f"Salvarea ({context})")
        log.info("Salvat (%s). SAP: %s", context, status)
        return status

    # -- activități ------------------------------------------------------------ #
    def open_activities(self, wbs_row: int) -> None:
        """Deschide activitățile WBS-ului selectat (secvența din înregistrare)."""
        sap = self.sap
        self._wbs_table().select_row(wbs_row, True)
        if sap.exists(ids.TAB_WBS_LETR):
            sap.select(ids.TAB_WBS_LETR)
        sap.select(ids.TAB_ACTIVITIES)
        sap.handle_popups(values=self.cfg.sap.network_popup_values, context="activități / antet rețea")
        sap.raise_on_error("Deschiderea activităților")
        if sap.exists(ids.TAB_ACT_EXTERNAL):
            sap.select(ids.TAB_ACT_EXTERNAL)

    def read_activities(self) -> List[ActivityRow]:
        if not self.sap.exists(ids.ACT_TABLE):
            if self.sap.exists(ids.TAB_ACT_EXTERNAL):
                self.sap.select(ids.TAB_ACT_EXTERNAL)
            else:
                return []
        table = SapTable(self.sap, ids.ACT_TABLE)
        return [ActivityRow(i, vornr, text) for i, (vornr, text) in
                enumerate(table.read_rows([ids.ACT_COL_VORNR, ids.ACT_COL_TEXT], key=ids.ACT_COL_VORNR))]

    def copy_standard_network(self, network: str) -> None:
        """Meniu -> copiere rețea standard (STDNR = rețeaua, STDAL = alternativa, ex. QS)."""
        sap = self.sap
        sap.select(ids.MENU_COPY_STD_NETWORK)
        sap.set_text(ids.STD_NETWORK_NR, network)
        sap.set_text(ids.STD_NETWORK_ALT, self.cfg.sap.std_network_alternative)
        sap.press(ids.POPUP_OK)
        sap.handle_popups(values=self.cfg.sap.network_popup_values, context=f"copiere rețea {network}")
        sap.raise_on_error(f"Copierea rețelei standard {network}")
        if sap.exists(ids.TAB_ACT_EXTERNAL):
            sap.select(ids.TAB_ACT_EXTERNAL)
        log.info("Rețeaua standard %s/%s copiată", network, self.cfg.sap.std_network_alternative)

    def delete_activities(self, rows: Sequence[int]) -> None:
        """Șterge activitățile nefolosite, pe loturi vizibile (indicii se recalculează)."""
        pending = sorted(set(rows))
        table = SapTable(self.sap, ids.ACT_TABLE)
        while pending:
            table.scroll_to(pending[0])
            visible = set(table.visible_rows())
            batch = [r for r in pending if r in visible]
            for row in batch:
                table.select_row(row, True)
            self.sap.press(ids.ACT_DELETE_BTN)
            self.sap.handle_popups(confirm=True, context="ștergere activități")
            self.sap.raise_on_error("Ștergerea activităților")
            pending = [r - sum(1 for b in batch if b < r) for r in pending if r not in batch]
        if rows:
            log.info("%d activități nefolosite șterse", len(set(rows)))

    def open_activity_detail(self, row: int) -> None:
        table = SapTable(self.sap, ids.ACT_TABLE)
        table.select_row(row, True)
        table.focus(row, ids.ACT_COL_VORNR)
        self.sap.send_vkey(2)  # F2 = detaliu pe rândul cu cursorul
        if not self.sap.exists(ids.ACT_DETAIL_SERVICES_BTN):
            text = self.sap.popup_text(self.sap.top_window())
            self.sap.close_all_popups()
            raise SapActionError(f"Detaliul activității nu s-a deschis: {text}")

    def deselect_activity(self, row: int) -> None:
        try:
            SapTable(self.sap, ids.ACT_TABLE).select_row(row, False)
        except Exception:  # noqa: BLE001 - doar cosmetic
            pass

    def fill_activity_detail(self, text: str) -> None:
        """Textul activității + grup de cumpărare (EKGRP) + solicitant (AFNAM)."""
        sap = self.sap
        if sap.exists(ids.ACT_DETAIL_TEXT):
            sap.set_text(ids.ACT_DETAIL_TEXT, text)
        else:
            log.warning("Câmpul de text al activității nu există în detaliu; textul rămâne neschimbat")
        if sap.exists(ids.ACT_DETAIL_TAB_EXTERNAL):
            sap.select(ids.ACT_DETAIL_TAB_EXTERNAL)
        for element, value in ((ids.ACT_DETAIL_EKGRP, self.cfg.sap.purchasing_group),
                               (ids.ACT_DETAIL_AFNAM, self.cfg.sap.requisitioner)):
            if value and sap.exists(element):
                sap.set_text(element, value)

    def open_services(self) -> None:
        sap = self.sap
        sap.press(ids.ACT_DETAIL_SERVICES_BTN)
        if sap.top_window() >= 2:
            text = sap.popup_text(sap.top_window())
            sap.handle_popups(base_level=1, context="deschidere servicii")
            raise SapActionError(f"Serviciile nu s-au deschis: {text}")
        if not sap.exists(ids.SRV_TABLE):
            sap.handle_popups(context="deschidere servicii")
            if not sap.exists(ids.SRV_TABLE):
                raise SapActionError(f"Ecranul de servicii nu s-a deschis ({sap.status()})")

    def read_services(self) -> List[Tuple[int, str]]:
        table = SapTable(self.sap, ids.SRV_TABLE)
        return [(i, code) for i, (code,) in
                enumerate(table.read_rows([ids.SRV_COL_CODE], key=ids.SRV_COL_CODE))]

    def add_norm_to_activity(self, row: int, code: str, quantity: str, price: Optional[str] = None,
                             unlimited: bool = True) -> StatusMessage:
        """Adaugă un serviciu (normă) pe rândul ``row`` al specificației și îl validează.

        La eroare rândul este golit și se ridică ``SapLineError`` (bucla poate continua).
        """
        sap = self.sap
        table = SapTable(sap, ids.SRV_TABLE)
        table.set_text(row, ids.SRV_COL_CODE, code)
        table.set_text(row, ids.SRV_COL_QTY, quantity)
        if price:
            table.set_text(row, ids.SRV_COL_PRICE, price)
        status = sap.enter()
        popups = sap.handle_popups(context=f"serviciu {code}")
        if not status.is_error:
            status = sap.status() if popups else status
        if status.is_error:
            message = status.text or " | ".join(popups)
            table = SapTable(sap, ids.SRV_TABLE)
            table.set_text(row, ids.SRV_COL_CODE, "")
            table.set_text(row, ids.SRV_COL_QTY, "")
            if price:
                table.set_text(row, ids.SRV_COL_PRICE, "")
            after = sap.enter()
            sap.handle_popups(context=f"anulare serviciu {code}")
            if after.is_error:
                raise SapActionError(f"Rândul cu eroare ({code}) nu a putut fi anulat: {after.text}", after)
            raise SapLineError(message, status)
        if unlimited:
            try:
                SapTable(sap, ids.SRV_TABLE).set_checkbox(row, ids.SRV_COL_UNLIMITED, True)
            except Exception as exc:  # noqa: BLE001
                log.warning("Nu pot bifa 'nelimitat' pentru %s: %s", code, exc)
        return status

    def confirm_services(self) -> StatusMessage:
        status = self.sap.enter()
        self.sap.handle_popups(context="confirmare servicii")
        return self.sap.raise_on_error("Confirmarea serviciilor", status if status.is_error else None)

    def close_services_and_detail(self) -> None:
        """F3 din servicii -> detaliul activității -> închidere (butonul din înregistrare)."""
        sap = self.sap
        sap.press(ids.BTN_BACK)
        sap.handle_popups(base_level=1, confirm=True, context="ieșire servicii")
        if sap.exists(ids.ACT_DETAIL_CLOSE_BTN):
            sap.press(ids.ACT_DETAIL_CLOSE_BTN)
        elif sap.top_window() >= 1:
            sap.press(ids.POPUP_OK)
        sap.handle_popups(context="închidere detaliu activitate")
        sap.raise_on_error("Închiderea detaliului activității")

    def leave_activities(self) -> None:
        self.sap.press(ids.BTN_BACK)
        self.sap.handle_popups(confirm=True, context="ieșire activități")


def plan_wbs_order(existing: Sequence[WbsRow], level3_nodes: Iterable) -> Tuple[List, Set[str]]:
    """Ordinea de procesare a grupurilor de nivel 3, respectând ierarhia din listă.

    Elementele noi se adaugă doar la final; deci un nivel 4 nou poate fi creat
    numai sub ULTIMUL nivel 3 existent, sau sub un nivel 3 nou creat. Copiii noi
    ai altor niveluri 3 existente sunt blocați (raportați, nu creați).
    """
    def exists(code: str) -> bool:
        return find_wbs(existing, code) is not None

    last_l3 = None
    for row in existing:
        if row.level <= 3:
            last_l3 = row
    tail_ident = last_l3.ident if last_l3 is not None and last_l3.level == 3 else None

    without_new, tail, new_groups = [], None, []
    blocked: Set[str] = set()
    for group in level3_nodes:
        if not exists(group.code):
            new_groups.append(group)
            continue
        new_children = [code for code in group.children if not exists(code)]  # cheile = coduri WBS
        if not new_children:
            without_new.append(group)
        elif tail_ident is not None and ident_matches(tail_ident, group.code):
            tail = group
        else:
            blocked.update(new_children)
            without_new.append(group)
    order = without_new + ([tail] if tail is not None else []) + new_groups
    return order, blocked


# =============================================================================
#  11. EXECUȚIA PLANULUI ÎN SAP + TRATAREA ERORILOR
# =============================================================================
#
#  Execuția planului în SAP, cu tratarea erorilor pe trei niveluri.
#
#  * rând de serviciu (normă): eroarea se loghează, rândul se anulează, bucla continuă;
#  * activitate: se revine la lista de activități și se trece la următoarea;
#  * element WBS de nivel 4: se renunță la modificările nesalvate (``/n``) și se
#    trece la următorul. După ``max_consecutive_failures`` eșecuri la rând
#    execuția se oprește (de ex. sesiune SAP închisă).
# =============================================================================

ST_OK, ST_EXISTENT, ST_ERROR, ST_BLOCKED, ST_PENDING, ST_SKIPPED = (
    "OK", "EXISTENT", "EROARE", "BLOCAT", "NEPROCESAT", "SARIT")


@dataclass
class ResultRow:
    level3: str
    wbs: str
    network: str
    activity: str
    code: str
    quantity: str
    sources: str
    status: str = ST_PENDING
    message: str = ""


@dataclass
class RunResults:
    rows: Dict[Tuple[str, str, str], ResultRow] = field(default_factory=dict)
    events: List[Tuple[str, str, str]] = field(default_factory=list)  # (wbs, status, mesaj)
    aborted: bool = False

    def set(self, wbs: str, activity: str, code: str, status: str, message: str = "") -> None:
        row = self.rows.get((wbs, activity, code))
        if row is not None:
            row.status, row.message = status, message

    def set_pending(self, wbs: str, status: str, message: str, activity: Optional[str] = None) -> None:
        for (w, a, _c), row in self.rows.items():
            if w == wbs and (activity is None or a == activity) and row.status == ST_PENDING:
                row.status, row.message = status, message

    def counts(self) -> Dict[str, int]:
        result: Dict[str, int] = {}
        for row in self.rows.values():
            result[row.status] = result.get(row.status, 0) + 1
        return result


class PlanRunner:
    def __init__(self, sap: SapSession, plan: ProjectPlan, cfg: AppConfig, step: bool = False,
                 max_n4: Optional[int] = None, max_consecutive_failures: int = 3):
        self.sap = sap
        self.plan = plan
        self.cfg = cfg
        self.builder = ProjectBuilder(sap, cfg)
        self.step = step
        self.max_n4 = max_n4
        self.max_failures = max_consecutive_failures
        self.results = RunResults()
        self._price_re = re.compile(cfg.sap.price_one_regex) if cfg.sap.price_one_regex else None
        for n3, n4 in plan.iter_wbs4():
            for act in n4.activities.values():
                for line in act.services.values():
                    self.results.rows[(n4.code, act.text, line.code)] = ResultRow(
                        n3.code, n4.code, n4.network, act.text, line.code,
                        self._qty(line), line.source_refs)

    def _qty(self, line: ServiceLine) -> str:
        return format_quantity(line.quantity, self.cfg.sap.quantity_decimals, self.cfg.sap.decimal_separator)

    # ------------------------------------------------------------------ #
    def run(self) -> RunResults:
        project = self.plan.project_definition
        log.info("=== Începe execuția în SAP pentru proiectul %s ===", project)
        self.builder.open_project(project)
        existing = self.builder.read_wbs_rows()
        log.info("Proiectul are %d elemente WBS în listă", len(existing))
        order, blocked = plan_wbs_order(existing, self.plan.level3.values())
        for code in sorted(blocked):
            log.error("WBS %s nu poate fi adăugat: nivelul 3 părinte nu este ultimul din listă "
                      "(creați-l manual sau mutați nivelul 3 la final)", code)
            self.results.set_pending(code, ST_BLOCKED, "părintele de nivel 3 nu este ultimul din listă")

        processed = 0
        failures = 0
        for n3 in order:
            if self.results.aborted:
                break
            try:
                self._ensure_level3(n3)
            except Exception as exc:  # noqa: BLE001
                log.error("Eroare la WBS %s: %s", n3.code, exc)
                for child in n3.children.values():
                    self.results.set_pending(child.code, ST_ERROR, f"nivelul 3 {n3.code}: {exc}")
                self._recover(n3.code)
                failures += 1
                if failures >= self.max_failures:
                    self._abort()
                continue
            for n4 in n3.children.values():
                if n4.code in blocked:
                    continue
                if self.max_n4 is not None and processed >= self.max_n4:
                    log.info("Limita --max-n4=%d atinsă; restul rămâne neprocesat", self.max_n4)
                    self.results.aborted = True
                    break
                if self.step and not self._confirm(n4):
                    self.results.set_pending(n4.code, ST_SKIPPED, "sărit de utilizator")
                    continue
                processed += 1
                try:
                    self._process_wbs4(n3, n4)
                    failures = 0
                except Exception as exc:  # noqa: BLE001
                    log.error("Eroare la WBS %s: %s", n4.code, exc)
                    self.results.set_pending(n4.code, ST_ERROR, str(exc))
                    self._recover(n4.code)
                    failures += 1
                    if failures >= self.max_failures:
                        self._abort()
                        break
        log.info("=== Execuție terminată: %s ===", self.results.counts())
        return self.results

    # ------------------------------------------------------------------ #
    def _abort(self) -> None:
        log.critical("%d eșecuri consecutive: opresc execuția (verificați sesiunea SAP)", self.max_failures)
        self.results.aborted = True

    def _confirm(self, n4: Wbs4Plan) -> bool:
        answer = input(f"Procesez {n4.code} '{n4.description}' ({n4.service_count()} servicii)? "
                       "[D]a / [N]u / [S]top: ").strip().lower()
        if answer.startswith("s"):
            self.results.aborted = True
            raise KeyboardInterrupt("oprit de utilizator")
        return not answer.startswith("n")

    def _recover(self, context: str) -> None:
        log.warning("Recuperare după eroare (%s): modificările nesalvate se abandonează", context)
        try:
            self.sap.close_all_popups()
            self.sap.start_transaction(self.cfg.sap.transaction)
        except Exception as exc:  # noqa: BLE001
            log.error("Recuperarea a eșuat: %s", exc)

    def _ensure_level3(self, n3: Wbs3Plan) -> None:
        builder = self.builder
        builder.open_project(self.plan.project_definition)
        if find_wbs(builder.read_wbs_rows(), n3.code):
            log.info("WBS nivel 3 %s există deja", n3.code)
            return
        row = builder.create_wbs_element(None, n3.code, 3, n3.description)
        builder.set_assignments(row, n3.assignments)
        builder.set_custom_fields(row, n3.custom_fields)
        builder.save(f"WBS {n3.code}")
        self.results.events.append((n3.code, ST_OK, "nivel 3 creat"))

    def _process_wbs4(self, n3: Wbs3Plan, n4: Wbs4Plan) -> None:
        builder = self.builder
        log.info("--- WBS %s '%s' [%s] ---", n4.code, n4.description, n4.network)
        builder.open_project(self.plan.project_definition)
        rows = builder.read_wbs_rows()
        existing = find_wbs(rows, n4.code)
        if existing is None:
            row = builder.create_wbs_element(n3.code, n4.code, 4, n4.description)
            builder.set_assignments(row, n4.assignments)
            builder.set_custom_fields(row, n4.custom_fields)
            builder.save(f"WBS {n4.code}")
            self.results.events.append((n4.code, ST_OK, "nivel 4 creat"))
            builder.open_project(self.plan.project_definition)
            existing = find_wbs(builder.read_wbs_rows(), n4.code)
            if existing is None:
                raise SapError(f"WBS {n4.code} nu se găsește după salvare")
        else:
            log.info("WBS %s există deja; completez doar activitățile/serviciile lipsă", n4.code)

        builder.open_activities(existing.row)
        current = builder.read_activities()
        reuse: List[Tuple[ActivityPlan, ActivityRow]] = []
        todo: List[ActivityPlan] = []
        for act in n4.activities.values():
            match = next((a for a in current if norm(a.text) == norm(act.text)), None)
            if match:
                reuse.append((act, match))
            else:
                todo.append(act)

        for act, act_row in reuse:
            log.info("Activitatea '%s' există deja (op. %s)", act.text, act_row.vornr)
            self._process_activity(n4, act, act_row.row)

        for round_acts in plan_copy_rounds(todo):
            before = {a.vornr for a in builder.read_activities()}
            builder.copy_standard_network(n4.network)
            new_rows = [a for a in builder.read_activities() if a.vornr not in before]
            keep, delete, missing = choose_rows_to_keep(
                new_rows, [a.key for a in round_acts], self.cfg.rules.activity_key_regex)
            for key in missing:
                for act in round_acts:
                    if act.key == key:
                        msg = f"activitatea {key} nu există în rețeaua standard {n4.network}"
                        log.error("Eroare la WBS %s: %s ('%s')", n4.code, msg, act.text)
                        self.results.set_pending(n4.code, ST_ERROR, msg, act.text)
            builder.delete_activities([r.row for r in delete])
            by_vornr = {a.vornr: a.row for a in builder.read_activities()}
            for act in round_acts:
                kept = keep.get(act.key)
                if kept is None:
                    continue
                if kept.vornr not in by_vornr:
                    self.results.set_pending(n4.code, ST_ERROR, "activitatea a dispărut după ștergere", act.text)
                    continue
                self._process_activity(n4, act, by_vornr[kept.vornr])

        builder.leave_activities()
        builder.save(f"activități WBS {n4.code}")

    def _process_activity(self, n4: Wbs4Plan, act: ActivityPlan, row: int) -> None:
        builder = self.builder
        try:
            builder.open_activity_detail(row)
            builder.fill_activity_detail(act.text)
            builder.open_services()
            existing_services = builder.read_services()
            existing_codes = {code for _, code in existing_services}
            next_row = len(existing_services)
            for line in act.services.values():
                if line.code in existing_codes:
                    self.results.set(n4.code, act.text, line.code, ST_EXISTENT, "serviciul exista deja")
                    log.info("WBS %s / %s: serviciul %s există deja, sărit", n4.code, act.text, line.code)
                    continue
                qty = self._qty(line)
                price = "1" if self._price_re and self._price_re.match(line.code) else None
                try:
                    status = builder.add_norm_to_activity(
                        next_row, line.code, qty, price, self.cfg.sap.set_unlimited_overfulfillment)
                    next_row += 1
                    self.results.set(n4.code, act.text, line.code, ST_OK, status.text)
                    log.info("WBS %s / %s: %s = %s OK", n4.code, act.text, line.code, qty)
                except SapLineError as exc:
                    log.error("Eroare la WBS %s, rândul %s (norma %s, activitatea '%s'): %s",
                              n4.code, line.source_refs, line.code, act.text, exc)
                    self.results.set(n4.code, act.text, line.code, ST_ERROR, str(exc))
            builder.confirm_services()
            builder.close_services_and_detail()
            builder.deselect_activity(row)
        except SapError as exc:
            log.error("Eroare la WBS %s, activitatea '%s': %s", n4.code, act.text, exc)
            self.results.set_pending(n4.code, ST_ERROR, str(exc), act.text)
            self._back_to_activity_list()

    def _back_to_activity_list(self) -> None:
        """Revenire la lista de activități după o eroare într-o activitate."""
        sap = self.sap
        sap.close_all_popups()
        if sap.exists(ids.SRV_TABLE):
            sap.press(ids.BTN_BACK)
            sap.handle_popups(base_level=1, confirm=True, context="recuperare servicii")
            sap.close_all_popups()
        if not sap.exists(ids.ACT_TABLE) and not sap.exists(ids.TAB_ACT_EXTERNAL):
            raise SapError("Nu am putut reveni la lista de activități")


# =============================================================================
#  12. RECALCULARE FORMULE PRIN EXCEL (opțional)
# =============================================================================
#
#  Recalcularea formulelor prin Excel (Windows), când fișierul nu are valori salvate.
#
#  openpyxl citește doar valorile salvate de Excel; dacă fișierul a fost generat
#  de alt program, formulele nu au valori. Atunci Excel deschide o copie, o
#  recalculează și o salvează într-un fișier temporar care va fi citit.
# =============================================================================

MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


def recalc_with_excel(path: Path) -> Path:
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("Recalcularea necesită Excel + pywin32 (Windows)") from exc
    source = Path(path).resolve()
    target = Path(tempfile.mkdtemp(prefix="sapwbs_")) / f"{source.stem}_recalculat{source.suffix}"
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE  # fără macro-uri la deschidere
    try:
        workbook = excel.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=True)
        try:
            excel.CalculateFull()
            workbook.SaveCopyAs(str(target))
        finally:
            workbook.Close(SaveChanges=False)
    finally:
        excel.Quit()
    log.info("Devizul a fost recalculat în Excel: %s", target)
    return target


# =============================================================================
#  13. LOGGING (fișier + consolă)
# =============================================================================
#
#  Configurarea modulului logging: fișier + consolă, cu timestamp.
# =============================================================================

FORMAT = "%(asctime)s | %(levelname)-8s | %(funcName)-26s | %(message)s"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"


def setup_logging(log_dir: Path, level: str = "INFO", run_id: str = "") -> Path:
    """Creează fișierul de log al rulării și trimite mesajele și în consolă."""
    log_dir.mkdir(parents=True, exist_ok=True)
    run_id = run_id or datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"sap_wbs_{run_id}.log"

    root = logging.getLogger()
    root.setLevel(logging.DEBUG)
    for handler in list(root.handlers):
        root.removeHandler(handler)

    formatter = logging.Formatter(FORMAT, DATE_FORMAT)
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    root.addHandler(file_handler)

    console = logging.StreamHandler(sys.stdout)
    console.setLevel(getattr(logging, level.upper(), logging.INFO))
    console.setFormatter(formatter)
    root.addHandler(console)

    for noisy in ("openpyxl",):
        logging.getLogger(noisy).setLevel(logging.WARNING)
    return log_file


# =============================================================================
#  14. LINIE DE COMANDĂ
# =============================================================================
#
#  Punctul de intrare în linie de comandă.
#
#  Exemple:
#      python run.py deviz.xlsm --dry-run
#      python run.py deviz.xlsm --max-n4 1 --pas-cu-pas
#      python run.py deviz.xlsm --proiect C-12345 --cod-buget REM_001604
# =============================================================================

DEFAULT_CONFIG = Path(__file__).resolve().parent / "config.yaml"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="sap_wbs.py",
        description="Creează WBS nivel 3/4 în SAP și încarcă normele din devizul general.")
    parser.add_argument("excel", type=Path, help="devizul (.xlsm / .xlsx)")
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG,
                        help="fișier YAML opțional care suprascrie configurarea din script")
    parser.add_argument("--dry-run", action="store_true",
                        help="doar citește devizul și generează raportul pre-flight (fără SAP)")
    parser.add_argument("--proiect", default="", help="definiția de proiect SAP (suprascrie Excel-ul)")
    parser.add_argument("--cod-buget", default="", help="codul de buget (suprascrie Excel-ul)")
    parser.add_argument("--recalculeaza", action="store_true",
                        help="recalculează formulele în Excel înainte de citire (Windows)")
    parser.add_argument("--max-n4", type=int, default=None,
                        help="procesează cel mult N elemente de nivel 4 (pentru teste)")
    parser.add_argument("--pas-cu-pas", action="store_true",
                        help="cere confirmare înainte de fiecare element de nivel 4")
    parser.add_argument("--continua-cu-avertizari", action="store_true",
                        help="rulează în SAP chiar dacă raportul are erori de planificare")
    parser.add_argument("--nivel-log", default=None, help="DEBUG / INFO / WARNING (consolă)")
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    if not args.config.exists() and args.config.resolve() != DEFAULT_CONFIG:
        print(f"Fișierul de configurare nu există: {args.config}", file=sys.stderr)
        return 2
    try:
        cfg = load_config(args.config if args.config.exists() else None)
    except ConfigError as exc:
        print(f"Eroare de configurare: {exc}", file=sys.stderr)
        return 2
    if not args.config.exists():
        cfg.base_dir = Path(__file__).resolve().parent  # logs/ și rapoarte/ lângă script

    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = setup_logging(cfg.resolve_path(cfg.output.log_dir), args.nivel_log or cfg.output.log_level, run_id)
    log.info("Log: %s | Config: %s", log_file, args.config if args.config.exists() else "implicit")

    excel_path = args.excel
    if not excel_path.exists():
        log.error("Fișierul nu există: %s", excel_path)
        return 2
    if args.recalculeaza:

        excel_path = recalc_with_excel(excel_path)

    read = WorkbookReader(excel_path, cfg).read()
    plan = Planner(cfg).build(read, project_override=args.proiect, budget_override=args.cod_buget)
    log.info("Plan dedus din deviz:\n%s", describe_plan(plan, cfg.sap.quantity_decimals, cfg.sap.decimal_separator))
    log.info("Statistici: %s", plan.stats)

    report_dir = cfg.resolve_path(cfg.output.report_dir)
    write_plan_report(plan, report_dir / f"plan_{run_id}.xlsx", read)

    errors = [i for i in plan.issues if i.severity == ERROR]
    if args.dry_run:
        log.info("Dry-run terminat: %d erori, %d avertizări de planificare", len(errors),
                 sum(1 for i in plan.issues if i.severity == "WARNING"))
        return 1 if errors else 0
    if errors and not args.continua_cu_avertizari:
        log.error("Planul are %d erori (vezi raportul). Corectați devizul sau folosiți "
                  "--continua-cu-avertizari. Nu am atins SAP.", len(errors))
        return 1
    if not plan.project_definition:
        log.error("Fără definiția de proiect nu pot deschide tranzacția. Oprire.")
        return 1


    try:
        sap = SapSession.attach(cfg.sap)
    except SapConnectionError as exc:
        log.error("Conexiune SAP: %s", exc)
        return 3
    runner = PlanRunner(sap, plan, cfg, step=args.pas_cu_pas, max_n4=args.max_n4)
    try:
        results = runner.run()
    except KeyboardInterrupt:
        log.warning("Execuție întreruptă de utilizator")
        results = runner.results
    except Exception:  # noqa: BLE001
        log.exception("Eroare neașteptată; rezultatele parțiale se salvează")
        results = runner.results
    write_results_report(results, report_dir / f"rezultate_{run_id}.xlsx")
    counts = results.counts()
    log.info("Rezumat: %s", counts)
    return 0 if not counts.get("EROARE") and not counts.get("BLOCAT") else 1


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
