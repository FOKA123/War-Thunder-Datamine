"""Acțiunile din tranzacția de proiect (CJ02), traduse din înregistrarea VBS.

Fiecare metodă publică este o operație parametrizată, verificată prin bara de
stare și cu tratarea pop-up-urilor. Funcțiile pure de la final (fără SAP) sunt
testate unitar.
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

from . import sap_ids as ids
from .config import AppConfig
from .sap_session import SapActionError, SapLineError, SapSession, SapTable, StatusMessage

log = logging.getLogger(__name__)


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
        for field, value in values.items():
            spec = ids.ASSIGN_COLS.get(field)
            if spec is None:
                raise SapActionError(f"Câmp de atribuire necunoscut: {field}")
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
