"""Conexiunea la SAP GUI Scripting și operațiile generice pe ecrane.

* ``SapSession.attach`` -> GetObject("SAPGUI") -> GetScriptingEngine -> sesiunea potrivită;
* ``status()`` citește bara de stare (``wnd[0]/sbar``) după fiecare acțiune critică;
* ``handle_popups()`` tratează ferestrele modale neașteptate (informări,
  confirmări, cereri de valori), cu log pentru fiecare;
* ``SapTable`` citește/scrie tabelele SAP (table control) cu derulare automată.

Modulul nu importă ``win32com`` decât la conectare, deci restul aplicației
(citire Excel, plan, teste) funcționează și fără SAP / Windows.
"""
from __future__ import annotations

import logging
import time
from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from . import sap_ids as ids
from .config import SapConfig

log = logging.getLogger(__name__)


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
