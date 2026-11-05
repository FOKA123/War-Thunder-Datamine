"""Un SAP GUI simulat (foarte simplificat) pentru testarea logicii de execuție.

Reproduce doar comportamentul folosit de ``ProjectBuilder`` / ``PlanRunner``,
cu aceleași ID-uri ca înregistrarea VBS: ecran inițial CJ02, lista WBS,
atribuiri, detaliu WBS, activități, copiere rețea standard, detaliu activitate
(fereastră modală) și specificația de servicii.
"""
from __future__ import annotations

import re
from typing import Dict, List, Optional

from sap_wbs import ids


class _Collection:
    def __init__(self, items):
        self._items = list(items)

    @property
    def Count(self):
        return len(self._items)

    def __call__(self, index):
        return self._items[index]


class _Element:
    def __init__(self, sap: "FakeSap", element_id: str, type_: str = "GuiButton", text: str = "",
                 on_press=None, on_select=None, getter=None, setter=None, children=None):
        self._sap = sap
        self.Id = element_id
        self.Type = type_
        self._text = text
        self._on_press = on_press
        self._on_select = on_select
        self._getter = getter
        self._setter = setter
        self.Children = _Collection(children or [])
        self.Changeable = True

    @property
    def Text(self):
        return self._getter() if self._getter else self._text

    @Text.setter
    def Text(self, value):
        if self._setter:
            self._setter(value)
        else:
            self._text = value

    def press(self):
        self._sap._reset_status()
        self._on_press()

    def select(self):
        self._sap._reset_status()
        if self._on_select:
            self._on_select()

    def sendVKey(self, key):
        self._sap._reset_status()
        self._sap._vkey(key)

    def maximize(self):
        pass

    def close(self):
        self._sap.popups.pop()


class _Scrollbar:
    def __init__(self, table):
        self._table = table
        self.Position = 0

    @property
    def Maximum(self):
        return max(0, self._table.RowCount - 1)


class _Cell:
    def __init__(self, table, row, col):
        self._t, self._row, self._col = table, row, col

    @property
    def Name(self):
        return self._t.col_names.get(self._col, f"COL{self._col}")

    @property
    def Text(self):
        rows = self._t.rows()
        if self._row < len(rows):
            return str(rows[self._row].get(self.Name, ""))
        return ""

    @Text.setter
    def Text(self, value):
        rows = self._t.rows()
        while len(rows) <= self._row:
            self._t.new_row()
        rows[self._row][self.Name] = value

    @property
    def Selected(self):
        rows = self._t.rows()
        return bool(rows[self._row].get(self.Name)) if self._row < len(rows) else False

    @Selected.setter
    def Selected(self, value):
        self._t.rows()[self._row][self.Name] = bool(value)

    def SetFocus(self):
        self._t.sap.focus_row = self._row


class _Row:
    def __init__(self, table, row):
        self._t, self._row = table, row

    @property
    def Selected(self):
        return self._row in self._t.selected

    @Selected.setter
    def Selected(self, value):
        (self._t.selected.add if value else self._t.selected.discard)(self._row)


class _Table:
    def __init__(self, sap, col_names: Dict[int, str], rows_fn, new_row_fn, visible=4):
        self.sap = sap
        self.col_names = col_names
        self.rows = rows_fn
        self.new_row = new_row_fn
        self.VisibleRowCount = visible
        self.VerticalScrollbar = _Scrollbar(self)
        self.selected = set()
        self.Columns = _Collection(range(max(col_names) + 1))

    @property
    def RowCount(self):
        return len(self.rows()) + self.VisibleRowCount

    def GetCell(self, vis, col):
        return _Cell(self, self.VerticalScrollbar.Position + vis, col)

    def GetAbsoluteRow(self, row):
        return _Row(self, row)


class FakeSap:
    """Sesiune SAP simulată. ``self.wbs`` = lista elementelor WBS (dict-uri)."""

    def __init__(self, project="P-1", wbs: Optional[List[dict]] = None,
                 std_networks: Optional[Dict[str, List[str]]] = None, invalid_services=()):
        self.project = project
        self.wbs = wbs if wbs is not None else [
            {"PRPS-STUFE": "1", "RCWBS-IDENT": project, "PRPS-POST1": "PROIECT"},
            {"PRPS-STUFE": "2", "RCWBS-IDENT": f"{project}-1", "PRPS-POST1": "LUCRARE"},
        ]
        self.std_networks = std_networks or {}
        self.invalid_services = set(invalid_services)
        self.screen = "menu"
        self.popups: List[dict] = []
        self.status = ("", "")
        self.focus_row = 0
        self.saves = 0
        self.current_wbs: Optional[dict] = None
        self.current_act: Optional[dict] = None
        self.customer_tab = False
        self.wbs_tab = "general"
        self.okcode = ""
        self._proj_input = ""
        self.Busy = False
        self.log: List[str] = []
        self._wbs_table = _Table(self, {1: "PRPS-STUFE", 2: "RCWBS-IDENT", 3: "PRPS-POST1"},
                                 lambda: self.wbs, lambda: self.wbs.append({}))
        self._assign_table = _Table(self, {7: "PRPS-STORT", 14: "PRPS-EQUNR", 15: "PRPS-TPLNR"},
                                    lambda: self.wbs, lambda: self.wbs.append({}))
        self._act_table = _Table(self, {0: "AFVGD-VORNR", 2: "AFVGD-LTXA1"},
                                 lambda: self.current_wbs.setdefault("acts", []), lambda: None, visible=5)
        self._srv_table = _Table(self, {2: "ESLL-SRVPOS", 4: "ESLL-MENGE", 6: "ESLL-TBTWR", 8: "ESLL-UEBTK"},
                                 lambda: self.current_act.setdefault("services", []),
                                 lambda: self.current_act["services"].append({}), visible=3)

    # -- API SAP ----------------------------------------------------------- #
    def findById(self, element_id, raise_error=True):
        element = self._elements().get(element_id)
        if element is None and raise_error:
            raise RuntimeError(f"not found: {element_id}")
        return element

    # -- stare --------------------------------------------------------------- #
    def _reset_status(self):
        self.status = ("", "")

    def _set_status(self, type_, text):
        self.status = (type_, text)

    def _push(self, kind, **data):
        self.popups.append({"kind": kind, **data})

    def _elements(self) -> Dict[str, object]:
        e: Dict[str, object] = {}
        add = lambda eid, **kw: e.__setitem__(eid, _Element(self, eid, **kw))  # noqa: E731
        add("wnd[0]", type_="GuiMainWindow")
        e["wnd[0]/sbar"] = type("Bar", (), {"MessageType": self.status[0], "Text": self.status[1],
                                            "MessageId": "", "MessageNumber": ""})()
        add(ids.OKCODE, type_="GuiOkCodeField", getter=lambda: self.okcode,
            setter=lambda v: setattr(self, "okcode", v))
        if self.screen == "initial":
            add(ids.INITIAL_PROJECT_FIELD, type_="GuiCTextField", getter=lambda: self._proj_input,
                setter=lambda v: setattr(self, "_proj_input", v))
            add(ids.INITIAL_WBS_FIELD, type_="GuiCTextField", setter=lambda v: None)
            add(ids.BTN_WBS_OVERVIEW, on_press=self._open_overview)
        elif self.screen == "overview":
            add(ids.TAB_WBS_GENERAL, type_="GuiTab", on_select=lambda: setattr(self, "wbs_tab", "general"))
            add(ids.TAB_WBS_ASSIGN, type_="GuiTab", on_select=lambda: setattr(self, "wbs_tab", "assign"))
            add(ids.TAB_WBS_LETR, type_="GuiTab", on_select=lambda: None)
            add(ids.TAB_ACTIVITIES, type_="GuiTab", on_select=self._open_activities)
            add(ids.MENU_WBS_DETAIL, type_="GuiMenu", on_select=self._open_wbs_detail)
            add(ids.BTN_SAVE, on_press=self._save)
            if self.wbs_tab == "general":
                e[ids.WBS_TABLE] = self._wbs_table
            else:
                e[ids.WBS_ASSIGN_TABLE] = self._assign_table
        elif self.screen == "wbs_detail":
            add(ids.BTN_WBS_CUSTOMER_TAB, on_press=lambda: setattr(self, "customer_tab", True))
            add(ids.BTN_SAVE, on_press=self._save)
            if self.customer_tab:
                for name, eid in ids.CUSTOM_FIELDS.items():
                    add(eid, type_="GuiCTextField",
                        getter=lambda n=name: self.current_wbs.get(n, ""),
                        setter=lambda v, n=name: self.current_wbs.__setitem__(n, v))
        elif self.screen == "activities":
            add(ids.MENU_COPY_STD_NETWORK, type_="GuiMenu", on_select=lambda: self._push("stdnet"))
            add(ids.TAB_ACT_EXTERNAL, type_="GuiTab", on_select=lambda: None)
            add(ids.ACT_DELETE_BTN, on_press=self._delete_request)
            add(ids.BTN_BACK, on_press=lambda: setattr(self, "screen", "overview"))
            add(ids.BTN_SAVE, on_press=self._save)
            e[ids.ACT_TABLE] = self._act_table
        elif self.screen == "services":
            add(ids.BTN_BACK, on_press=self._leave_services)
            e[ids.SRV_TABLE] = self._srv_table
        self._popup_elements(e, add)
        return e

    def _popup_elements(self, e, add):
        for level, popup in enumerate(self.popups, start=1):
            w = f"wnd[{level}]"
            kind = popup["kind"]
            label = _Element(self, f"{w}/usr/lbl", type_="GuiLabel", text=popup.get("text", kind))
            add(w, type_="GuiModalWindow", text=popup.get("title", kind))
            add(f"{w}/usr", type_="GuiUserArea", children=[label])
            if kind == "spo4":
                field = _Element(self, f"{w}/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]",
                                 type_="GuiCTextField", getter=lambda p=popup: p.get("value", ""),
                                 setter=lambda v, p=popup: p.__setitem__("value", v))
                add(f"{w}/usr/sub:SAPLSPO4:0300", type_="GuiSimpleContainer", children=[field])
                add(f"{w}/tbar[0]/btn[0]", on_press=self._answer_spo4)
            elif kind == "stdnet":
                add(ids.STD_NETWORK_NR, type_="GuiCTextField", setter=lambda v, p=popup: p.__setitem__("nr", v))
                add(ids.STD_NETWORK_ALT, type_="GuiTextField", setter=lambda v, p=popup: p.__setitem__("alt", v))
                add(ids.POPUP_OK, on_press=self._copy_network)
            elif kind == "confirm":
                add(f"{w}/{ids.POPUP_YES}", on_press=lambda: self._confirm(True))
                add(f"{w}/{ids.POPUP_NO}", on_press=lambda: self._confirm(False))
            elif kind == "detail":
                act = popup["act"]
                add(ids.ACT_DETAIL_TEXT, type_="GuiTextField", getter=lambda: act["AFVGD-LTXA1"],
                    setter=lambda v: act.__setitem__("AFVGD-LTXA1", v))
                add(ids.ACT_DETAIL_TAB_EXTERNAL, type_="GuiTab", on_select=lambda: None)
                add(ids.ACT_DETAIL_EKGRP, type_="GuiCTextField", setter=lambda v: act.__setitem__("EKGRP", v))
                add(ids.ACT_DETAIL_AFNAM, type_="GuiTextField", setter=lambda v: act.__setitem__("AFNAM", v))
                add(ids.ACT_DETAIL_SERVICES_BTN, on_press=lambda: self._open_services(act))
                add(ids.ACT_DETAIL_CLOSE_BTN, on_press=lambda: self.popups.pop())
            else:  # info
                add(f"{w}/tbar[0]/btn[0]", on_press=lambda: self.popups.pop())

    # -- comportament -------------------------------------------------------- #
    def _vkey(self, key):
        if key == 0 and self.okcode.upper().startswith("/N"):
            self.okcode = ""
            self.popups.clear()
            self.screen = "initial"
            self._proj_input = ""
            return
        if key == 2 and self.screen == "activities":
            acts = self.current_wbs.get("acts", [])
            if self.focus_row < len(acts):
                self._push("detail", act=acts[self.focus_row])
            return
        if key == 0 and self.screen == "services":
            self._validate_services()

    def _open_overview(self):
        if self._proj_input != self.project:
            self._set_status("E", f"Proiectul {self._proj_input} nu există")
            return
        self.screen, self.wbs_tab = "overview", "general"
        self._wbs_table.selected.clear()
        self._assign_table.selected.clear()

    def _selected_wbs(self) -> Optional[dict]:
        table = self._wbs_table if self.wbs_tab == "general" else self._assign_table
        rows = sorted(table.selected)
        return self.wbs[rows[0]] if rows else None

    def _open_wbs_detail(self):
        self.current_wbs = self._selected_wbs()
        self.screen, self.customer_tab = "wbs_detail", False

    def _open_activities(self):
        self.current_wbs = self._selected_wbs()
        self.screen = "activities"
        self._act_table.selected.clear()
        self._act_table.VerticalScrollbar.Position = 0
        if not self.current_wbs.get("network"):
            self.current_wbs["network"] = True
            self._push("spo4", title="Valori", text="Introduceți valoarea")
            self._push("spo4", title="Valori", text="Introduceți valoarea")

    def _answer_spo4(self):
        popup = self.popups.pop()
        self.current_wbs.setdefault("spo4_values", []).append(popup.get("value"))

    def _copy_network(self):
        popup = self.popups.pop()
        template = self.std_networks.get(popup.get("nr"))
        if template is None:
            self._set_status("E", f"Rețeaua standard {popup.get('nr')} nu există")
            return
        acts = self.current_wbs.setdefault("acts", [])
        last = max((int(a["AFVGD-VORNR"]) for a in acts), default=0)
        for i, text in enumerate(template, start=1):
            acts.append({"AFVGD-VORNR": f"{last + 10 * i:04d}", "AFVGD-LTXA1": text})
        self._set_status("S", "Rețea copiată")

    def _delete_request(self):
        if self._act_table.selected:
            self._push("confirm", title="Ștergere", text="Ștergeți activitățile?")

    def _confirm(self, yes):
        self.popups.pop()
        if yes:
            acts = self.current_wbs["acts"]
            for idx in sorted(self._act_table.selected, reverse=True):
                del acts[idx]
        self._act_table.selected.clear()

    def _open_services(self, act):
        if not act.get("EKGRP"):
            self._push("info", title="Eroare", text="Introduceți grupul de cumpărare")
            return
        self.popups.pop()
        self.current_act = act
        self._srv_table.VerticalScrollbar.Position = 0
        self.screen = "services"

    def _validate_services(self):
        services = self.current_act.setdefault("services", [])
        services[:] = [s for s in services if s.get("ESLL-SRVPOS") or s.get("ESLL-MENGE")]
        for srv in services:
            if srv.get("ok"):
                continue
            code = srv.get("ESLL-SRVPOS", "")
            if code in self.invalid_services:
                self._set_status("E", f"Serviciul {code} nu există")
                return
            if not re.match(r"^\d+(,\d+)?$", srv.get("ESLL-MENGE", "")):
                self._set_status("E", f"Cantitate invalidă pentru {code}")
                return
            srv["ok"] = True

    def _leave_services(self):
        self.screen = "activities"
        self._push("detail", act=self.current_act)

    def _save(self):
        self.saves += 1
        self.screen = "initial"
        self._proj_input = ""
        self._set_status("S", "Proiectul a fost salvat")
