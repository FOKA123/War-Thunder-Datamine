"""Execuția planului în SAP, cu tratarea erorilor pe trei niveluri.

* rând de serviciu (normă): eroarea se loghează, rândul se anulează, bucla continuă;
* activitate: se revine la lista de activități și se trece la următoarea;
* element WBS de nivel 4: se renunță la modificările nesalvate (``/n``) și se
  trece la următorul. După ``max_consecutive_failures`` eșecuri la rând
  execuția se oprește (de ex. sesiune SAP închisă).
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from . import sap_ids as ids
from .cj02 import (ActivityRow, ProjectBuilder, choose_rows_to_keep, find_wbs, plan_copy_rounds,
                   plan_wbs_order)
from .config import AppConfig
from .model import ActivityPlan, ProjectPlan, ServiceLine, Wbs3Plan, Wbs4Plan
from .sap_session import SapError, SapLineError, SapSession
from .textutil import format_quantity, norm

log = logging.getLogger(__name__)

OK, EXISTENT, ERROR, BLOCKED, PENDING, SKIPPED = (
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
    status: str = PENDING
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
            if w == wbs and (activity is None or a == activity) and row.status == PENDING:
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
            self.results.set_pending(code, BLOCKED, "părintele de nivel 3 nu este ultimul din listă")

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
                    self.results.set_pending(child.code, ERROR, f"nivelul 3 {n3.code}: {exc}")
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
                    self.results.set_pending(n4.code, SKIPPED, "sărit de utilizator")
                    continue
                processed += 1
                try:
                    self._process_wbs4(n3, n4)
                    failures = 0
                except Exception as exc:  # noqa: BLE001
                    log.error("Eroare la WBS %s: %s", n4.code, exc)
                    self.results.set_pending(n4.code, ERROR, str(exc))
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
        self.results.events.append((n3.code, OK, "nivel 3 creat"))

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
            self.results.events.append((n4.code, OK, "nivel 4 creat"))
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
                        self.results.set_pending(n4.code, ERROR, msg, act.text)
            builder.delete_activities([r.row for r in delete])
            by_vornr = {a.vornr: a.row for a in builder.read_activities()}
            for act in round_acts:
                kept = keep.get(act.key)
                if kept is None:
                    continue
                if kept.vornr not in by_vornr:
                    self.results.set_pending(n4.code, ERROR, "activitatea a dispărut după ștergere", act.text)
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
                    self.results.set(n4.code, act.text, line.code, EXISTENT, "serviciul exista deja")
                    log.info("WBS %s / %s: serviciul %s există deja, sărit", n4.code, act.text, line.code)
                    continue
                qty = self._qty(line)
                price = "1" if self._price_re and self._price_re.match(line.code) else None
                try:
                    status = builder.add_norm_to_activity(
                        next_row, line.code, qty, price, self.cfg.sap.set_unlimited_overfulfillment)
                    next_row += 1
                    self.results.set(n4.code, act.text, line.code, OK, status.text)
                    log.info("WBS %s / %s: %s = %s OK", n4.code, act.text, line.code, qty)
                except SapLineError as exc:
                    log.error("Eroare la WBS %s, rândul %s (norma %s, activitatea '%s'): %s",
                              n4.code, line.source_refs, line.code, act.text, exc)
                    self.results.set(n4.code, act.text, line.code, ERROR, str(exc))
            builder.confirm_services()
            builder.close_services_and_detail()
            builder.deselect_activity(row)
        except SapError as exc:
            log.error("Eroare la WBS %s, activitatea '%s': %s", n4.code, act.text, exc)
            self.results.set_pending(n4.code, ERROR, str(exc), act.text)
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
