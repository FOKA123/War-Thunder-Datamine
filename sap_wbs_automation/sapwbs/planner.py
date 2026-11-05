"""Construirea planului ierarhic: Nivel 3 -> Nivel 4 -> Activitate -> Servicii.

Reguli de alocare pe nivelul 4 (stabilite de noi):

* fiecare rețea standard (ex. A453) are propriile elemente de nivel 4;
* cantitatea din coloana PTk merge pe nivelul 4 al PT-ului k: ``+{rețea}R{k}``
  (ex. ``+441R2`` = clădirea PT2);
* restul (cantitatea totală minus suma PT) = lucrări comune. Dacă rețeaua nu are
  nimic alocat pe PT, lucrările comune primesc indicele 1 (``+452R1``, ca în
  înregistrarea SAP); altfel primesc primul indice liber după ultimul PT din deviz
  (cu 3 PT-uri: ``+453R4``).
"""
from __future__ import annotations

import logging
from collections import Counter
from typing import Dict, List, Optional, Tuple

from .config import AppConfig
from .model import (ERROR, INFO, WARNING, ActivityPlan, Issue, ProjectPlan, ReadResult,
                    SourceRow, Wbs3Plan, Wbs4Plan)
from .rules import (Resolution, RuleEngine, build_activity_text, level3_code, level3_description,
                    level4_code, level4_description)

log = logging.getLogger(__name__)

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
                field = wcfg.assignment_field_by_network.get(node4.network, wcfg.assignment_default_field)
                value = ""
                if node4.pt is not None and node4.pt in ctx.pts:
                    value = ctx.pts[node4.pt].values.get(field, "")
                if not value:
                    value = wcfg.default_equnr if field == "EQUNR" else wcfg.default_tplnr
                if value:
                    node4.assignments[field] = value
                if 4 in wcfg.budget_on_levels and budget:
                    node4.custom_fields["ZZCOD_BUGET"] = budget
                for name, val in wcfg.custom_fields_by_network.get(node4.network, {}).items():
                    if val not in (None, ""):
                        node4.custom_fields[name] = str(val)


def describe_plan(plan: ProjectPlan, decimals: int = 3, decimal_separator: str = ",") -> str:
    """Reprezentare text a arborelui (pentru log / consolă)."""
    from .textutil import format_quantity

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
