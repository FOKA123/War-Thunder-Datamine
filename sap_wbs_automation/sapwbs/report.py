"""Rapoartele Excel: planul (pre-flight) și rezultatele execuției."""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

from .model import ProjectPlan, ReadResult

log = logging.getLogger(__name__)


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
