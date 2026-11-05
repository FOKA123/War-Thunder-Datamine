"""Punctul de intrare în linie de comandă.

Exemple:
    python run.py deviz.xlsm --dry-run
    python run.py deviz.xlsm --max-n4 1 --pas-cu-pas
    python run.py deviz.xlsm --proiect C-12345 --cod-buget REM_001604
"""
from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from .config import ConfigError, load_config
from .excel_reader import WorkbookReader
from .logging_setup import setup_logging
from .model import ERROR
from .planner import Planner, describe_plan
from .report import write_plan_report, write_results_report

log = logging.getLogger("sapwbs")

DEFAULT_CONFIG = Path(__file__).resolve().parent.parent / "config.yaml"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="sapwbs",
        description="Creează WBS nivel 3/4 în SAP și încarcă normele din devizul general.")
    parser.add_argument("excel", type=Path, help="devizul (.xlsm / .xlsx)")
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG, help="fișierul YAML de configurare")
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

    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = setup_logging(cfg.resolve_path(cfg.output.log_dir), args.nivel_log or cfg.output.log_level, run_id)
    log.info("Log: %s | Config: %s", log_file, args.config if args.config.exists() else "implicit")

    excel_path = args.excel
    if not excel_path.exists():
        log.error("Fișierul nu există: %s", excel_path)
        return 2
    if args.recalculeaza:
        from .excel_recalc import recalc_with_excel

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

    from .runner import PlanRunner
    from .sap_session import SapConnectionError, SapSession

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
