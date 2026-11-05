"""Configurarea modulului logging: fișier + consolă, cu timestamp."""
from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path

FORMAT = "%(asctime)s | %(levelname)-8s | %(name)-18s | %(message)s"
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
