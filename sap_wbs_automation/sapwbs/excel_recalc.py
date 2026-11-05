"""Recalcularea formulelor prin Excel (Windows), când fișierul nu are valori salvate.

openpyxl citește doar valorile salvate de Excel; dacă fișierul a fost generat
de alt program, formulele nu au valori. Atunci Excel deschide o copie, o
recalculează și o salvează într-un fișier temporar care va fi citit.
"""
from __future__ import annotations

import logging
import tempfile
from pathlib import Path

log = logging.getLogger(__name__)

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
