"""Fixturi comune: un deviz sintetic cu structură asemănătoare celui real.

Antetele sunt intenționat pe rânduri și coloane diferite între foi, cu
denumiri diferite ("Cod Norma" / "Cod normă", "Cant." / "Cantitate totala"),
ca să verificăm că nimic nu e indexat fix.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def build_workbook(path: Path, with_budget: bool = True) -> Path:
    import openpyxl
    from openpyxl.workbook.defined_name import DefinedName

    wb = openpyxl.Workbook()
    date = wb.active
    date.title = "Date"
    date["B3"], date["C3"] = "Nivel Tensiune", "MT"
    date["B14"], date["C14"] = "Solutie proiectata", "LES"
    date["A21"], date["B21"], date["C21"] = "PT1", "Solutie proiectata", "PTAB"
    date["D21"], date["E21"], date["F21"], date["G21"] = "Cod Post", "DA102028076", "Cod Linie", "DA1033617"
    date["A25"], date["D25"], date["E25"] = "PT2", "Cod Post", "DA102028077"
    if with_budget:
        date["B8"], date["C8"] = "Cod buget", "REM_001604"
    date["B9"], date["C9"] = "Locatie SAP", "DA1C111"
    date["H40"] = "C-0001"
    if hasattr(wb.defined_names, "add"):
        wb.defined_names.add(DefinedName("COD_PROIECT", attr_text="Date!$H$40"))
    else:  # openpyxl >= 3.1
        wb.defined_names["COD_PROIECT"] = DefinedName("COD_PROIECT", attr_text="Date!$H$40")
    date["P4"], date["Q4"] = "Activitate SAP", "Nivel 3"
    for i, (net, n3) in enumerate([("A412", "M1"), ("A441", "M1"), ("A451", "M1"), ("A452", "M1"),
                                   ("A453", "M1"), ("A463", "B1"), ("A61F", "T1")], start=5):
        date[f"P{i}"], date[f"Q{i}"] = net, n3

    norme = wb.create_sheet("Norme MT")
    norme["A3"] = "Prestatie Terti"
    header = ["Partea", "Cod Norma", "Descriere", "UM", "T.I.", "Cod Sap", "Puncte", "Cant.", "PT1", "PT2"]
    for col, value in enumerate(header, start=1):
        norme.cell(row=17, column=col, value=value)
    norme.cell(row=17, column=20, value="Activitate SAP")
    norme.cell(row=17, column=21, value="Denumire activitate SAP")
    data = [
        # partea, cod, descriere, cant, pt1, pt2, retea, activitate
        ("L", "L22101", "Pozare cablu MT subteran", 348.6, None, None, "A453", "XL4 Pozare LES MT KM"),
        ("L", "L22104", "Exec.3 mans. MT", 2.162, None, None, "A453", "XN4 Mans/termin Linii MT cablu subt NUM"),
        ("L", "L31104", "Constr. PT box", 2, 1, 1, "A441", "YA1Furniz/ const/ achiz post prefabr NUM"),
        ("L", "L31130", "Rezistenta anticondens", 9, None, None, "A441", "YQ1 Alte act invest PT(exclus PTA) NUM"),
        ("I", "I11101", "Verticalizare stalp", 5, None, None, "A451/A452", "XA2 Mont stalpi NUM / XA3 Mont stalp iz NUM"),
        ("L", "L11201", "Montare stalp", 4, None, None, "A451/A452", "XA2 Mont stalpi neiz NUM / XA3 Mont stalp iz NUM"),
        ("L", "L32201", "Realiz.inlocurii recup. C-JT", 3, None, None, "A61F",
         "YF1 Montare aparate TLC NUM / YI1 Demontare aparate TLC NUM / YR1 Inloc aparate TLC NUM"),
        ("L", "L99999", "Fara activitate", 1, None, None, "—", "Fara activitate in catalog"),
        ("L", "L22101", "Pozare cablu (rand duplicat)", 1.4, None, None, "A453", "XL4 Pozare LES MT KM"),
        (None, None, "TOTAL CAPITOL", 999, None, None, None, None),
        ("M", "M30007", "Furniz. leg. pamant", 60, None, None, "A412", "YE1 Mont compartim MT-JT In post NUM"),
    ]
    for r, (part, code, desc, qty, pt1, pt2, net, act) in enumerate(data, start=18):
        norme.cell(row=r, column=1, value=part)
        norme.cell(row=r, column=2, value=code)
        norme.cell(row=r, column=3, value=desc)
        norme.cell(row=r, column=8, value=qty)
        norme.cell(row=r, column=9, value=pt1)
        norme.cell(row=r, column=10, value=pt2)
        norme.cell(row=r, column=20, value=net)
        norme.cell(row=r, column=21, value=act)
    norme.cell(row=40, column=2, value="L21102")  # cantitate zero -> ignorat
    norme.cell(row=40, column=8, value=0)

    mat = wb.create_sheet("MAT MT")
    header = ["Cod normă", "Instalația", "Matricola", "Descriere normă", "Fișă", "UM", "Puncte",
              "Cantitate totala", "PT1\nDA102028076", "PT2"]
    for col, value in enumerate(header, start=1):
        mat.cell(row=13, column=col, value=value)
    mat.cell(row=13, column=34, value="Activitate SAP")
    mat.cell(row=13, column=35, value="Denumire activitate SAP")
    mat_rows = [
        ("E20056", "CBL MT 3P", 393.6, None, None, "A453", "XL4 Pozare LES MT KM"),
        ("E20005", "PRESETUPA", 4, 2, 2, "A453", "XN4 Mans/termin Linii MT cablu subt NUM"),
        ("E10029", "Separator", 1, None, None, "A451/A452", "XT2 Mont armat/desc/separatoare NUM"),
        ("E31000", "Cablu neizolat", 2, None, None, "A451", "XD2 Intind cond neizolat linii MT KM"),
    ]
    for r, (code, desc, qty, pt1, pt2, net, act) in enumerate(mat_rows, start=14):
        mat.cell(row=r, column=1, value=code)
        mat.cell(row=r, column=4, value=desc)
        mat.cell(row=r, column=8, value=qty)
        mat.cell(row=r, column=9, value=pt1)
        mat.cell(row=r, column=10, value=pt2)
        mat.cell(row=r, column=34, value=net)
        mat.cell(row=r, column=35, value=act)

    backup = wb.create_sheet("BackupFormule")  # exclusă din config
    backup["A1"], backup["B1"], backup["C1"] = "Cod Norma", "Cant.", "Activitate SAP"
    backup["A2"], backup["B2"], backup["C2"] = "L22101", 100, "A453"

    wb.save(path)
    return path


@pytest.fixture
def workbook(tmp_path) -> Path:
    return build_workbook(tmp_path / "deviz_test.xlsx")


@pytest.fixture
def cfg():
    from sapwbs.config import AppConfig

    return AppConfig()
