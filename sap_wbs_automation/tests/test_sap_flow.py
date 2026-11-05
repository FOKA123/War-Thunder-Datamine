"""Testarea fluxului SAP pe un SAP GUI simulat (tests/fake_sap.py)."""
from __future__ import annotations

from sapwbs.cj02 import (ActivityRow, WbsRow, choose_rows_to_keep, ident_matches, plan_copy_rounds,
                         plan_wbs_order)
from sapwbs.excel_reader import WorkbookReader
from sapwbs.model import ActivityPlan
from sapwbs.planner import Planner
from sapwbs.runner import ERROR, EXISTENT, OK, PlanRunner
from sapwbs.sap_session import SapSession

from fake_sap import FakeSap

STD_NETWORKS = {
    "A453": ["187 Realiz excav", "XL4 Pozare LES MT KM", "XN4 Mans/termin", "AUT Autorizatii", "XW0 Refaceri"],
    "A441": ["YA1Furniz/ const", "YQ1 Alte act invest", "PRG Proiectare"],
    "A451": ["XA2 Mont stalpi", "XD2 Intind cond", "XT2 Mont armat", "AUT Autorizatii"],
    "A412": ["YE1 Mont compartim", "YH1 Demont compartim"],
    "A61F": ["YF1 Montare aparate", "YI1 Demontare", "YR1 Inloc aparate"],
}


# --------------------------------------------------------------------------- #
#  Funcții pure
# --------------------------------------------------------------------------- #
def test_ident_matches():
    assert ident_matches("+453R1", "+453R1")
    assert ident_matches("C-0001-453R1", "+453R1")
    assert ident_matches("C-0001-M1", "+M1")
    assert not ident_matches("C-0001-1453R1", "+453R1")
    assert not ident_matches("", "+M1")


def test_copy_rounds_and_rows_to_keep():
    acts = [ActivityPlan("XL4", p, "XL4 x", f"XL4 x ({p})") for p in ("L", "E")]
    acts.append(ActivityPlan("AUT", "L", "AUT y", "AUT y (L)"))
    rounds = plan_copy_rounds(acts)
    assert [[a.text for a in r] for r in rounds] == [["XL4 x (L)", "AUT y (L)"], ["XL4 x (E)"]]

    new = [ActivityRow(4, "0050", "187 Realiz"), ActivityRow(5, "0060", "XL4 Pozare"),
           ActivityRow(6, "0070", "AUT Autorizatii"), ActivityRow(7, "0080", "XL4 Pozare")]
    keep, delete, missing = choose_rows_to_keep(new, ["XL4", "AUT", "ZZ9"], r"^([A-Z0-9]{3})")
    assert keep["XL4"].row == 5 and keep["AUT"].row == 6
    assert [r.row for r in delete] == [4, 7]
    assert missing == ["ZZ9"]


def test_wbs_order_respects_hierarchy():
    class G:  # grup de nivel 3 minimal
        def __init__(self, code, children):
            self.code = code
            self.children = {c: None for c in children}

    existing = [WbsRow(0, 1, "P", ""), WbsRow(1, 3, "P-B1", ""), WbsRow(2, 4, "P-463R1", ""),
                WbsRow(3, 3, "P-M1", ""), WbsRow(4, 4, "P-453R1", "")]
    groups = [G("+M1", ["+453R1", "+453R2"]), G("+B1", ["+463R1", "+463R2"]), G("+T1", ["+61FR1"])]
    order, blocked = plan_wbs_order(existing, groups)
    assert [g.code for g in order] == ["+B1", "+M1", "+T1"]
    assert blocked == {"+463R2"}  # B1 nu e ultimul nivel 3 din listă


# --------------------------------------------------------------------------- #
#  Flux complet pe SAP simulat
# --------------------------------------------------------------------------- #
def _plan(workbook, cfg):
    return Planner(cfg).build(WorkbookReader(workbook, cfg).read())


def _fake():
    wbs = [
        {"PRPS-STUFE": "1", "RCWBS-IDENT": "C-0001", "PRPS-POST1": "PROIECT"},
        {"PRPS-STUFE": "2", "RCWBS-IDENT": "C-0001-1", "PRPS-POST1": "LUCRARE"},
        {"PRPS-STUFE": "3", "RCWBS-IDENT": "C-0001-M1", "PRPS-POST1": "MEDIE TENSIUNE"},
        {"PRPS-STUFE": "4", "RCWBS-IDENT": "C-0001-453R3", "PRPS-POST1": "RETEA LES MT COMUN"},
    ]
    return FakeSap(project="C-0001", wbs=wbs, std_networks=STD_NETWORKS, invalid_services={"E31000"})


def _parent_codes(fake):
    """Pentru fiecare element, codul părintelui (rândul anterior cu nivel mai mic)."""
    result = {}
    for i, row in enumerate(fake.wbs):
        level = int(row["PRPS-STUFE"])
        parent = next((r for r in reversed(fake.wbs[:i]) if int(r["PRPS-STUFE"]) < level), None)
        result[row["RCWBS-IDENT"]] = parent["RCWBS-IDENT"] if parent else None
    return result


def test_full_run_and_rerun_is_idempotent(workbook, cfg):
    plan = _plan(workbook, cfg)
    fake = _fake()
    sap = SapSession(fake, cfg.sap)
    results = PlanRunner(sap, plan, cfg).run()

    statuses = {(r.wbs, r.code): r.status for r in results.rows.values()}
    assert statuses[("+451R1", "E31000")] == ERROR          # serviciu inexistent -> eroare pe rând
    assert statuses[("+451R1", "E10029")] == OK             # ... dar restul continuă
    assert all(s == OK for k, s in statuses.items() if k[1] != "E31000"), statuses

    parents = _parent_codes(fake)
    for code in ("+412R1", "+441R1", "+441R2", "+441R3", "+451R1", "+453R1", "+453R2"):
        assert parents[code] == "C-0001-M1", (code, parents[code])
    assert parents["+T1"] == "C-0001-1"
    assert parents["+61FR1"] == "+T1"

    by_ident = {row["RCWBS-IDENT"]: row for row in fake.wbs}
    cladire = by_ident["+441R1"]
    assert cladire["PRPS-POST1"] == "CLADIRE PT1"
    assert cladire["PRPS-EQUNR"] == "DA102028076" and cladire["PRPS-STORT"] == "DA1C111"
    assert cladire["ZZCOD_BUGET"] == "REM_001604" and cladire["ZZPARTE"] == "1"
    assert cladire["spo4_values"] == ["AMC", "MBT1"]
    texts = [a["AFVGD-LTXA1"] for a in cladire["acts"]]
    assert sorted(texts) == ["YA1Furniz/ const/ achiz post prefabr (L)"]
    services = {s["ESLL-SRVPOS"]: s for s in cladire["acts"][0]["services"]}
    assert services["L31104"]["ESLL-MENGE"] == "1"
    assert services["L31104"]["ESLL-UEBTK"] is True
    assert cladire["acts"][0]["EKGRP"] == "AMC"

    # WBS existent (+453R3): activitățile L și E pentru XL4 au necesitat două copieri ale rețelei
    les = by_ident["C-0001-453R3"]
    les_texts = sorted(a["AFVGD-LTXA1"] for a in les["acts"])
    assert "XL4 Pozare LES MT KM (L)" in les_texts and "XL4 Pozare LES MT KM (E)" in les_texts
    assert not any(t.startswith(("187", "XW0")) for t in les_texts)  # activitățile nefolosite șterse
    xl4_l = next(a for a in les["acts"] if a["AFVGD-LTXA1"] == "XL4 Pozare LES MT KM (L)")
    assert [(s["ESLL-SRVPOS"], s["ESLL-MENGE"]) for s in xl4_l["services"]] == [("L22101", "350")]

    # a doua rulare nu dublează nimic
    rows_before = len(fake.wbs)
    acts_before = {r["RCWBS-IDENT"]: len(r.get("acts", [])) for r in fake.wbs}
    results2 = PlanRunner(sap, plan, cfg).run()
    assert len(fake.wbs) == rows_before
    assert {r["RCWBS-IDENT"]: len(r.get("acts", [])) for r in fake.wbs} == acts_before
    statuses2 = {(r.wbs, r.code): r.status for r in results2.rows.values()}
    assert statuses2[("+441R1", "L31104")] == EXISTENT
    assert statuses2[("+451R1", "E31000")] == ERROR  # tot invalid


def test_errors_are_contained_and_run_continues(workbook, cfg):
    """Eroare în activitate (lipsă grup cumpărare) și la nivel WBS (rețea standard inexistentă)."""
    plan = _plan(workbook, cfg)
    networks = {k: v for k, v in STD_NETWORKS.items() if k != "A61F"}  # A61F lipsește din SAP
    fake = FakeSap(project="C-0001", std_networks=networks)
    cfg.sap.purchasing_group = ""  # SAP refuză serviciile fără grup de cumpărare
    results = PlanRunner(SapSession(fake, cfg.sap), plan, cfg).run()

    statuses = {(r.wbs, r.code): (r.status, r.message) for r in results.rows.values()}
    assert statuses[("+441R1", "L31104")][0] == ERROR
    assert "grupul de cumpărare" in statuses[("+441R1", "L31104")][1]
    assert statuses[("+61FR1", "L32201")][0] == ERROR
    assert "A61F" in statuses[("+61FR1", "L32201")][1]
    assert not results.aborted          # erorile nu au oprit execuția
    assert fake.popups == []            # nicio fereastră rămasă deschisă
    created = [r["RCWBS-IDENT"] for r in fake.wbs]
    assert "+453R3" in created and "+61FR1" in created
