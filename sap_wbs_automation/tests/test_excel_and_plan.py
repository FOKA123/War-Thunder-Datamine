from __future__ import annotations

import pytest

from sap_wbs import AppConfig, load_config, ConfigError
from sap_wbs import WorkbookReader, detect_header, match_columns, SheetGrid, score_header
from sap_wbs import ERROR, SourceRow, WorkbookContext
from sap_wbs import Planner
from sap_wbs import RuleEngine, build_activity_text, clean_activity_name, split_activities
from sap_wbs import format_quantity, norm, parse_number


# --------------------------------------------------------------------------- #
#  Utilitare
# --------------------------------------------------------------------------- #
def test_norm_and_numbers():
    assert norm("Cod normă") == "cod norma"
    assert norm("  Cant. ") == "cant"
    assert norm("Instalația") == "instalatia"
    assert parse_number("1.234,5") == pytest.approx(1234.5)
    assert parse_number("2,162") == pytest.approx(2.162)
    assert parse_number("-") is None
    assert format_quantity(348.6) == "348,6"
    assert format_quantity(40.0) == "40"
    assert format_quantity(3.5999999999999996) == "3,6"
    assert format_quantity(17649, 2, ",") == "17649"


def test_score_header_prefers_exact():
    assert score_header("Cod Norma", ["Cod Norma"]) == 100
    assert score_header("Cod Sap", ["E"]) == 0  # alias scurt doar exact
    cols = match_columns(["Partea", "Cod Norma", "Descriere", "Cod Sap", "Cant."],
                         AppConfig().excel.columns, ["code", "quantity"])
    assert cols["code"][0] == 1
    assert cols["quantity"][0] == 4
    assert cols["part"][0] == 0


def test_header_detection_ignores_position():
    cfg = AppConfig()
    rows = [[None] * 6 for _ in range(30)]
    rows[22] = ["x", "Retea standard", "Cantitate", None, "Cod articol", "PT 1"]
    header = detect_header(SheetGrid("S", rows), cfg)
    assert header.header_row == 23
    assert header.columns == {"network": 1, "quantity": 2, "code": 4}
    assert header.pt_columns == {1: 5}


# --------------------------------------------------------------------------- #
#  Citirea devizului sintetic
# --------------------------------------------------------------------------- #
def test_reader_detects_sheets_columns_and_context(workbook, cfg):
    result = WorkbookReader(workbook, cfg).read()
    sheets = {h.sheet: h for h in result.headers}
    assert set(sheets) == {"Norme MT", "MAT MT"}  # BackupFormule e exclusă
    assert sheets["Norme MT"].header_row == 17
    assert sheets["MAT MT"].header_row == 13
    assert sheets["MAT MT"].pt_columns == {1: 8, 2: 9}  # "PT1\nDA102..." recunoscut

    codes = [r.code for r in result.rows]
    assert "L21102" not in codes          # cantitate zero
    assert codes.count("L22101") == 2     # două rânduri, agregate mai târziu
    l31104 = next(r for r in result.rows if r.code == "L31104")
    assert l31104.pt_quantities == {1: 1.0, 2: 1.0}
    assert l31104.common_quantity == pytest.approx(0)

    ctx = result.context
    assert (ctx.voltage, ctx.solution) == ("MT", "LES")
    assert ctx.level3_by_network["A61F"] == "T1"
    assert ctx.fields["budget_code"] == "REM_001604"
    assert ctx.fields["project_definition"] == "C-0001"
    assert "COD_PROIECT" in ctx.field_origin["project_definition"]
    assert ctx.fields["location"] == "DA1C111"
    assert ctx.pts[1].values == {"EQUNR": "DA102028076", "TPLNR": "DA1033617"}
    assert ctx.pts[2].values["EQUNR"] == "DA102028077"


def test_budget_from_explicit_cell(tmp_path, cfg):
    from conftest import build_workbook

    path = build_workbook(tmp_path / "fara_buget.xlsx", with_budget=False)
    assert WorkbookReader(path, cfg).read().context.fields["budget_code"] == ""
    cfg.wbs.budget_code.cell = "Date!C9"
    assert WorkbookReader(path, cfg).read().context.fields["budget_code"] == "DA1C111"


# --------------------------------------------------------------------------- #
#  Reguli
# --------------------------------------------------------------------------- #
def _row(**kw) -> SourceRow:
    base = dict(sheet="S", row=1, code="L11201", quantity=1.0, pt_quantities={},
                network_raw="", activity_raw="", part="L", description="")
    base.update(kw)
    return SourceRow(**base)


def test_activity_names(cfg):
    assert split_activities("XQ7 Recup / abandon cablu JT subt KM", cfg) == ["XQ7 Recup / abandon cablu JT subt KM"]
    assert split_activities("YE1 Mont NUM / YK1 Inloc NUM", cfg) == ["YE1 Mont NUM", "YK1 Inloc NUM"]
    assert clean_activity_name("XN4 Mans/termin Linii MT cablu subt NUM", ["NUM"]) == "XN4 Mans/termin Linii MT cablu subt"
    text = build_activity_text("YA1Furniz/ const/ achiz post prefabr", "L", 40)
    assert text == "YA1Furniz/ const/ achiz post prefabr (L)" and len(text) <= 40
    assert len(build_activity_text("187 Realiz excaV si/sau poz conducte KM", "L", 40)) <= 40


def test_rules_disambiguation(cfg):
    ctx = WorkbookContext(voltage="MT", solution="LES")
    engine = RuleEngine(cfg, ctx)
    engine.learn_used_networks([_row(network_raw="A451", code="E31000", part="")])

    # perechi poziționale + rețea deja folosită (A451) -> XA2
    res = engine.resolve(_row(network_raw="A451/A452", activity_raw="XA2 Mont neiz NUM / XA3 Mont iz NUM"))
    assert (res.network, res.key, res.activity_base) == ("A451", "XA2", "XA2 Mont neiz")

    # context JT/LEA exclude A451/A452 când tensiunea e JT
    engine_jt = RuleEngine(cfg, WorkbookContext(voltage="JT", solution="LEA"))
    res = engine_jt.resolve(_row(network_raw="A451/A452/A462", activity_raw="XA2 a NUM / XA3 b NUM / XA6 c NUM"))
    assert (res.network, res.key) == ("A462", "XA6")

    # rețea unică, activitate cu variante -> cuvinte cheie din descriere
    res = engine.resolve(_row(network_raw="A61F", description="Realiz.inlocurii recup.",
                              activity_raw="YF1 Montare NUM / YI1 Demontare NUM / YR1 Inloc NUM"))
    assert res.key == "YR1"
    res = engine.resolve(_row(network_raw="A414", description="Mont.sau recup.trafo",
                              activity_raw="YL1 Mont transf NUM / YM1 Demont transf NUM"))
    assert res.key == "YL1"

    # coloana de forțare
    res = engine.resolve(_row(network_raw="A451/A452", activity_raw="XA2 x NUM / XA3 y NUM", override="A452"))
    assert (res.network, res.key) == ("A452", "XA3")

    # partea I este ignorată; rețea lipsă -> rând sărit
    assert not engine.resolve(_row(part="I", network_raw="A453", activity_raw="XL4 x")).ok
    assert "fără rețea" in engine.resolve(_row(network_raw="—", activity_raw="XL4 x")).skip_reason


def test_ambiguity_skip_policy(cfg):
    cfg.rules.ambiguity_fallback = "skip"
    engine = RuleEngine(cfg, WorkbookContext())
    res = engine.resolve(_row(network_raw="A451/A452", activity_raw="XT2 Mont armat NUM"))
    assert not res.ok and "ambiguă" in res.skip_reason


# --------------------------------------------------------------------------- #
#  Plan
# --------------------------------------------------------------------------- #
def test_plan_structure(workbook, cfg):
    read = WorkbookReader(workbook, cfg).read()
    plan = Planner(cfg).build(read)
    assert plan.project_definition == "C-0001" and plan.budget_code == "REM_001604"
    assert not plan.has_errors()
    assert list(plan.level3) == ["M1", "T1"]

    m1 = plan.level3["M1"].children
    # A441: PT1, PT2 + lucrări comune (R3 = primul indice după ultimul PT din deviz)
    assert {"+441R1", "+441R2", "+441R3"} <= set(m1)
    assert m1["+441R1"].description == "CLADIRE PT1"
    assert m1["+441R3"].description == "CLADIRE COMUN"
    assert m1["+441R1"].assignments == {"STORT": "DA1C111", "EQUNR": "DA102028076"}
    assert m1["+441R1"].custom_fields == {"ZZCOD_BUGET": "REM_001604", "ZZPARTE": "1"}
    # A451 nu are nimic pe PT -> R1
    assert "+451R1" in m1 and m1["+451R1"].description == "LEA MT COND NEIZOLAT"
    # agregare: L22101 = 348.6 + 1.4
    xl4 = m1["+453R3"].activities["XL4 Pozare LES MT KM (L)"]
    assert xl4.services["L22101"].quantity == pytest.approx(350.0)
    assert len(xl4.services["L22101"].sources) == 2
    # materialele primesc sufixul (E); E20005 împărțit pe PT1/PT2
    assert m1["+453R1"].activities["XN4 Mans/termin Linii MT cablu subt (E)"].services["E20005"].quantity == 2
    # control de sumă
    assert plan.stats["cantitate_citita"] == pytest.approx(plan.stats["cantitate_planificata"])
    # rândul fără rețea apare ca avertizare
    assert any("L99999" == i.code for i in plan.issues)


def test_missing_budget_is_error(tmp_path, cfg):
    from conftest import build_workbook

    read = WorkbookReader(build_workbook(tmp_path / "x.xlsx", with_budget=False), cfg).read()
    plan = Planner(cfg).build(read)
    assert any(i.severity == ERROR and "buget" in i.message for i in plan.issues)
    assert not Planner(cfg).build(read, budget_override="B1").has_errors()


def test_config_yaml_roundtrip(tmp_path):
    path = tmp_path / "c.yaml"
    path.write_text("rules:\n  networks:\n    A441: {desc: CLADIRE NOUA}\nsap:\n  transaction: CJ20\n",
                    encoding="utf-8")
    cfg = load_config(path)
    assert cfg.rules.networks["A441"].desc == "CLADIRE NOUA"
    assert cfg.rules.networks["A441"].level3 == "M1"      # îmbinat cu implicitul
    assert "A453" in cfg.rules.networks
    assert cfg.sap.transaction == "CJ20"
    path.write_text("sap:\n  tranzactie: CJ02\n", encoding="utf-8")
    with pytest.raises(ConfigError):
        load_config(path)
