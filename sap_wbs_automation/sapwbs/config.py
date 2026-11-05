"""Configurația aplicației.

Valorile implicite sunt definite aici; fișierul ``config.yaml`` le suprascrie.
Orice cheie necunoscută din YAML produce o eroare, ca să nu treacă neobservate
greșelile de scriere.
"""
from __future__ import annotations

import dataclasses
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional


class ConfigError(Exception):
    """Configurație invalidă."""


# --------------------------------------------------------------------------- #
#  Secțiuni
# --------------------------------------------------------------------------- #
def _default_columns() -> Dict[str, List[str]]:
    return {
        "code": ["Cod Norma", "Cod normă", "Cod norma", "Cod articol", "Cod serviciu", "E"],
        "quantity": [
            "Cant.", "Cant", "Cantitate", "Cantitate totala", "Cant. Totala",
            "Cantitate totală", "Norma", "Norme_ore", "Norme ore",
        ],
        "network": ["Activitate SAP", "Retea standard", "Retea SAP", "Rețea standard"],
        "activity": ["Denumire activitate SAP", "Activitate SAP denumire", "Denumire activitate"],
        "part": ["Partea", "Parte"],
        "description": ["Descriere", "Descriere normă", "Descriere norma", "Denumire"],
        "install_type": ["T.I.", "Instalatia", "Instalația", "Tip instalatie"],
        "override": ["Activitate SAP ales", "Activitate SAP aleasa", "Activitate aleasa"],
    }


@dataclass
class ExcelConfig:
    header_scan_rows: int = 40
    max_scan_col: int = 120
    skip_hidden_sheets: bool = True
    include_sheets: List[str] = field(default_factory=list)  # regex; gol = toate
    exclude_sheets: List[str] = field(default_factory=lambda: ["^BackupFormule$", "^Anexa", "vechi"])
    stop_after_empty_rows: int = 60
    columns: Dict[str, List[str]] = field(default_factory=_default_columns)
    required_columns: List[str] = field(default_factory=lambda: ["code", "quantity", "network"])
    pt_column_regex: str = r"^PT\s*(\d+)$"
    code_regex: str = r"^[A-Z]\d{5}$"


@dataclass
class NetworkInfo:
    level3: str = ""
    voltage: Optional[str] = None   # MT / JT / None (oricare)
    solution: Optional[str] = None  # LEA / LES / None (oricare)
    desc: str = ""


def _default_networks() -> Dict[str, NetworkInfo]:
    return {
        "A412": NetworkInfo(level3="M1", desc="APARATAJ"),
        "A414": NetworkInfo(level3="M1", desc="TRANSFORMATOR"),
        "A441": NetworkInfo(level3="M1", desc="CLADIRE"),
        "A451": NetworkInfo(level3="M1", voltage="MT", solution="LEA", desc="LEA MT COND NEIZOLAT"),
        "A452": NetworkInfo(level3="M1", voltage="MT", solution="LEA", desc="LEA MT COND IZOLAT"),
        "A453": NetworkInfo(level3="M1", voltage="MT", solution="LES", desc="RETEA LES MT"),
        "A462": NetworkInfo(level3="B1", voltage="JT", solution="LEA", desc="RETEA LEA JT"),
        "A463": NetworkInfo(level3="B1", voltage="JT", solution="LES", desc="RETEA LES JT"),
        "A471": NetworkInfo(level3="B1", voltage="JT", desc="BRANSAMENTE JT"),
        "A61F": NetworkInfo(level3="T1", desc="TLC"),
    }


def _default_keyword_rules() -> List[Dict[str, str]]:
    return [
        {"description_regex": r"^(demont|demol|dezafect|desfiint)", "prefer_regex": r"\bdemont|\bdem\b"},
        {"description_regex": r"\binloc", "prefer_regex": r"\binloc"},
        {"description_regex": r".*", "prefer_regex": r"\bmont|\bfurniz|\bpozare|\binstal"},
    ]


@dataclass
class RulesConfig:
    allowed_parts: List[str] = field(default_factory=lambda: ["L", "M", "S", "E"])
    part_order: List[str] = field(default_factory=lambda: ["L", "M", "S", "E"])
    activity_suffix_strip: List[str] = field(default_factory=lambda: ["NUM"])
    activity_key_regex: str = r"^([A-Z0-9]{3})"
    activity_split_regex: str = r"\s+/\s+(?=[A-Z0-9]{3})"
    network_regex: str = r"^A[0-9A-Z]{3}$"
    network_split_regex: str = r"\s*/\s*"
    activity_name_max_len: int = 40
    ambiguity_fallback: str = "first"  # first | skip
    networks: Dict[str, NetworkInfo] = field(default_factory=_default_networks)
    activity_keyword_rules: List[Dict[str, str]] = field(default_factory=_default_keyword_rules)
    context_sheet: str = "Date"
    voltage_labels: List[str] = field(default_factory=lambda: ["Nivel Tensiune"])
    solution_labels: List[str] = field(default_factory=lambda: ["Solutie proiectata"])
    level3_table_headers: List[str] = field(default_factory=lambda: ["Activitate SAP", "Nivel 3"])


@dataclass
class FieldSource:
    """De unde se citește o valoare de proiect (nume definit > etichetă > adresă > implicit)."""

    defined_names: List[str] = field(default_factory=list)
    labels: List[str] = field(default_factory=list)
    sheets: List[str] = field(default_factory=list)  # gol = toate foile
    cell: str = ""  # adresă explicită, ex. "Date!C5"
    default: str = ""


@dataclass
class WbsConfig:
    level3_code: str = "+{n3}"
    level3_desc: Dict[str, str] = field(
        default_factory=lambda: {"M1": "MEDIE TENSIUNE", "B1": "JOASA TENSIUNE", "T1": "TELECONTROL"}
    )
    level4_code: str = "+{net}R{idx}"
    level4_desc_pt: str = "{base} PT{pt}"
    level4_desc_common: str = "{base}"
    level4_desc_common_with_pt: str = "{base} COMUN"
    desc_max_len: int = 40
    project_definition: FieldSource = field(
        default_factory=lambda: FieldSource(
            defined_names=["COD_PROIECT"],
            labels=["Cod proiect SAP", "Definitie proiect", "Proiect SAP", "Cod proiect"],
            sheets=["Date"],
        )
    )
    budget_code: FieldSource = field(
        default_factory=lambda: FieldSource(
            defined_names=["COD_BUGET"],
            labels=["Cod buget", "Cod bugetar", "Cod buget SAP"],
            sheets=["Date"],
        )
    )
    location: FieldSource = field(
        default_factory=lambda: FieldSource(labels=["Locatie SAP", "Locatie"], sheets=["Date"])
    )
    assignment_field_by_network: Dict[str, str] = field(
        default_factory=lambda: {"A412": "EQUNR", "A414": "EQUNR", "A441": "EQUNR"}
    )
    assignment_default_field: str = "TPLNR"
    pt_assignment_labels: Dict[str, List[str]] = field(
        default_factory=lambda: {"EQUNR": ["Cod Post"], "TPLNR": ["Cod Linie"]}
    )
    default_tplnr: str = ""
    default_equnr: str = ""
    level3_tplnr: Dict[str, str] = field(default_factory=dict)
    budget_on_levels: List[int] = field(default_factory=lambda: [4])
    custom_fields_by_network: Dict[str, Dict[str, str]] = field(
        default_factory=lambda: {"A441": {"ZZPARTE": "1"}}
    )


@dataclass
class SapConfig:
    transaction: str = "CJ02"
    system_name: str = ""
    client: str = ""
    user: str = ""
    std_network_alternative: str = "QS"
    network_popup_values: List[str] = field(default_factory=lambda: ["AMC", "MBT1"])
    purchasing_group: str = "AMC"   # AFVGD-EKGRP
    requisitioner: str = "AMC"      # AFVGD-AFNAM
    set_unlimited_overfulfillment: bool = True  # ESLL-UEBTK
    price_one_regex: str = r"^E00\d{3}$"        # servicii valorice -> preț brut 1
    decimal_separator: str = ","
    quantity_decimals: int = 3
    busy_timeout_s: float = 120.0
    max_popups: int = 10
    auto_confirm_info_popups: bool = True


@dataclass
class OutputConfig:
    log_dir: str = "logs"
    report_dir: str = "rapoarte"
    log_level: str = "INFO"


@dataclass
class AppConfig:
    excel: ExcelConfig = field(default_factory=ExcelConfig)
    rules: RulesConfig = field(default_factory=RulesConfig)
    wbs: WbsConfig = field(default_factory=WbsConfig)
    sap: SapConfig = field(default_factory=SapConfig)
    output: OutputConfig = field(default_factory=OutputConfig)
    base_dir: Path = field(default_factory=Path.cwd)

    def resolve_path(self, value: str) -> Path:
        path = Path(value)
        return path if path.is_absolute() else self.base_dir / path


# --------------------------------------------------------------------------- #
#  Încărcare YAML
# --------------------------------------------------------------------------- #
def _apply(obj: Any, data: Dict[str, Any], path: str) -> Any:
    """Aplică valorile din YAML peste un dataclass.

    Dicționarele se îmbină (cheile din YAML le suprascriu pe cele implicite),
    listele și valorile simple se înlocuiesc.
    """
    if not isinstance(data, dict):
        raise ConfigError(f"'{path}' trebuie să fie un dicționar")
    known = {f.name for f in dataclasses.fields(obj)}
    for key, value in data.items():
        if key not in known or key == "base_dir":
            raise ConfigError(f"Cheie necunoscută în configurație: '{path}.{key}'")
        current = getattr(obj, key)
        if dataclasses.is_dataclass(current):
            _apply(current, value or {}, f"{path}.{key}")
        elif isinstance(obj, RulesConfig) and key == "networks":
            if not isinstance(value, dict):
                raise ConfigError(f"'{path}.networks' trebuie să fie un dicționar")
            for net, spec in value.items():
                net_code = str(net).upper()
                base = current.get(net_code, NetworkInfo())
                current[net_code] = _apply(base, spec or {}, f"{path}.networks.{net}")
        elif isinstance(current, dict):
            if not isinstance(value, dict):
                raise ConfigError(f"'{path}.{key}' trebuie să fie un dicționar")
            current.update(value)
        else:
            setattr(obj, key, value)
    return obj


def load_config(path: Optional[Path]) -> AppConfig:
    """Încarcă configurația; fără fișier se folosesc valorile implicite."""
    cfg = AppConfig()
    if path is None:
        return cfg
    path = Path(path)
    if not path.exists():
        raise ConfigError(f"Fișierul de configurare nu există: {path}")
    try:
        import yaml  # import local: PyYAML e necesar doar când există fișier
    except ImportError as exc:  # pragma: no cover
        raise ConfigError("Lipsește pachetul PyYAML (pip install pyyaml)") from exc
    with path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}
    _apply(cfg, data, "config")
    cfg.base_dir = path.resolve().parent
    if cfg.rules.ambiguity_fallback not in ("first", "skip"):
        raise ConfigError("rules.ambiguity_fallback trebuie să fie 'first' sau 'skip'")
    return cfg
