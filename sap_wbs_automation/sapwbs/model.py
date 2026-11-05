"""Structurile de date comune: rânduri sursă, probleme, planul ierarhic."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

INFO = "INFO"
WARNING = "WARNING"
ERROR = "ERROR"


@dataclass
class Issue:
    """O problemă găsită la citire/planificare (apare în raportul pre-flight)."""

    severity: str
    message: str
    sheet: str = ""
    row: Optional[int] = None
    code: str = ""

    def location(self) -> str:
        if self.sheet and self.row:
            return f"{self.sheet}!{self.row}"
        return self.sheet


# --------------------------------------------------------------------------- #
#  Excel
# --------------------------------------------------------------------------- #
@dataclass
class HeaderMap:
    """Rezultatul detecției antetului într-o foaie."""

    sheet: str
    header_row: int                      # 1-based (ca în Excel)
    columns: Dict[str, int]              # coloană logică -> index 0-based
    pt_columns: Dict[int, int]           # număr PT -> index 0-based
    score: float = 0.0

    def letter(self, logical: str) -> str:
        from openpyxl.utils import get_column_letter

        idx = self.columns.get(logical)
        return get_column_letter(idx + 1) if idx is not None else "-"


@dataclass
class SourceRow:
    """Un rând de normă/material cu cantitate nenulă, citit din deviz."""

    sheet: str
    row: int
    code: str
    quantity: float
    pt_quantities: Dict[int, float]
    network_raw: str
    activity_raw: str
    part: str = ""
    description: str = ""
    install_type: str = ""
    override: str = ""

    @property
    def ref(self) -> str:
        return f"{self.sheet}!{self.row}"

    @property
    def common_quantity(self) -> float:
        """Cantitatea nealocată pe PT-uri (coloana totală minus suma PT)."""
        return self.quantity - sum(self.pt_quantities.values())


@dataclass
class PtInfo:
    number: int
    values: Dict[str, str] = field(default_factory=dict)  # ex. {"EQUNR": "...", "TPLNR": "..."}


@dataclass
class WorkbookContext:
    """Informații de proiect citite din foaia de date (context pentru reguli)."""

    voltage: str = ""
    solution: str = ""
    level3_by_network: Dict[str, str] = field(default_factory=dict)
    pts: Dict[int, PtInfo] = field(default_factory=dict)
    fields: Dict[str, str] = field(default_factory=dict)        # project_definition, budget_code, location
    field_origin: Dict[str, str] = field(default_factory=dict)  # de unde a fost citită valoarea


@dataclass
class ReadResult:
    rows: List[SourceRow]
    headers: List[HeaderMap]
    context: WorkbookContext
    issues: List[Issue]


# --------------------------------------------------------------------------- #
#  Plan
# --------------------------------------------------------------------------- #
@dataclass
class ServiceLine:
    code: str
    quantity: float
    description: str = ""
    sources: List[Tuple[str, int]] = field(default_factory=list)

    @property
    def source_refs(self) -> str:
        return ", ".join(f"{sheet}!{row}" for sheet, row in self.sources)


@dataclass
class ActivityPlan:
    key: str            # ex. "XL4"
    part: str           # L / M / S / E
    base_name: str      # ex. "XL4 Pozare LES MT KM"
    text: str           # ex. "XL4 Pozare LES MT KM (L)" (textul final în SAP)
    services: Dict[str, ServiceLine] = field(default_factory=dict)

    def add(self, code: str, quantity: float, description: str, source: Tuple[str, int]) -> None:
        line = self.services.get(code)
        if line is None:
            line = self.services[code] = ServiceLine(code=code, quantity=0.0, description=description)
        line.quantity += quantity
        line.sources.append(source)


@dataclass
class Wbs4Plan:
    code: str                    # ex. "+453R1"
    description: str
    network: str                 # ex. "A453"
    level3: str                  # ex. "M1"
    pt: Optional[int]            # None = lucrări comune (nealocate pe PT)
    assignments: Dict[str, str] = field(default_factory=dict)   # STORT/TPLNR/EQUNR
    custom_fields: Dict[str, str] = field(default_factory=dict)  # ZZCOD_BUGET, ZZPARTE ...
    activities: Dict[str, ActivityPlan] = field(default_factory=dict)

    def service_count(self) -> int:
        return sum(len(a.services) for a in self.activities.values())


@dataclass
class Wbs3Plan:
    n3: str                      # ex. "M1"
    code: str                    # ex. "+M1"
    description: str
    assignments: Dict[str, str] = field(default_factory=dict)
    custom_fields: Dict[str, str] = field(default_factory=dict)
    children: Dict[str, Wbs4Plan] = field(default_factory=dict)


@dataclass
class ProjectPlan:
    project_definition: str
    budget_code: str
    level3: Dict[str, Wbs3Plan]
    issues: List[Issue]
    stats: Dict[str, float] = field(default_factory=dict)

    def iter_wbs4(self):
        for n3 in self.level3.values():
            for n4 in n3.children.values():
                yield n3, n4

    def has_errors(self) -> bool:
        return any(i.severity == ERROR for i in self.issues)
