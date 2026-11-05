"""Regulile de deducție: rețea standard, activitate, sufix, denumiri.

Legătura normă -> activitate vine din coloanele "Activitate SAP" (rețeaua
standard, ex. A453) și "Denumire activitate SAP" (ex. "XL4 Pozare LES MT KM").
Când există mai multe variante ("A451/A452", "YL1 ... / YM1 ...") se aplică,
în ordine:

1. coloana opțională de forțare ("Activitate SAP ales");
2. perechi poziționale (rețeaua i <-> activitatea i, dacă listele au aceeași lungime);
3. contextul din foaia Date (nivel tensiune MT/JT, soluție LEA/LES);
4. rețelele deja folosite de rânduri neambigue din același deviz;
5. reguli de cuvinte cheie pe descrierea normei (demontare / înlocuire / montare);
6. politica de rezervă din config (prima variantă cu avertizare, sau sări rândul).
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Iterable, List, Optional, Sequence, Set

from .config import AppConfig, NetworkInfo
from .model import SourceRow, WorkbookContext
from .textutil import clean_spaces, norm, truncate


@dataclass
class Resolution:
    network: Optional[str] = None
    activity_base: Optional[str] = None
    key: Optional[str] = None
    part: str = ""
    skip_reason: str = ""
    warnings: List[str] = field(default_factory=list)
    ambiguous_network: bool = False

    @property
    def ok(self) -> bool:
        return not self.skip_reason


# --------------------------------------------------------------------------- #
#  Funcții pure (ușor de testat)
# --------------------------------------------------------------------------- #
def split_networks(raw: str, cfg: AppConfig) -> List[str]:
    net_re = re.compile(cfg.rules.network_regex)
    parts = [p.strip().upper() for p in re.split(cfg.rules.network_split_regex, raw or "") if p.strip()]
    return [p for p in parts if net_re.match(p)]


def split_activities(raw: str, cfg: AppConfig) -> List[str]:
    key_re = re.compile(cfg.rules.activity_key_regex)
    parts = [clean_spaces(p) for p in re.split(cfg.rules.activity_split_regex, raw or "") if p.strip()]
    return [p for p in parts if key_re.match(p)]


def clean_activity_name(raw: str, strip_suffixes: Sequence[str]) -> str:
    """Elimină unitatea de măsură de la final (ex. "... NUM")."""
    text = clean_spaces(raw)
    for suffix in strip_suffixes:
        text = re.sub(rf"\s+{re.escape(suffix)}$", "", text)
    return text


def activity_key(name: str, key_regex: str) -> Optional[str]:
    match = re.match(key_regex, (name or "").strip().upper())
    return match.group(1) if match else None


def build_activity_text(base: str, part: str, max_len: int) -> str:
    """Textul final al activității în SAP, ex. "XL4 Pozare LES MT KM (L)"."""
    suffix = f" ({part})"
    return truncate(base, max_len - len(suffix)) + suffix


def part_of(row: SourceRow) -> str:
    """Partea (L/M/S/E): coloana "Partea" sau, lipsind, prima literă a codului."""
    return (row.part or row.code[:1]).upper()


def network_matches_context(info: Optional[NetworkInfo], ctx: WorkbookContext) -> bool:
    if info is None:
        return True
    voltage = norm(ctx.voltage).upper().replace(" ", "+")
    solution = norm(ctx.solution).upper().replace(" ", "+")
    if info.voltage and voltage and info.voltage.upper() not in voltage.split("+"):
        return False
    if info.solution and solution and info.solution.upper() not in solution.split("+"):
        return False
    return True


def choose_by_keywords(description: str, options: Sequence[str], rules: Iterable[dict]) -> Optional[str]:
    desc = norm(description)
    for rule in rules:
        if not re.search(rule.get("description_regex", ".*"), desc):
            continue
        prefer = re.compile(rule.get("prefer_regex", ""), re.IGNORECASE)
        for opt in options:
            if prefer.search(norm(opt) + " "):
                return opt
    return None


# --------------------------------------------------------------------------- #
#  Motorul de reguli
# --------------------------------------------------------------------------- #
class RuleEngine:
    def __init__(self, cfg: AppConfig, context: WorkbookContext):
        self.cfg = cfg
        self.ctx = context
        self.used_networks: Set[str] = set()

    def learn_used_networks(self, rows: Iterable[SourceRow]) -> None:
        """Rețelele care apar neambiguu în deviz (folosite la dezambiguizare)."""
        for row in rows:
            nets = split_networks(row.network_raw, self.cfg)
            if len(nets) == 1 and part_of(row) in self.cfg.rules.allowed_parts:
                self.used_networks.add(nets[0])

    def resolve(self, row: SourceRow) -> Resolution:
        rcfg = self.cfg.rules
        res = Resolution(part=part_of(row))
        if res.part not in rcfg.allowed_parts:
            res.skip_reason = f"partea '{res.part}' nu este procesată"
            return res

        nets = split_networks(row.network_raw, self.cfg)
        acts = split_activities(row.activity_raw, self.cfg)
        override_net, override_key = self._parse_override(row.override)

        if override_net:
            if nets and override_net not in nets:
                res.warnings.append(f"rețeaua forțată {override_net} nu e printre variantele {nets}")
            nets_for_pairing = nets
            chosen_net = override_net
        else:
            nets_for_pairing = nets
            chosen_net = None

        if not nets and not chosen_net:
            res.skip_reason = (f"fără rețea standard validă în 'Activitate SAP' "
                               f"('{row.network_raw or 'gol'}')")
            return res
        if not acts:
            res.skip_reason = (f"fără activitate validă în 'Denumire activitate SAP' "
                               f"('{row.activity_raw or 'gol'}')")
            return res

        if chosen_net is None:
            chosen_net = self._choose_network(nets, res)
            if chosen_net is None:
                return res

        res.network = chosen_net
        chosen_act = self._choose_activity(row, chosen_net, nets_for_pairing, acts, override_key, res)
        if chosen_act is None:
            return res

        base = clean_activity_name(chosen_act, rcfg.activity_suffix_strip)
        res.activity_base = base
        res.key = activity_key(base, rcfg.activity_key_regex)
        if not res.key:
            res.skip_reason = f"nu pot extrage codul activității din '{base}'"
        return res

    # -- detalii ----------------------------------------------------------- #
    def _parse_override(self, raw: str):
        if not raw:
            return None, None
        tokens = re.split(r"[\s;|,/]+", raw.strip().upper())
        net_re = re.compile(self.cfg.rules.network_regex)
        net = next((t for t in tokens if net_re.match(t)), None)
        key = next((t for t in tokens if t != net and re.match(r"^[A-Z0-9]{3}$", t)), None)
        return net, key

    def _choose_network(self, nets: List[str], res: Resolution) -> Optional[str]:
        if len(nets) == 1:
            return nets[0]
        rcfg = self.cfg.rules
        candidates = list(nets)
        by_ctx = [n for n in candidates if network_matches_context(rcfg.networks.get(n), self.ctx)]
        if by_ctx:
            candidates = by_ctx
        if len(candidates) > 1:
            by_use = [n for n in candidates if n in self.used_networks]
            if by_use:
                candidates = by_use
        if len(candidates) == 1:
            res.warnings.append(f"rețea aleasă {candidates[0]} din {'/'.join(nets)} (context/utilizare)")
            return candidates[0]
        res.ambiguous_network = True
        if rcfg.ambiguity_fallback == "skip":
            res.skip_reason = (f"rețea ambiguă {'/'.join(nets)}; completați coloana "
                               f"'Activitate SAP ales'")
            return None
        res.warnings.append(f"rețea ambiguă {'/'.join(nets)} -> aleasă prima variantă "
                            f"compatibilă: {candidates[0]}")
        return candidates[0]

    def _choose_activity(self, row: SourceRow, chosen_net: str, nets: List[str], acts: List[str],
                         override_key: Optional[str], res: Resolution) -> Optional[str]:
        rcfg = self.cfg.rules
        if override_key:
            for act in acts:
                if activity_key(act, rcfg.activity_key_regex) == override_key:
                    return act
            res.warnings.append(f"activitatea forțată {override_key} nu e printre variante")
        if len(acts) == 1:
            return acts[0]
        if len(nets) == len(acts) and chosen_net in nets:
            return acts[nets.index(chosen_net)]
        picked = choose_by_keywords(row.description, acts, rcfg.activity_keyword_rules)
        if picked:
            if picked != acts[0]:
                res.warnings.append(f"activitate aleasă după descriere: '{picked}'")
            else:
                res.warnings.append(f"activitate cu variante; aleasă '{picked}' după descriere")
            return picked
        if rcfg.ambiguity_fallback == "skip":
            res.skip_reason = f"activitate ambiguă: {' / '.join(acts)}"
            return None
        res.warnings.append(f"activitate ambiguă -> aleasă prima variantă: '{acts[0]}'")
        return acts[0]


# --------------------------------------------------------------------------- #
#  Denumiri WBS
# --------------------------------------------------------------------------- #
def network_short(network: str) -> str:
    """A453 -> 453, A61F -> 61F (partea folosită în codul WBS de nivel 4)."""
    return network[1:] if network[:1].upper() == "A" else network


def level4_code(cfg: AppConfig, network: str, index: int) -> str:
    return cfg.wbs.level4_code.format(net=network_short(network), idx=index, network=network)


def level4_description(cfg: AppConfig, network: str, pt: Optional[int], network_has_pt: bool) -> str:
    info = cfg.rules.networks.get(network)
    base = (info.desc if info and info.desc else network)
    if pt is not None:
        text = cfg.wbs.level4_desc_pt.format(base=base, pt=pt, network=network)
    elif network_has_pt:
        text = cfg.wbs.level4_desc_common_with_pt.format(base=base, network=network)
    else:
        text = cfg.wbs.level4_desc_common.format(base=base, network=network)
    return truncate(text, cfg.wbs.desc_max_len)


def level3_code(cfg: AppConfig, n3: str) -> str:
    return cfg.wbs.level3_code.format(n3=n3)


def level3_description(cfg: AppConfig, n3: str) -> str:
    return truncate(cfg.wbs.level3_desc.get(n3, n3), cfg.wbs.desc_max_len)
