"""Funcții utilitare pentru text și numere (fără dependențe externe)."""
from __future__ import annotations

import math
import re
import unicodedata
from typing import Any, Optional

_TOKEN_RE = re.compile(r"[a-z0-9]+")


def strip_diacritics(text: str) -> str:
    """Elimină diacriticele (ă, â, î, ș, ş, ț, ţ ...) păstrând litera de bază."""
    decomposed = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def norm(value: Any) -> str:
    """Normalizează un text pentru comparații robuste.

    "Cod normă" -> "cod norma", "Cant." -> "cant", "T.I." -> "t i".
    """
    if value is None:
        return ""
    text = strip_diacritics(str(value).lower())
    return " ".join(_TOKEN_RE.findall(text))


def first_line(value: Any) -> str:
    """Prima linie (ne-goală) dintr-o celulă cu text pe mai multe rânduri."""
    if value is None:
        return ""
    for line in str(value).splitlines():
        if line.strip():
            return line.strip()
    return ""


def clean_spaces(value: Any) -> str:
    """Colapsează spațiile multiple și elimină spațiile de la capete."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def parse_number(value: Any) -> Optional[float]:
    """Convertește o valoare din Excel în float; întoarce None dacă nu e număr.

    Acceptă numere native și texte de tip "1.234,56", "1,5", "2.75", "-".
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    text = str(value).strip().replace(" ", "").replace(" ", "")
    if not text or not re.search(r"\d", text):
        return None
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def format_quantity(value: float, decimals: int = 3, decimal_separator: str = ",") -> str:
    """Formatează o cantitate pentru câmpurile SAP (ex. 348.6 -> "348,6", 40.0 -> "40").

    Zecimalele inutile sunt eliminate; nu se folosește separator de mii
    (SAP acceptă valoarea fără el, indiferent de setările utilizatorului).
    """
    rounded = round(float(value), decimals)
    if rounded == 0:
        rounded = 0.0  # evită "-0"
    text = f"{rounded:.{decimals}f}"
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text.replace(".", decimal_separator)


def truncate(text: str, max_len: int) -> str:
    """Taie textul la lungimea maximă acceptată de câmpul SAP."""
    return text if len(text) <= max_len else text[:max_len].rstrip()
