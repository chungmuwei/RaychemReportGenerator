from __future__ import annotations

import re
from datetime import datetime


class UserInputError(ValueError):
    """Raised when user-entered GUI values cannot be used to generate a report."""


def require_text(value: object, label: str) -> str:
    text = "" if value is None else str(value).strip()
    if not text:
        raise UserInputError(f"{label}不可空白。")
    return text


def parse_positive_float(value: object, label: str) -> float:
    text = require_text(value, label)
    try:
        parsed = float(text)
    except ValueError as exc:
        raise UserInputError(f"{label}必須是數字。") from exc
    if parsed <= 0:
        raise UserInputError(f"{label}必須大於 0。")
    return parsed


def parse_positive_int(value: object, label: str) -> int:
    text = require_text(value, label)
    try:
        parsed = int(text)
    except ValueError as exc:
        raise UserInputError(f"{label}必須是整數。") from exc
    if parsed <= 0:
        raise UserInputError(f"{label}必須大於 0。")
    return parsed


def validate_report_date(value: object, label: str = "檢測日期") -> str:
    text = require_text(value, label)
    if not re.fullmatch(r"\d{4}/\d{2}/\d{2}", text):
        raise UserInputError(f"{label}格式必須為 YYYY/MM/DD。")
    try:
        datetime.strptime(text, "%Y/%m/%d")
    except ValueError as exc:
        raise UserInputError(f"{label}不是有效日期。") from exc
    return text


def format_numeric_text(value: object) -> str:
    text = str(value)

    def format_match(match: re.Match) -> str:
        raw_number = match.group(0)
        normalized = raw_number.replace(",", "")
        if "." in normalized:
            whole, decimal = normalized.split(".", 1)
            return f"{int(whole):,}.{decimal}"
        return f"{int(normalized):,}"

    return re.sub(r"\d[\d,]*(?:\.\d+)?", format_match, text)


def format_type_1_viscosity(viscosity: float) -> str:
    formatted = f"{viscosity:.4g}" if viscosity < 1000 else str(round(viscosity))
    return format_numeric_text(formatted)


def mousewheel_scroll_units(delta: int) -> int:
    if delta == 0:
        return 0
    if abs(delta) >= 120:
        return int(-delta / 120)
    return -1 if delta > 0 else 1

