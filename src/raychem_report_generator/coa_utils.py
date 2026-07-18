from __future__ import annotations

import re
from datetime import datetime


class UserInputError(ValueError):
    """Raised when user-entered GUI values cannot be used to generate a report."""


def require_text(value: object, label: str) -> str:
    """Return a stripped value or reject empty input.

    Args:
        value: User-provided value to normalize.
        label: Field name included in validation errors.

    Returns:
        The normalized, non-empty text.

    Raises:
        UserInputError: If the value is empty after normalization.
    """
    text = "" if value is None else str(value).strip()
    if not text:
        raise UserInputError(f"{label}不可空白。")
    return text


def parse_positive_float(value: object, label: str) -> float:
    """Parse a user-entered value as a positive floating-point number.

    Args:
        value: User-provided numeric value.
        label: Field name included in validation errors.

    Returns:
        The parsed positive floating-point number.

    Raises:
        UserInputError: If the value is empty, nonnumeric, or not positive.
    """
    text = require_text(value, label)
    try:
        parsed = float(text)
    except ValueError as exc:
        raise UserInputError(f"{label}必須是數字。") from exc
    if parsed <= 0:
        raise UserInputError(f"{label}必須大於 0。")
    return parsed


def parse_positive_int(value: object, label: str) -> int:
    """Parse a user-entered value as a positive integer.

    Args:
        value: User-provided numeric value.
        label: Field name included in validation errors.

    Returns:
        The parsed positive integer.

    Raises:
        UserInputError: If the value is empty, not an integer, or not positive.
    """
    text = require_text(value, label)
    try:
        parsed = int(text)
    except ValueError as exc:
        raise UserInputError(f"{label}必須是整數。") from exc
    if parsed <= 0:
        raise UserInputError(f"{label}必須大於 0。")
    return parsed


def validate_report_date(value: object, label: str = "檢測日期") -> str:
    """Validate and return a report date formatted as ``YYYY/MM/DD``.

    Args:
        value: User-provided date value.
        label: Field name included in validation errors.

    Returns:
        The validated date string.

    Raises:
        UserInputError: If the value is empty, malformed, or not a real date.
    """
    text = require_text(value, label)
    if not re.fullmatch(r"\d{4}/\d{2}/\d{2}", text):
        raise UserInputError(f"{label}格式必須為 YYYY/MM/DD。")
    try:
        datetime.strptime(text, "%Y/%m/%d")
    except ValueError as exc:
        raise UserInputError(f"{label}不是有效日期。") from exc
    return text


def format_numeric_text(value: object) -> str:
    """Add thousands separators to every numeric substring in a value.

    Args:
        value: Value containing one or more numeric substrings.

    Returns:
        Text with grouped integer portions and unchanged decimal portions.
    """
    text = str(value)

    def format_match(match: re.Match) -> str:
        """Format one matched number with thousands separators.

        Args:
            match: Regular-expression match containing a numeric substring.

        Returns:
            The grouped numeric substring.
        """
        raw_number = match.group(0)
        normalized = raw_number.replace(",", "")
        if "." in normalized:
            whole, decimal = normalized.split(".", 1)
            return f"{int(whole):,}.{decimal}"
        return f"{int(normalized):,}"

    return re.sub(r"\d[\d,]*(?:\.\d+)?", format_match, text)


def format_type_1_viscosity(viscosity: float) -> str:
    """Format a type-1 viscosity value for display in a report.

    Args:
        viscosity: Validated viscosity measurement.

    Returns:
        A compact, thousands-separated viscosity string.
    """
    formatted = f"{viscosity:.4g}" if viscosity < 1000 else str(round(viscosity))
    return format_numeric_text(formatted)
