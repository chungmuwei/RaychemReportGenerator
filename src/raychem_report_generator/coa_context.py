from __future__ import annotations

import re
import time
from datetime import date

from dateutil.relativedelta import relativedelta

from .coa_config import COUPLE
from .coa_utils import (
    UserInputError,
    format_numeric_text,
    format_type_1_viscosity,
    parse_positive_float,
    parse_positive_int,
    require_text,
    validate_report_date,
)


def type_1_uses_user_quantity(product_name: str, qty_products: set[str] | None) -> bool:
    """Return whether a type-1 product requires a user-entered quantity.

    Args:
        product_name: Selected product name.
        qty_products: Products configured to accept an entered quantity.

    Returns:
        ``True`` when the selected product uses an entered quantity.
    """
    return qty_products is not None and product_name in qty_products


def format_hy2537_base_weight(weight: str) -> str:
    """Remove the barrel multiplier from an HY2537 weight description.

    Args:
        weight: Configured HY2537 weight description.

    Returns:
        The base weight without a barrel multiplier.
    """
    return re.sub(r"\*\s*\d+\s*桶", "", weight, count=1)


def build_type_1_context(
    company: str,
    values: dict,
    product_specs: dict,
    qty_products: set[str] | None = None,
    include_gel_time: bool = True,
) -> dict:
    """Validate type-1 inputs and build a template rendering context.

    Args:
        company: Company identifier used to select product specifications.
        values: Raw values collected from the company form.
        product_specs: Product specifications keyed by company and product.
        qty_products: Products that require user-entered quantities.
        include_gel_time: Whether gel time is required and rendered.

    Returns:
        A validated and display-formatted template context.

    Raises:
        UserInputError: If required input or product configuration is invalid.
    """
    product_name = require_text(values.get("product_name"), "品名")
    if company not in product_specs:
        raise UserInputError(f"找不到 {company} 的產品規格。")
    if product_name not in product_specs[company]:
        raise UserInputError(f"找不到品名 {product_name} 的產品規格。")
    if product_name not in COUPLE:
        raise UserInputError(f"找不到品名 {product_name} 的搭配產品。")

    test_date = validate_report_date(values.get("date"))
    lot_no = require_text(values.get("lot_no"), "批號")
    viscosity = parse_positive_float(values.get("viscosity"), "黏度 cPs")
    uses_user_quantity = type_1_uses_user_quantity(product_name, qty_products)
    qty = (
        parse_positive_int(values.get("quantity"), "數量")
        if uses_user_quantity
        else None
    )
    gel_time = (
        parse_positive_int(values.get("gel_time"), "凝膠時間 sec")
        if include_gel_time
        else None
    )
    spec = product_specs[company][product_name]

    context = {
        "product_name": product_name,
        "date": test_date,
        "lot_no": lot_no,
        "weight": format_numeric_text(spec["weight"]),
        "unit_weight": format_numeric_text(spec.get("unit_weight", "")),
        "qty": format_numeric_text(spec.get("qty", "")),
        "viscosity_range": format_numeric_text(spec["viscosity_range"]),
        "appearance": spec["appearance"],
        "obs_appearance": spec["appearance"],
        "couple": COUPLE[product_name],
        "hardness": format_numeric_text(spec["hardness"]),
        "gel_time_range": format_numeric_text(spec["gel_time_range"]),
        "viscosity": format_type_1_viscosity(viscosity),
    }
    if include_gel_time:
        context["gel_time"] = format_numeric_text(gel_time)
    if uses_user_quantity:
        context["qty"] = format_numeric_text(qty)
    # Only UIC products currently define unit weight and quantity separately.
    if company == "uic":
        total_weight = int(context["unit_weight"]) * int(context["qty"])
        context["weight"] = format_numeric_text(total_weight)

    return context


def build_yuasa_context(values: dict) -> dict:
    """Validate Yuasa inputs and build a template rendering context.

    Args:
        values: Raw values collected from the Yuasa form.

    Returns:
        A validated context with calculated quantities and expiration date.

    Raises:
        UserInputError: If input is invalid or weights are not valid multiples.
    """
    test_date = validate_report_date(values.get("date"))
    lot_no = require_text(values.get("lot_no"), "批號")
    if not re.fullmatch(r"[A-Za-z]\d{6}.*", lot_no):
        raise UserInputError(
            "湯淺批號格式錯誤，前 7 碼需為 1 個英文字母"
            "加 6 個日期數字，"
            "例如 T260101。"
        )
    ay8000r_unit_weight = 4
    ay8000b_unit_weight = 4
    hy8000_unit_weight = 10

    ay8000r_weight = parse_positive_int(
        values.get("ay8000r_weight"), "AY8000R重量 Kg"
    )
    ay8000b_weight = parse_positive_int(
        values.get("ay8000b_weight"), "AY8000B重量 Kg"
    )
    hy8000_weight = parse_positive_int(values.get("hy8000_weight"), "HY8000重量 Kg")

    if ay8000r_weight % ay8000r_unit_weight != 0:
        raise UserInputError("AY8000R重量必須是4的倍數。")
    ay8000r_quantity = ay8000r_weight // ay8000r_unit_weight

    if ay8000b_weight % ay8000b_unit_weight != 0:
        raise UserInputError("AY8000B重量必須是4的倍數。")
    ay8000b_quantity = ay8000b_weight // ay8000b_unit_weight

    if hy8000_weight % hy8000_unit_weight != 0:
        raise UserInputError("HY8000重量必須是10的倍數。")
    hy8000_quantity = hy8000_weight // hy8000_unit_weight

    try:
        year = 2000 + int(lot_no[1:3])
        month = int(lot_no[3:5])
        day = int(lot_no[5:7])
        due_date = date(year, month, day) + relativedelta(months=+6)
    except ValueError as exc:
        raise UserInputError("湯淺批號內的日期不是有效日期。") from exc

    before_tensile_strength = parse_positive_int(
        values.get("before_tensile_strength"),
        "浸酸前引張強度 Kgf/cm2",
    )
    after_tensile_strength = parse_positive_int(
        values.get("after_tensile_strength"),
        "浸酸後引張強度 Kgf/cm2",
    )
    tensile_strength_diff = round(
        100
        * (before_tensile_strength - after_tensile_strength)
        / before_tensile_strength,
        2,
    )

    return {
        "product_name": "AY8000RB",
        "date": test_date,
        "lot_no": lot_no,
        "ay8000r_quant": format_numeric_text(ay8000r_quantity),
        "ay8000b_quant": format_numeric_text(ay8000b_quantity),
        "hy8000_quant": format_numeric_text(hy8000_quantity),
        "due_date": time.strftime("%Y/%m/%d", due_date.timetuple()),
        "ay8000r_viscosity": format_numeric_text(
            parse_positive_int(
                values.get("ay8000r_viscosity"), "AY8000R 黏度 cPs"
            )
        ),
        "ay8000b_viscosity": format_numeric_text(
            parse_positive_int(
                values.get("ay8000b_viscosity"), "AY8000B 黏度 cPs"
            )
        ),
        "hy8000_viscosity": format_numeric_text(
            "{:.1f}".format(
                parse_positive_float(
                    values.get("hy8000_viscosity"), "HY8000 黏度 cPs"
                )
            )
        ),
        "ay8000r_gel_time": format_numeric_text(
            parse_positive_int(
                values.get("ay8000r_gel_time"), "AY8000R 凝膠時間 sec"
            )
        ),
        "ay8000b_gel_time": format_numeric_text(
            parse_positive_int(
                values.get("ay8000b_gel_time"), "AY8000B 凝膠時間 sec"
            )
        ),
        "before_tensile_strength": format_numeric_text(before_tensile_strength),
        "after_tensile_strength": format_numeric_text(after_tensile_strength),
        "tensile_strength_diff": format_numeric_text(tensile_strength_diff),
        "acid_resistance": format_numeric_text(
            "{:.2f}".format(
                parse_positive_float(values.get("acid_resistance"), "耐酸性 %")
            )
        ),
    }
