from __future__ import annotations

import re
import time
from datetime import date

from dateutil.relativedelta import relativedelta

from coa_config import COUPLE
from coa_utils import (
    UserInputError,
    format_numeric_text,
    format_type_1_viscosity,
    parse_positive_float,
    parse_positive_int,
    require_text,
    validate_report_date,
)


def type_1_uses_user_quantity(product_name: str, qty_products: set[str] | None) -> bool:
    return qty_products is not None and product_name in qty_products


def format_hy2537_base_weight(weight: str) -> str:
    return re.sub(r"\*\s*\d+\s*桶", "", weight, count=1)


def build_type_1_context(
    company: str,
    values: dict,
    product_specs: dict,
    qty_products: set[str] | None = None,
    include_gel_time: bool = True,
) -> dict:
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
    qty = parse_positive_int(values.get("quantity"), "數量") if uses_user_quantity else None
    gel_time = parse_positive_int(values.get("gel_time"), "凝膠時間 sec") if include_gel_time else None
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
    # if uses_user_quantity and company == "etacom" and product_name == "硬化劑HY2537":
    #     context["weight"] = format_numeric_text(format_hy2537_base_weight(spec["weight"]))
    
    # Only recalculate weight for UIC company
    # Currently, only UIC products have unit_weight and qty in the specs
    if company  == "uic":
        context["weight"] = format_numeric_text(int(context["unit_weight"]) * int(context["qty"]))

    return context


def build_yuasa_context(values: dict) -> dict:
    test_date = validate_report_date(values.get("date"))
    lot_no = require_text(values.get("lot_no"), "批號")
    if not re.fullmatch(r"[A-Za-z]\d{6}.*", lot_no):
        raise UserInputError("湯淺批號格式錯誤，前 7 碼需為 1 個英文字母加 6 個日期數字，例如 T260101。")

    try:
        year = 2000 + int(lot_no[1:3])
        month = int(lot_no[3:5])
        day = int(lot_no[5:7])
        due_date = date(year, month, day) + relativedelta(months=+6)
    except ValueError as exc:
        raise UserInputError("湯淺批號內的日期不是有效日期。") from exc

    before_tensile_strength = parse_positive_int(values.get("before_tensile_strength"), "浸酸前引張強度 Kgf/cm2")
    after_tensile_strength = parse_positive_int(values.get("after_tensile_strength"), "浸酸後引張強度 Kgf/cm2")
    tensile_strength_diff = round(
        (100 * (before_tensile_strength - after_tensile_strength) / before_tensile_strength),
        2,
    )

    return {
        "product_name": "AY8000RB",
        "date": test_date,
        "lot_no": lot_no,
        "ay8000r_quant": format_numeric_text(parse_positive_int(values.get("ay8000r_quantity"), "AY8000R數量")),
        "ay8000b_quant": format_numeric_text(parse_positive_int(values.get("ay8000b_quantity"), "AY8000B數量")),
        "hy8000_quant": format_numeric_text(parse_positive_int(values.get("hy8000_quantity"), "HY8000數量")),
        "due_date": time.strftime("%Y/%m/%d", due_date.timetuple()),
        "ay8000r_viscosity": format_numeric_text(parse_positive_int(values.get("ay8000r_viscosity"), "AY8000R 黏度 cPs")),
        "ay8000b_viscosity": format_numeric_text(parse_positive_int(values.get("ay8000b_viscosity"), "AY8000B 黏度 cPs")),
        "hy8000_viscosity": format_numeric_text(
            "{:.1f}".format(parse_positive_float(values.get("hy8000_viscosity"), "HY8000 黏度 cPs"))
        ),
        "ay8000r_gel_time": format_numeric_text(parse_positive_int(values.get("ay8000r_gel_time"), "AY8000R 凝膠時間 sec")),
        "ay8000b_gel_time": format_numeric_text(parse_positive_int(values.get("ay8000b_gel_time"), "AY8000B 凝膠時間 sec")),
        "before_tensile_strength": format_numeric_text(before_tensile_strength),
        "after_tensile_strength": format_numeric_text(after_tensile_strength),
        "tensile_strength_diff": format_numeric_text(tensile_strength_diff),
        "acid_resistance": format_numeric_text(
            "{:.2f}".format(parse_positive_float(values.get("acid_resistance"), "耐酸性 %"))
        ),
    }
