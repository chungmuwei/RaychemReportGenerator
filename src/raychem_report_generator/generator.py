"""Generate Certificate of Analysis reports from Word templates."""

import os
import re
import sys

from docxtpl import DocxTemplate

TEST_EXPORT_PATH = "/Users/raymond/Desktop/code/python/RaychemReportGenerator/output"

DOCX_FILE_EXTENSION = ".docx"


def resource_path(relative_path: str) -> str:
    """Return an absolute resource path for source and PyInstaller runs."""
    package_root = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(package_root, os.pardir, os.pardir))
    base_path = getattr(sys, "_MEIPASS", project_root)
    return os.path.join(base_path, relative_path)


def generate_coa_report(
    template_file: str,
    context: dict[str, str],
    output_path: str | None = None,
) -> str:
    """Render and save a COA report, returning its sequenced filename."""
    print(
        "Generate COA report with template file: "
        f"{template_file} and context: {context}"
    )

    template = DocxTemplate(template_file=resource_path(template_file))
    template.render(context=context)

    product_name = context["product_name"]
    lot_no = context["lot_no"]
    target_directory = output_path if output_path else TEST_EXPORT_PATH
    safe_product_name = re.sub(r"[^a-zA-Z0-9]", "", product_name)
    filepath = os.path.join(
        target_directory,
        f"COA_{safe_product_name}_{lot_no}",
    )
    filepath = sequence_filename(filepath)
    print(f"Export docx at {filepath}")
    template.save(filename=filepath)
    return filepath


def sequence_filename(path: str) -> str:
    """Return an available filename while preserving an existing first file."""
    first_path = f"{path}{DOCX_FILE_EXTENSION}"
    if os.path.exists(first_path):
        os.rename(first_path, f"{path}-1{DOCX_FILE_EXTENSION}")
        return f"{path}-2{DOCX_FILE_EXTENSION}"

    order = 2
    sequenced_path = f"{path}-{order}{DOCX_FILE_EXTENSION}"
    while os.path.exists(sequenced_path):
        order += 1
        sequenced_path = f"{path}-{order}{DOCX_FILE_EXTENSION}"

    if order > 2:
        return sequenced_path
    return first_path
