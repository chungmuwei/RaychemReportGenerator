import os
import tempfile
import unittest
from unittest.mock import patch

from docx import Document
from docx.oxml.ns import qn

from raychem_report_generator import generator
from raychem_report_generator.coa_config import (
    BUSWAY_TEMPLATE_FILE,
    ETACOM_TEMPLATE_FILE,
    UIC_TEMPLATE_FILE,
)


class GeneratorTests(unittest.TestCase):
    def test_generated_product_name_uses_east_asian_and_latin_fonts(self):
        context = {"product_name": "樹脂CY2536L", "lot_no": "T261128"}

        for template in (ETACOM_TEMPLATE_FILE, BUSWAY_TEMPLATE_FILE, UIC_TEMPLATE_FILE):
            with self.subTest(template=template), tempfile.TemporaryDirectory() as tmpdir:
                with patch("builtins.print"):
                    path = generator.generate_coa_report(template, context, tmpdir)

                self.assertEqual(os.path.basename(path), "COA_CY2536L_T261128.docx")
                document = Document(path)
                product_fonts = [
                    run.xpath("./w:rPr/w:rFonts")[0]
                    for run in document.element.xpath(".//w:r")
                    if run.text == context["product_name"]
                ]
                self.assertEqual(len(product_fonts), 1)
                self.assertEqual(product_fonts[0].get(qn("w:eastAsia")), "DFKai-SB")
                self.assertEqual(product_fonts[0].get(qn("w:ascii")), "Arial")
                self.assertEqual(product_fonts[0].get(qn("w:hAnsi")), "Arial")

    def test_sequence_filename_returns_docx_when_file_does_not_exist(self):
        """Return the base DOCX path when no report exists.

        Returns:
            None.
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            base_path = os.path.join(tmpdir, "COA_Test_T260614")

            self.assertEqual(
                generator.sequence_filename(base_path),
                base_path + generator.DOCX_FILE_EXTENSION,
            )

    def test_sequence_filename_renames_existing_file_and_returns_next_suffix(self):
        """Preserve the first report and return the second sequenced path.

        Returns:
            None.
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            base_path = os.path.join(tmpdir, "COA_Test_T260614")
            existing = base_path + generator.DOCX_FILE_EXTENSION
            with open(existing, "w", encoding="utf-8") as existing_file:
                existing_file.write("old")

            next_path = generator.sequence_filename(base_path)

            self.assertEqual(
                next_path,
                base_path + "-2" + generator.DOCX_FILE_EXTENSION,
            )
            self.assertFalse(os.path.exists(existing))
            self.assertTrue(
                os.path.exists(base_path + "-1" + generator.DOCX_FILE_EXTENSION)
            )

    def test_resource_path_resolves_project_relative_file(self):
        """Resolve data resources relative to the project root.

        Returns:
            None.
        """
        path = generator.resource_path("product_specs.json")

        self.assertTrue(path.endswith("product_specs.json"))
        self.assertTrue(os.path.exists(path))


if __name__ == "__main__":
    unittest.main()
