import os
import tempfile
import unittest

from raychem_report_generator import generator


class GeneratorTests(unittest.TestCase):
    def test_sequence_filename_returns_docx_when_file_does_not_exist(self):
        """Return the base DOCX path when no report exists."""
        with tempfile.TemporaryDirectory() as tmpdir:
            base_path = os.path.join(tmpdir, "COA_Test_T260614")

            self.assertEqual(
                generator.sequence_filename(base_path),
                base_path + generator.DOCX_FILE_EXTENSION,
            )

    def test_sequence_filename_renames_existing_file_and_returns_next_suffix(self):
        """Preserve the first report and return the second sequenced path."""
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
        """Resolve data resources relative to the project root."""
        path = generator.resource_path("product_specs.json")

        self.assertTrue(path.endswith("product_specs.json"))
        self.assertTrue(os.path.exists(path))


if __name__ == "__main__":
    unittest.main()
