import os
import tempfile
import unittest

import app


PRODUCT_SPECS = {
    "etacom": {
        "樹脂CY2536L": {
            "weight": "1300KG",
            "viscosity_range": "900~1500",
            "appearance": "淡黃色透明液體",
            "hardness": ">70",
            "gel_time_range": "40~80",
        }
    },
    "busway": {
        "CY2533L7": {
            "weight": "1100KG",
            "viscosity_range": "1000~1500",
            "appearance": "Light yellow transparent liquid",
            "hardness": ">70",
            "gel_time_range": "60~100",
        }
    },
}


class DummyDialog:
    def __init__(self):
        self.infos = []
        self.errors = []

    def showinfo(self, title, message):
        self.infos.append((title, message))

    def showerror(self, title, message):
        self.errors.append((title, message))


class AppValidationTests(unittest.TestCase):
    def test_build_type_1_context_formats_and_populates_specs(self):
        context = app.build_type_1_context(
            "etacom",
            {
                "product_name": "樹脂CY2536L",
                "date": "20260614",
                "lot_no": "T260614",
                "viscosity": "950.25",
                "gel_time": "55",
            },
            PRODUCT_SPECS,
        )

        self.assertEqual(context["product_name"], "樹脂CY2536L")
        self.assertEqual(context["date"], "20260614")
        self.assertEqual(context["weight"], "1300KG")
        self.assertEqual(context["couple"], "HY2536")
        self.assertEqual(context["viscosity"], "950.2")
        self.assertEqual(context["gel_time"], 55)

    def test_build_type_1_context_rejects_missing_lot_no(self):
        with self.assertRaisesRegex(app.UserInputError, "批號不可空白"):
            app.build_type_1_context(
                "etacom",
                {
                    "product_name": "樹脂CY2536L",
                    "date": "20260614",
                    "lot_no": "",
                    "viscosity": "950",
                    "gel_time": "55",
                },
                PRODUCT_SPECS,
            )

    def test_build_type_1_context_rejects_invalid_date(self):
        with self.assertRaisesRegex(app.UserInputError, "檢測日期不是有效日期"):
            app.build_type_1_context(
                "etacom",
                {
                    "product_name": "樹脂CY2536L",
                    "date": "20260230",
                    "lot_no": "T260614",
                    "viscosity": "950",
                    "gel_time": "55",
                },
                PRODUCT_SPECS,
            )

    def test_build_yuasa_context_calculates_due_date_and_tensile_diff(self):
        context = app.build_yuasa_context(
            {
                "date": "20260614",
                "lot_no": "T260101",
                "ay8000r_quantity": "1",
                "ay8000r_viscosity": "120",
                "ay8000r_gel_time": "60",
                "ay8000b_quantity": "2",
                "ay8000b_viscosity": "130",
                "ay8000b_gel_time": "70",
                "hy8000_quantity": "3",
                "hy8000_viscosity": "88.5",
                "before_tensile_strength": "100",
                "after_tensile_strength": "92",
                "acid_resistance": "98.234",
            }
        )

        self.assertEqual(context["product_name"], "AY8000RB")
        self.assertEqual(context["due_date"], "2026-07-01")
        self.assertEqual(context["hy8000_viscosity"], "88.5")
        self.assertEqual(context["tensile_strength_diff"], 8.0)
        self.assertEqual(context["acid_resistance"], "98.23")

    def test_build_yuasa_context_rejects_invalid_lot_number(self):
        with self.assertRaisesRegex(app.UserInputError, "湯淺批號格式錯誤"):
            app.build_yuasa_context(
                {
                    "date": "20260614",
                    "lot_no": "T26",
                    "ay8000r_quantity": "1",
                    "ay8000r_viscosity": "120",
                    "ay8000r_gel_time": "60",
                    "ay8000b_quantity": "2",
                    "ay8000b_viscosity": "130",
                    "ay8000b_gel_time": "70",
                    "hy8000_quantity": "3",
                    "hy8000_viscosity": "88.5",
                    "before_tensile_strength": "100",
                    "after_tensile_strength": "92",
                    "acid_resistance": "98.234",
                }
            )

    def test_mousewheel_scroll_units_handles_trackpad_and_wheel_deltas(self):
        self.assertEqual(app.mousewheel_scroll_units(120), -1)
        self.assertEqual(app.mousewheel_scroll_units(-120), 1)
        self.assertEqual(app.mousewheel_scroll_units(1), -1)
        self.assertEqual(app.mousewheel_scroll_units(-1), 1)
        self.assertEqual(app.mousewheel_scroll_units(0), 0)


class AppCallbackTests(unittest.TestCase):
    def make_uninitialized_app(self):
        app_obj = object.__new__(app.COAApp)
        app_obj.dialog = DummyDialog()
        return app_obj

    def test_safe_callback_shows_user_input_error_popup(self):
        app_obj = self.make_uninitialized_app()
        callback = app.COAApp.safe_callback(
            app_obj,
            lambda: (_ for _ in ()).throw(app.UserInputError("黏度 cPs必須是數字。")),
        )

        callback()

        self.assertEqual(app_obj.dialog.errors, [("錯誤", "黏度 cPs必須是數字。")])

    def test_safe_callback_shows_unexpected_runtime_error_popup(self):
        app_obj = self.make_uninitialized_app()
        callback = app.COAApp.safe_callback(
            app_obj,
            lambda: (_ for _ in ()).throw(Exception("boom")),
        )

        callback()

        self.assertEqual(len(app_obj.dialog.errors), 1)
        self.assertIn("發生未預期錯誤", app_obj.dialog.errors[0][1])
        self.assertIn("boom", app_obj.dialog.errors[0][1])

    def test_export_type_1_report_uses_native_directory_and_generator(self):
        calls = []
        with tempfile.TemporaryDirectory() as tmpdir:
            app_obj = self.make_uninitialized_app()
            app_obj.product_specs = PRODUCT_SPECS
            app_obj.ask_output_directory = lambda company, title: tmpdir
            app_obj.get_type_1_values = lambda company: {
                "product_name": "樹脂CY2536L",
                "date": "20260614",
                "lot_no": "T260614",
                "viscosity": "950",
                "gel_time": "55",
            }

            def fake_generator(template_file, context, output_path):
                calls.append((template_file, context, output_path))
                return os.path.join(output_path, "COA_test.docx")

            app_obj.report_generator = fake_generator

            app.COAApp.export_type_1_report(app_obj, "etacom", "template.docx")

        self.assertEqual(calls[0][0], "template.docx")
        self.assertEqual(calls[0][1]["lot_no"], "T260614")
        self.assertTrue(calls[0][2])
        self.assertEqual(app_obj.dialog.infos[0][0], "成功")
        self.assertIn("COA_test.docx", app_obj.dialog.infos[0][1])

    def test_export_type_1_validation_error_shows_before_directory_dialog(self):
        app_obj = self.make_uninitialized_app()
        app_obj.product_specs = PRODUCT_SPECS
        app_obj.get_type_1_values = lambda company: {
            "product_name": "樹脂CY2536L",
            "date": "20260614",
            "lot_no": "",
            "viscosity": "950",
            "gel_time": "55",
        }
        app_obj.ask_output_directory = lambda company, title: self.fail("directory dialog should not open")
        app_obj.report_generator = lambda **kwargs: self.fail("report should not generate")

        callback = app.COAApp.safe_callback(
            app_obj,
            lambda: app.COAApp.export_type_1_report(app_obj, "etacom", "template.docx"),
        )
        callback()

        self.assertEqual(app_obj.dialog.errors, [("錯誤", "批號不可空白。")])
        self.assertEqual(app_obj.dialog.infos, [])

    def test_yuasa_validation_error_shows_before_directory_dialog(self):
        app_obj = self.make_uninitialized_app()
        app_obj.get_yuasa_values = lambda: {
            "date": "20260614",
            "lot_no": "T26",
            "ay8000r_quantity": "1",
            "ay8000r_viscosity": "120",
            "ay8000r_gel_time": "60",
            "ay8000b_quantity": "2",
            "ay8000b_viscosity": "130",
            "ay8000b_gel_time": "70",
            "hy8000_quantity": "3",
            "hy8000_viscosity": "88.5",
            "before_tensile_strength": "100",
            "after_tensile_strength": "92",
            "acid_resistance": "98.234",
        }
        app_obj.ask_output_directory = lambda company, title: self.fail("directory dialog should not open")
        app_obj.report_generator = lambda **kwargs: self.fail("report should not generate")

        callback = app.COAApp.safe_callback(
            app_obj,
            lambda: app.COAApp.export_yuasa_report(app_obj, "template.docx"),
        )
        callback()

        self.assertEqual(len(app_obj.dialog.errors), 1)
        self.assertIn("湯淺批號格式錯誤", app_obj.dialog.errors[0][1])
        self.assertEqual(app_obj.dialog.infos, [])


if __name__ == "__main__":
    unittest.main()
