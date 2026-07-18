import os
import tempfile
import unittest
from types import SimpleNamespace

from raychem_report_generator import app


PRODUCT_SPECS = {
    "etacom": {
        "樹脂CY2536L": {
            "weight": "1300KG",
            "viscosity_range": "900~1500",
            "appearance": "淡黃色透明液體",
            "hardness": ">70",
            "gel_time_range": "40~80",
        },
        "硬化劑HY2537": {
            "weight": "180KG",
            "viscosity_range": "20~50",
            "appearance": "透明無雜質液體",
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
    "uic": {
        "CY8101R": {
            "unit_weight": 20,
            "viscosity_range": "10,000~20,000",
            "appearance": "Red liquid",
        },
        "HY8101": {
            "unit_weight": 15,
            "viscosity_range": "15~50",
            "appearance": "Transparent liquid",
        },
    },
}


class DummyDialog:
    def __init__(self):
        """Initialize captured informational and error dialogs.

        Returns:
            None.
        """
        self.infos = []
        self.errors = []

    def showinfo(self, title, message):
        """Capture an informational dialog call.

        Args:
            title: Dialog title supplied by the application.
            message: Dialog message supplied by the application.

        Returns:
            None.
        """
        self.infos.append((title, message))

    def showerror(self, title, message):
        """Capture an error dialog call.

        Args:
            title: Dialog title supplied by the application.
            message: Dialog message supplied by the application.

        Returns:
            None.
        """
        self.errors.append((title, message))


class FakeVar:
    def __init__(self, value):
        """Store a test value.

        Args:
            value: Value returned by :meth:`get`.

        Returns:
            None.
        """
        self.value = value

    def get(self):
        """Return the stored test value.

        Returns:
            The value supplied during initialization.
        """
        return self.value


class FakeWidget:
    def __init__(self):
        """Initialize an empty widget configuration.

        Returns:
            None.
        """
        self.config = {}

    def configure(self, **kwargs):
        """Capture widget configuration changes.

        Args:
            **kwargs: Widget options applied by the application.

        Returns:
            None.
        """
        self.config.update(kwargs)


class AppValidationTests(unittest.TestCase):
    def test_build_type_1_context_formats_and_populates_specs(self):
        """Build a type-1 context with formatted specification values.

        Returns:
            None.
        """
        context = app.build_type_1_context(
            "etacom",
            {
                "product_name": "樹脂CY2536L",
                "date": "2026/06/14",
                "lot_no": "T260614",
                "viscosity": "950.25",
                "gel_time": "55",
            },
            PRODUCT_SPECS,
        )

        self.assertEqual(context["product_name"], "樹脂CY2536L")
        self.assertEqual(context["date"], "2026/06/14")
        self.assertEqual(context["weight"], "1,300KG")
        self.assertEqual(context["couple"], "HY2536")
        self.assertEqual(context["viscosity"], "950.2")
        self.assertEqual(context["gel_time"], "55")

    def test_build_type_1_context_rejects_missing_lot_no(self):
        """Reject type-1 inputs without a lot number.

        Returns:
            None.
        """
        with self.assertRaisesRegex(app.UserInputError, "批號不可空白"):
            app.build_type_1_context(
                "etacom",
                {
                    "product_name": "樹脂CY2536L",
                    "date": "2026/06/14",
                    "lot_no": "",
                    "viscosity": "950",
                    "gel_time": "55",
                },
                PRODUCT_SPECS,
            )

    def test_build_type_1_context_rejects_invalid_date(self):
        """Reject type-1 inputs with a nonexistent calendar date.

        Returns:
            None.
        """
        with self.assertRaisesRegex(
            app.UserInputError,
            "檢測日期不是有效日期",
        ):
            app.build_type_1_context(
                "etacom",
                {
                    "product_name": "樹脂CY2536L",
                    "date": "2026/02/30",
                    "lot_no": "T260614",
                    "viscosity": "950",
                    "gel_time": "55",
                },
                PRODUCT_SPECS,
            )

    def test_build_type_1_context_rejects_compact_date_format(self):
        """Reject compact dates that do not use slash separators.

        Returns:
            None.
        """
        with self.assertRaisesRegex(app.UserInputError, "YYYY/MM/DD"):
            app.build_type_1_context(
                "etacom",
                {
                    "product_name": "樹脂CY2536L",
                    "date": "20260614",
                    "lot_no": "T260614",
                    "viscosity": "950",
                    "gel_time": "55",
                },
                PRODUCT_SPECS,
            )

    def test_build_uic_context_populates_template_fields_without_gel_time(self):
        """Build UIC fields without requiring gel time.

        Returns:
            None.
        """
        context = app.build_type_1_context(
            "uic",
            {
                "product_name": "CY8101R",
                "date": "2026/06/14",
                "lot_no": "T260528",
                "total_weight": "480",
                "viscosity": "12158",
            },
            PRODUCT_SPECS,
            include_total_weight=True,
            include_gel_time=False,
        )

        self.assertEqual(context["product_name"], "CY8101R")
        self.assertEqual(context["weight"], "480")
        self.assertEqual(context["unit_weight"], "20")
        self.assertEqual(context["qty"], "24")
        self.assertEqual(context["viscosity_range"], "10,000~20,000")
        self.assertEqual(context["appearance"], "Red liquid")
        self.assertEqual(context["obs_appearance"], "Red liquid")
        self.assertEqual(context["viscosity"], "12,158")
        self.assertNotIn("gel_time", context)

    def test_build_uic_context_rejects_weight_not_divisible_by_unit_weight(self):
        """Reject a UIC total weight that cannot form whole packages.

        Returns:
            None.
        """
        with self.assertRaisesRegex(app.UserInputError, "20 Kg 的倍數"):
            app.build_type_1_context(
                "uic",
                {
                    "product_name": "CY8101R",
                    "date": "2026/06/14",
                    "lot_no": "T260528",
                    "viscosity": "12158",
                    "total_weight": "74",
                },
                PRODUCT_SPECS,
                include_total_weight=True,
                include_gel_time=False,
            )

    def test_build_etacom_hy2537_context_uses_entered_qty(self):
        """Use the entered quantity for Etacom HY2537.

        Returns:
            None.
        """
        context = app.build_type_1_context(
            "etacom",
            {
                "product_name": "硬化劑HY2537",
                "date": "2026/06/14",
                "lot_no": "T260614",
                "viscosity": "30",
                "quantity": "3",
                "gel_time": "55",
            },
            PRODUCT_SPECS,
            app.ETACOM_QTY_PRODUCTS,
        )

        self.assertEqual(context["weight"], "180KG")
        self.assertEqual(context["qty"], "3")

    def test_build_etacom_hy2537_context_requires_qty_when_enabled(self):
        """Require quantity for Etacom HY2537.

        Returns:
            None.
        """
        with self.assertRaisesRegex(app.UserInputError, "數量不可空白"):
            app.build_type_1_context(
                "etacom",
                {
                    "product_name": "硬化劑HY2537",
                    "date": "2026/06/14",
                    "lot_no": "T260614",
                    "viscosity": "30",
                    "quantity": "",
                    "gel_time": "55",
                },
                PRODUCT_SPECS,
                app.ETACOM_QTY_PRODUCTS,
            )

    def test_uic_context_renders_uic_template(self):
        """Render a UIC context with the configured template.

        Returns:
            None.
        """
        context = app.build_type_1_context(
            "uic",
            {
                "product_name": "HY8101",
                "date": "2026/06/14",
                "lot_no": "T260528",
                "total_weight": "135",
                "viscosity": "30",
            },
            PRODUCT_SPECS,
            include_total_weight=True,
            include_gel_time=False,
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            output = app.generator.generate_coa_report(
                app.UIC_TEMPLATE_FILE,
                context,
                tmpdir,
            )

            self.assertTrue(os.path.exists(output))
            self.assertTrue(output.endswith(".docx"))

    def test_build_yuasa_context_calculates_due_date_and_tensile_diff(self):
        """Calculate Yuasa derived dates, quantities, and strength values.

        Returns:
            None.
        """
        context = app.build_yuasa_context(
            {
                "date": "2026/06/14",
                "lot_no": "T260101",
                "ay8000r_weight": "200",
                "ay8000r_viscosity": "1258",
                "ay8000r_gel_time": "60",
                "ay8000b_weight": "200",
                "ay8000b_viscosity": "130",
                "ay8000b_gel_time": "70",
                "hy8000_weight": "100",
                "hy8000_viscosity": "88.5",
                "before_tensile_strength": "100",
                "after_tensile_strength": "92",
                "acid_resistance": "98.234",
            }
        )

        self.assertEqual(context["product_name"], "AY8000RB")
        self.assertEqual(context["date"], "2026/06/14")
        self.assertEqual(context["due_date"], "2026/07/01")
        self.assertEqual(context["ay8000r_quant"], "50")
        self.assertEqual(context["ay8000r_viscosity"], "1,258")
        self.assertEqual(context["hy8000_viscosity"], "88.5")
        self.assertEqual(context["tensile_strength_diff"], "8.0")
        self.assertEqual(context["acid_resistance"], "98.23")

    def test_build_yuasa_context_rejects_invalid_lot_number(self):
        """Reject a Yuasa lot number with an invalid shape.

        Returns:
            None.
        """
        with self.assertRaisesRegex(app.UserInputError, "湯淺批號格式錯誤"):
            app.build_yuasa_context(
                {
                    "date": "2026/06/14",
                    "lot_no": "T26",
                    "ay8000r_weight": "100",
                    "ay8000r_viscosity": "120",
                    "ay8000r_gel_time": "60",
                    "ay8000b_weight": "200",
                    "ay8000b_viscosity": "130",
                    "ay8000b_gel_time": "70",
                    "hy8000_weight": "100",
                    "hy8000_viscosity": "88.5",
                    "before_tensile_strength": "100",
                    "after_tensile_strength": "92",
                    "acid_resistance": "98.234",
                }
            )

    def test_touchpad_scroll_uses_tk_precise_pixel_scrolling(self):
        """Unpack the Tk 9 touchpad delta and scroll the canvas by pixels."""
        calls = []
        frame = SimpleNamespace(
            canvas=SimpleNamespace(_w=".canvas"),
            tk=SimpleNamespace(call=lambda *args: calls.append(args) or (2, -7)),
        )

        result = app.ScrollableFrame.on_touchpad_scroll(
            frame, SimpleNamespace(delta=196601)
        )

        self.assertEqual(calls[0], ("tk::PreciseScrollDeltas", 196601))
        self.assertEqual(calls[1], ("tk::ScrollByPixels", ".canvas", 0, -7))
        self.assertEqual(result, "break")

    def test_mousewheel_uses_tk_9_mousewheel_handler(self):
        """Delegate physical mouse-wheel movement to Tk 9."""
        calls = []
        frame = SimpleNamespace(
            canvas=SimpleNamespace(_w=".canvas"),
            tk=SimpleNamespace(call=lambda *args: calls.append(args)),
        )

        result = app.ScrollableFrame.on_mousewheel(
            frame, SimpleNamespace(delta=120)
        )

        self.assertEqual(calls, [("tk::MouseWheel", ".canvas", "y", 120, -40.0)])
        self.assertEqual(result, "break")

    def test_type_1_viscosity_format_avoids_forced_trailing_zeroes(self):
        """Format viscosity without unnecessary trailing zeroes.

        Returns:
            None.
        """
        self.assertEqual(app.format_type_1_viscosity(30), "30")
        self.assertEqual(app.format_type_1_viscosity(950.25), "950.2")
        self.assertEqual(app.format_type_1_viscosity(12158), "12,158")

    def test_numeric_text_format_adds_commas_without_changing_decimals(self):
        """Add thousands separators while preserving decimal text.

        Returns:
            None.
        """
        self.assertEqual(app.format_numeric_text("1300KG"), "1,300KG")
        self.assertEqual(app.format_numeric_text("1000~1500"), "1,000~1,500")
        self.assertEqual(app.format_numeric_text("1234.5"), "1,234.5")


class AppCallbackTests(unittest.TestCase):
    def make_uninitialized_app(self):
        """Build a minimal application object for callback tests.

        Returns:
            A partially initialized application with a dummy dialog provider.
        """
        app_obj = object.__new__(app.COAApp)
        app_obj.dialog = DummyDialog()
        return app_obj

    def test_safe_callback_shows_user_input_error_popup(self):
        """Show validation errors through the dialog adapter.

        Returns:
            None.
        """
        app_obj = self.make_uninitialized_app()
        callback = app.COAApp.safe_callback(
            app_obj,
            lambda: (_ for _ in ()).throw(
                app.UserInputError("黏度 cPs必須是數字。")
            ),
        )

        callback()

        self.assertEqual(
            app_obj.dialog.errors,
            [("錯誤", "黏度 cPs必須是數字。")],
        )

    def test_safe_callback_shows_unexpected_runtime_error_popup(self):
        """Show unexpected callback errors through the dialog adapter.

        Returns:
            None.
        """
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
        """Export type-1 reports to the selected native directory.

        Returns:
            None.
        """
        calls = []
        with tempfile.TemporaryDirectory() as tmpdir:
            app_obj = self.make_uninitialized_app()
            app_obj.product_specs = PRODUCT_SPECS
            app_obj.ask_output_directory = lambda company, title: tmpdir
            app_obj.get_type_1_values = (
                lambda company, qty_products=None, include_total_weight=False,
                include_gel_time=True: {
                "product_name": "樹脂CY2536L",
                "date": "2026/06/14",
                "lot_no": "T260614",
                "viscosity": "950",
                "gel_time": "55",
                }
            )

            def fake_generator(template_file, context, output_path):
                """Capture report generator arguments and return a fake path.

                Args:
                    template_file: Template path passed by the application.
                    context: Rendering context passed by the application.
                    output_path: Destination passed by the application.

                Returns:
                    A synthetic report filename in the destination directory.
                """
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
        """Validate type-1 inputs before asking for an output directory.

        Returns:
            None.
        """
        app_obj = self.make_uninitialized_app()
        app_obj.product_specs = PRODUCT_SPECS
        app_obj.get_type_1_values = (
            lambda company, qty_products=None, include_total_weight=False,
            include_gel_time=True: {
            "product_name": "樹脂CY2536L",
            "date": "2026/06/14",
            "lot_no": "",
            "viscosity": "950",
            "gel_time": "55",
            }
        )
        app_obj.ask_output_directory = lambda company, title: self.fail(
            "directory dialog should not open"
        )
        app_obj.report_generator = lambda **kwargs: self.fail(
            "report should not generate"
        )

        callback = app.COAApp.safe_callback(
            app_obj,
            lambda: app.COAApp.export_type_1_report(app_obj, "etacom", "template.docx"),
        )
        callback()

        self.assertEqual(app_obj.dialog.errors, [("錯誤", "批號不可空白。")])
        self.assertEqual(app_obj.dialog.infos, [])

    def test_uic_export_does_not_require_gel_time(self):
        """Export UIC reports without a gel-time input.

        Returns:
            None.
        """
        calls = []
        with tempfile.TemporaryDirectory() as tmpdir:
            app_obj = self.make_uninitialized_app()
            app_obj.product_specs = PRODUCT_SPECS
            app_obj.ask_output_directory = lambda company, title: tmpdir
            app_obj.get_type_1_values = (
                lambda company, qty_products=None, include_total_weight=False,
                include_gel_time=True: {
                "product_name": "HY8101",
                "date": "2026/06/14",
                "lot_no": "T260528",
                "viscosity": "30",
                "total_weight": "135",
                }
            )

            def fake_generator(template_file, context, output_path):
                """Capture UIC generator arguments and return a fake path.

                Args:
                    template_file: Template path passed by the application.
                    context: Rendering context passed by the application.
                    output_path: Destination passed by the application.

                Returns:
                    A synthetic UIC report filename.
                """
                calls.append((template_file, context, output_path))
                return os.path.join(output_path, "COA_HY8101_T260528.docx")

            app_obj.report_generator = fake_generator

            app.COAApp.export_type_1_report(
                app_obj,
                "uic",
                "template.docx",
                include_total_weight=True,
                include_gel_time=False,
            )

        self.assertEqual(calls[0][1]["product_name"], "HY8101")
        self.assertEqual(calls[0][1]["unit_weight"], "15")
        self.assertEqual(calls[0][1]["qty"], "9")
        self.assertNotIn("gel_time", calls[0][1])
        self.assertEqual(app_obj.dialog.infos[0][0], "成功")

    def test_yuasa_validation_error_shows_before_directory_dialog(self):
        """Validate Yuasa inputs before asking for an output directory.

        Returns:
            None.
        """
        app_obj = self.make_uninitialized_app()
        app_obj.get_yuasa_values = lambda: {
            "date": "2026/06/14",
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
        app_obj.ask_output_directory = lambda company, title: self.fail(
            "directory dialog should not open"
        )
        app_obj.report_generator = lambda **kwargs: self.fail(
            "report should not generate"
        )

        callback = app.COAApp.safe_callback(
            app_obj,
            lambda: app.COAApp.export_yuasa_report(app_obj, "template.docx"),
        )
        callback()

        self.assertEqual(len(app_obj.dialog.errors), 1)
        self.assertIn("湯淺批號格式錯誤", app_obj.dialog.errors[0][1])
        self.assertEqual(app_obj.dialog.infos, [])


class AppWidgetStateTests(unittest.TestCase):
    def make_quantity_state_app(self, product_name):
        """Build a minimal application object for quantity-state tests.

        Args:
            product_name: Product returned by the fake list-box accessor.

        Returns:
            A partially initialized application with fake quantity widgets.
        """
        app_obj = object.__new__(app.COAApp)
        app_obj.labels = {"etacom_quantity": FakeWidget()}
        app_obj.entries = {"etacom_quantity": FakeWidget()}
        app_obj.get_listbox_value = lambda key: product_name
        return app_obj

    def test_etacom_quantity_label_and_entry_disable_for_non_qty_product(self):
        """Disable quantity controls for products with fixed quantities.

        Returns:
            None.
        """
        app_obj = self.make_quantity_state_app("樹脂CY2536L")

        app.COAApp.update_type_1_quantity_state(
            app_obj,
            "etacom",
            app.ETACOM_QTY_PRODUCTS,
        )

        self.assertEqual(app_obj.labels["etacom_quantity"].config["state"], "disabled")
        self.assertEqual(app_obj.entries["etacom_quantity"].config["state"], "disabled")

    def test_etacom_quantity_label_and_entry_enable_for_hy2537(self):
        """Enable quantity controls for Etacom HY2537.

        Returns:
            None.
        """
        app_obj = self.make_quantity_state_app("硬化劑HY2537")

        app.COAApp.update_type_1_quantity_state(
            app_obj,
            "etacom",
            app.ETACOM_QTY_PRODUCTS,
        )

        self.assertEqual(app_obj.labels["etacom_quantity"].config["state"], "normal")
        self.assertEqual(app_obj.entries["etacom_quantity"].config["state"], "normal")

    def test_type_1_values_only_collects_qty_for_selected_qty_product(self):
        """Collect quantity only for products configured to require it.

        Returns:
            None.
        """
        app_obj = object.__new__(app.COAApp)
        app_obj.get_listbox_value = lambda key: "樹脂CY2536L"
        app_obj.vars = {
            "etacom_date": FakeVar("2026/06/14"),
            "etacom_lot_no": FakeVar("T260614"),
            "etacom_viscosity": FakeVar("950"),
            "etacom_quantity": FakeVar("3"),
            "etacom_gel_time": FakeVar("55"),
        }

        values = app.COAApp.get_type_1_values(
            app_obj,
            "etacom",
            app.ETACOM_QTY_PRODUCTS,
        )

        self.assertNotIn("quantity", values)


if __name__ == "__main__":
    unittest.main()
