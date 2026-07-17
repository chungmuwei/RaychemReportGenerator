from __future__ import annotations

import os
import time
import traceback
from tkinter import Tk, filedialog, messagebox, ttk
import tkinter as tk
from tkinter import font as tkfont

from . import generator
from .coa_config import (
    BUSWAY_PRODUCT_NAME,
    BUSWAY_TEMPLATE_FILE,
    DEBUG,
    ETACOM_PRODUCT_NAME,
    ETACOM_QTY_PRODUCTS,
    ETACOM_TEMPLATE_FILE,
    UIC_PRODUCT_NAME,
    UIC_QTY_PRODUCTS,
    UIC_TEMPLATE_FILE,
    YUASA_TEMPLATE_FILE,
    export_paths,
    load_export_paths,
    load_product_specs,
    save_all_paths,
)
from .coa_context import (
    build_type_1_context,
    build_yuasa_context,
    type_1_uses_user_quantity,
)
from .coa_utils import (
    UserInputError,
    format_numeric_text,
    format_type_1_viscosity,
    mousewheel_scroll_units,
)


class ScrollableFrame(ttk.Frame):
    """A themed frame whose contents can scroll vertically."""

    def __init__(self, parent):
        """Initialize the canvas, scrollbar, and inner content frame."""
        super().__init__(parent)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.body = ttk.Frame(self.canvas, padding=10)
        self.canvas_window = self.canvas.create_window(
            (0, 0),
            window=self.body,
            anchor="nw",
        )
        self.body.columnconfigure(0, weight=1)

        self.body.bind("<Configure>", self.update_scroll_region)
        self.canvas.bind("<Configure>", self.update_canvas_width)

    def update_scroll_region(self, _event=None):
        """Resize the canvas scroll region to include all child widgets."""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def update_canvas_width(self, event):
        """Keep the inner frame width synchronized with the canvas."""
        self.canvas.itemconfigure(self.canvas_window, width=event.width)

    def scroll_units(self, units: int):
        """Scroll the content vertically by the requested number of units."""
        self.canvas.yview_scroll(units, "units")


class COAApp:
    """Desktop interface for validating inputs and exporting COA reports."""

    def __init__(
        self,
        root: Tk,
        product_specs: dict | None = None,
        report_generator=generator.generate_coa_report,
        dialog=messagebox,
        file_dialog=filedialog,
    ):
        """Initialize application state, window behavior, and widgets."""
        self.root = root
        self.product_specs = (
            product_specs if product_specs is not None else load_product_specs()
        )
        self.report_generator = report_generator
        self.dialog = dialog
        self.file_dialog = file_dialog
        self.vars: dict[str, tk.StringVar] = {}
        self.listboxes: dict[str, tk.Listbox] = {}
        self.labels: dict[str, ttk.Label] = {}
        self.entries: dict[str, ttk.Entry] = {}
        self.company_frames: dict[str, ttk.Frame] = {}
        self.company_buttons: dict[str, ttk.Button] = {}
        self.active_company: str | None = None

        self.root.title("瑞肯材料品檢報告產生器")
        self.root.geometry("720x500")
        self.root.minsize(660, 460)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.report_callback_exception = self.report_callback_exception

        self.configure_style()
        self.build_layout()

    def configure_style(self):
        """Configure shared fonts and themed widget styles."""
        default_font = tkfont.nametofont("TkDefaultFont")
        default_font.configure(size=16)
        text_font = tkfont.nametofont("TkTextFont")
        text_font.configure(size=16)
        style = ttk.Style(self.root)
        style.configure("TButton", padding=(8, 4))
        style.configure("CompanyTab.TButton", padding=(16, 6))
        style.configure("SelectedCompanyTab.TButton", padding=(16, 6))
        style.configure("TLabel", padding=(0, 2))
        style.configure("TFrame", padding=0)

    def build_layout(self):
        """Build company tabs, report forms, and scrolling behavior."""
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        mainframe = ttk.Frame(self.root, padding=(10, 10, 10, 10))
        mainframe.grid(row=0, column=0, sticky="nsew")
        mainframe.columnconfigure(0, weight=1)
        mainframe.rowconfigure(1, weight=1)

        self.tab_bar = ttk.Frame(mainframe)
        self.tab_bar.grid(row=0, column=0, sticky="w", pady=(0, 8))

        self.scroll_frame = ScrollableFrame(mainframe)
        self.scroll_frame.grid(row=1, column=0, sticky="nsew")

        etacom = self.add_company_tab("etacom", "安達康", 0)
        self.build_type_1_section(
            etacom,
            company="etacom",
            products=ETACOM_PRODUCT_NAME,
            default_product="樹脂CY2536L",
            visible_products=4,
            template=ETACOM_TEMPLATE_FILE,
            qty_products=ETACOM_QTY_PRODUCTS,
        )

        busway = self.add_company_tab("busway", "巴斯威爾", 1)
        self.build_type_1_section(
            busway,
            company="busway",
            products=BUSWAY_PRODUCT_NAME,
            default_product="CY2533L7",
            visible_products=2,
            template=BUSWAY_TEMPLATE_FILE,
        )

        uic = self.add_company_tab("uic", "盛英", 2)
        self.build_type_1_section(
            uic,
            company="uic",
            products=UIC_PRODUCT_NAME,
            default_product="CY8101R",
            visible_products=2,
            template=UIC_TEMPLATE_FILE,
            qty_products=UIC_QTY_PRODUCTS,
            include_gel_time=False,
        )

        yuasa = self.add_company_tab("yuasa", "湯淺", 3)
        self.build_yuasa_section(yuasa)
        self.switch_company("etacom")
        self.root.bind_all("<MouseWheel>", self.on_mousewheel)
        self.root.bind_all("<Button-4>", self.on_scroll_up)
        self.root.bind_all("<Button-5>", self.on_scroll_down)

    def add_company_tab(self, company: str, title: str, column: int) -> ttk.Frame:
        """Add a company selector and return its associated content frame."""
        button = ttk.Button(
            self.tab_bar,
            text=title,
            style="CompanyTab.TButton",
            command=lambda: self.switch_company(company),
        )
        button.grid(row=0, column=column, sticky="w", padx=(0, 4))
        frame = ttk.Frame(self.scroll_frame.body)
        frame.columnconfigure(0, weight=1)
        self.company_buttons[company] = button
        self.company_frames[company] = frame
        return frame

    def switch_company(self, company: str):
        """Display one company's form and hide the previously active form."""
        if company == self.active_company:
            return
        if self.active_company:
            self.company_frames[self.active_company].grid_remove()
            self.company_buttons[self.active_company].configure(
                style="CompanyTab.TButton"
            )
        self.company_frames[company].grid(row=0, column=0, sticky="nw")
        self.company_buttons[company].configure(style="SelectedCompanyTab.TButton")
        self.active_company = company
        self.scroll_frame.canvas.yview_moveto(0)
        self.root.after_idle(self.scroll_frame.update_scroll_region)

    def on_mousewheel(self, event):
        """Scroll the active form in response to mouse-wheel input."""
        units = mousewheel_scroll_units(event.delta)
        if units:
            self.scroll_frame.scroll_units(units)
        return "break"

    def on_scroll_up(self, _event):
        """Scroll the active form upward for Linux button events."""
        self.scroll_frame.scroll_units(-1)
        return "break"

    def on_scroll_down(self, _event):
        """Scroll the active form downward for Linux button events."""
        self.scroll_frame.scroll_units(1)
        return "break"

    def build_type_1_section(
        self,
        parent,
        company: str,
        products: list[str],
        default_product: str,
        visible_products: int,
        template: str,
        qty_products: set[str] | None = None,
        include_gel_time: bool = True,
    ):
        """Build the shared form used by type-1 company reports."""
        listbox = self.add_listbox(
            parent,
            0,
            "品名",
            f"{company}_product_name",
            products,
            default_product,
            visible_products,
        )
        self.add_entry(parent, 1, "批號", f"{company}_lot_no", "T")
        next_row = 2
        if qty_products:
            self.add_entry(parent, next_row, "數量", f"{company}_quantity")
            listbox.bind(
                "<<ListboxSelect>>",
                lambda _event, c=company, qp=qty_products: (
                    self.update_type_1_quantity_state(c, qp)
                ),
            )
            self.update_type_1_quantity_state(company, qty_products)
            next_row += 1
        self.add_entry(parent, next_row, "黏度 cPs", f"{company}_viscosity")
        next_row += 1
        if include_gel_time:
            self.add_entry(parent, next_row, "凝膠時間 sec", f"{company}_gel_time")
            next_row += 1
        self.add_entry(
            parent,
            next_row,
            "檢測日期(YYYY/MM/DD)",
            f"{company}_date",
            time.strftime("%Y/%m/%d"),
        )
        ttk.Button(
            parent,
            text="輸出報告",
            command=self.safe_callback(
                lambda: self.export_type_1_report(
                    company,
                    template,
                    qty_products,
                    include_gel_time,
                )
            ),
        ).grid(row=next_row + 1, column=1, sticky="w", pady=(8, 0))

    def build_yuasa_section(self, parent):
        """Build the Yuasa-specific report input form."""
        self.add_entry(parent, 0, "批號", "yuasa_lot_no", "T")

        ay8000r = ttk.LabelFrame(parent, text="AY8000R", padding=(8, 6))
        ay8000r.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 2))
        self.add_entry(ay8000r, 0, "AY8000R重量 Kg", "ay8000r_weight")
        self.add_entry(ay8000r, 1, "黏度 cPs", "ay8000r_viscosity")
        self.add_entry(ay8000r, 2, "凝膠時間 sec", "ay8000r_gel_time")

        ay8000b = ttk.LabelFrame(parent, text="AY8000B", padding=(8, 6))
        ay8000b.grid(row=2, column=0, columnspan=2, sticky="ew", pady=2)
        self.add_entry(ay8000b, 0, "AY8000B重量 Kg", "ay8000b_weight")
        self.add_entry(ay8000b, 1, "黏度 cPs", "ay8000b_viscosity")
        self.add_entry(ay8000b, 2, "凝膠時間 sec", "ay8000b_gel_time")

        hy8000 = ttk.LabelFrame(parent, text="HY8000", padding=(8, 6))
        hy8000.grid(row=3, column=0, columnspan=2, sticky="ew", pady=2)
        self.add_entry(hy8000, 0, "HY8000重量 Kg", "hy8000_weight")
        self.add_entry(hy8000, 1, "黏度 cPs", "hy8000_viscosity")

        self.add_entry(
            parent, 4, "浸酸前引張強度 Kgf/cm2", "before_tensile_strength"
        )
        self.add_entry(
            parent, 5, "浸酸後引張強度 Kgf/cm2", "after_tensile_strength"
        )
        self.add_entry(parent, 6, "耐酸性 %", "acid_resistance")
        self.add_entry(
            parent,
            7,
            "檢測日期(YYYY/MM/DD)",
            "yuasa_date",
            time.strftime("%Y/%m/%d"),
        )
        ttk.Button(
            parent,
            text="輸出報告",
            command=self.safe_callback(
                lambda: self.export_yuasa_report(YUASA_TEMPLATE_FILE)
            ),
        ).grid(row=8, column=1, sticky="w", pady=(8, 0))

    def add_entry(self, parent, row: int, label: str, key: str, default: str = ""):
        """Add a labeled text entry and register its state by key."""
        label_widget = ttk.Label(parent, text=label)
        label_widget.grid(row=row, column=0, sticky="w", padx=(0, 8), pady=2)
        var = tk.StringVar(value=default)
        entry = ttk.Entry(parent, textvariable=var, width=26)
        entry.grid(row=row, column=1, sticky="w", pady=2)
        self.vars[key] = var
        self.labels[key] = label_widget
        self.entries[key] = entry
        return entry

    def add_listbox(
        self,
        parent,
        row: int,
        label: str,
        key: str,
        items: list[str],
        default: str,
        visible_items: int,
    ):
        """Add a labeled list box, populate it, and select its default item."""
        ttk.Label(parent, text=label).grid(
            row=row, column=0, sticky="nw", padx=(0, 8), pady=2
        )
        listbox = tk.Listbox(
            parent,
            height=visible_items,
            width=26,
            exportselection=False,
        )
        for item in items:
            listbox.insert(tk.END, item)
        if default in items:
            index = items.index(default)
            listbox.selection_set(index)
            listbox.see(index)
        listbox.grid(row=row, column=1, sticky="w", pady=2)
        self.listboxes[key] = listbox
        return listbox

    def get_listbox_value(self, key: str) -> str:
        """Return the selected list-box value or an empty string."""
        selection = self.listboxes[key].curselection()
        if not selection:
            return ""
        return self.listboxes[key].get(selection[0])

    def update_type_1_quantity_state(self, company: str, qty_products: set[str]):
        """Enable quantity input only for products that require it."""
        key = f"{company}_quantity"
        label = self.labels.get(key)
        entry = self.entries.get(f"{company}_quantity")
        if label is None or entry is None:
            return
        selected_product = self.get_listbox_value(f"{company}_product_name")
        state = "normal" if selected_product in qty_products else "disabled"
        label.configure(state=state)
        entry.configure(state=state)

    def get_type_1_values(
        self,
        company: str,
        qty_products: set[str] | None = None,
        include_gel_time: bool = True,
    ) -> dict:
        """Collect the current type-1 form values for context building."""
        product_name = self.get_listbox_value(f"{company}_product_name")
        values = {
            "product_name": product_name,
            "date": self.vars[f"{company}_date"].get(),
            "lot_no": self.vars[f"{company}_lot_no"].get(),
            "viscosity": self.vars[f"{company}_viscosity"].get(),
        }
        if type_1_uses_user_quantity(product_name, qty_products):
            values["quantity"] = self.vars[f"{company}_quantity"].get()
        if include_gel_time:
            values["gel_time"] = self.vars[f"{company}_gel_time"].get()
        return values

    def get_yuasa_values(self) -> dict:
        """Collect and normalize the current Yuasa form values."""
        keys = [
            "yuasa_lot_no",
            "ay8000r_weight",
            "ay8000r_viscosity",
            "ay8000r_gel_time",
            "ay8000b_weight",
            "ay8000b_viscosity",
            "ay8000b_gel_time",
            "hy8000_weight",
            "hy8000_viscosity",
            "before_tensile_strength",
            "after_tensile_strength",
            "acid_resistance",
            "yuasa_date",
        ]
        values = {key: self.vars[key].get() for key in keys}
        values["lot_no"] = values.pop("yuasa_lot_no")
        values["date"] = values.pop("yuasa_date")
        return values

    def ask_output_directory(self, company: str, title: str) -> str:
        """Ask for and persist a company's report output directory."""
        output_dir = self.file_dialog.askdirectory(
            initialdir=export_paths[company],
            title=title,
            mustexist=True,
        )
        if not output_dir:
            return ""
        export_paths[company] = output_dir
        save_all_paths(export_paths)
        return output_dir

    def export_type_1_report(
        self,
        company: str,
        template: str,
        qty_products: set[str] | None = None,
        include_gel_time: bool = True,
    ):
        """Validate, generate, and announce a type-1 report export."""
        context = build_type_1_context(
            company,
            self.get_type_1_values(company, qty_products, include_gel_time),
            self.product_specs,
            qty_products,
            include_gel_time,
        )
        output_dir = self.ask_output_directory(company, "選擇報告輸出資料夾")
        if not output_dir:
            return
        filename = self.generate_report(template, context, output_dir)
        self.show_message(
            "成功",
            f"報告 {os.path.basename(filename)} 已成功匯出至 {output_dir}！",
        )

    def export_yuasa_report(self, template: str):
        """Validate, generate, and announce a Yuasa report export."""
        context = build_yuasa_context(self.get_yuasa_values())
        output_dir = self.ask_output_directory("yuasa", "選擇報告輸出資料夾")
        if not output_dir:
            return
        filename = self.generate_report(template, context, output_dir)
        self.show_message(
            "成功",
            f"報告 {os.path.basename(filename)} 已成功匯出至 {output_dir}！",
        )

    def generate_report(self, template: str, context: dict, output_dir: str) -> str:
        """Generate a report and normalize generator failures for the GUI."""
        try:
            return self.report_generator(
                template_file=template,
                context=context,
                output_path=output_dir,
            )
        except Exception as exc:
            raise RuntimeError(f"匯出失敗：\n{exc}") from exc

    def show_message(self, title: str, message: str):
        """Display an informational dialog."""
        self.dialog.showinfo(title, message)

    def show_error(self, message: str):
        """Display an error dialog."""
        self.dialog.showerror("錯誤", message)

    def safe_callback(self, callback):
        """Wrap a GUI callback with user-facing exception handling."""

        def wrapped():
            """Run the callback and convert known failures into dialogs."""
            try:
                return callback()
            except UserInputError as exc:
                self.show_error(str(exc))
            except RuntimeError as exc:
                self.show_error(str(exc))
            except Exception as exc:
                if DEBUG:
                    traceback.print_exc()
                self.show_error(f"發生未預期錯誤：\n{exc}")

        return wrapped

    def report_callback_exception(self, exc_type, exc_value, exc_traceback):
        """Handle exceptions raised directly by Tkinter callbacks."""
        if DEBUG:
            traceback.print_exception(exc_type, exc_value, exc_traceback)
        self.show_error(f"發生未預期錯誤：\n{exc_value}")

    def on_close(self):
        """Persist settings before closing the root window."""
        try:
            save_all_paths(export_paths)
        except Exception as exc:
            self.show_error(f"儲存設定失敗：\n{exc}")
            return
        self.root.destroy()


def run():
    """Start the Tkinter desktop application."""
    load_export_paths()
    root = Tk()
    COAApp(root)
    root.mainloop()


if __name__ == "__main__":
    run()
