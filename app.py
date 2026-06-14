from __future__ import annotations

import json
import os
import re
import time
import traceback
from datetime import date, datetime
from tkinter import Tk, filedialog, messagebox, ttk
import tkinter as tk
from tkinter import font as tkfont

from dateutil.relativedelta import relativedelta

import generator

DEBUG = False

# ETACOM_TEMPLATE_FILE = generator.resource_path("templates/COA_Etacom_template.docx")
ETACOM_TEMPLATE_FILE = generator.resource_path("templates/COA_Etacom_template_font_revision.docx")  # 新細明體
BUSWAY_TEMPLATE_FILE = generator.resource_path("templates/COA_Busway_template.docx")
YUASA_TEMPLATE_FILE = generator.resource_path("templates/COA_Yuasa_template.docx")
ETACOM_PRODUCT_NAME = ["樹脂CY2536L", "樹脂CY2536", "硬化劑HY2536", "硬化劑HY2537"]
BUSWAY_PRODUCT_NAME = ["CY2533L7", "HY2533"]
COUPLE = {
    "樹脂CY2536L": "HY2536",
    "樹脂CY2536": "HY2536",
    "硬化劑HY2536": "CY2536L",
    "硬化劑HY2537": "CY2536L",
    "CY2533L7": "HY2533",
    "HY2533": "CY2533L7",
}

# PATHS
APP_SUPPORT_DIR = os.path.join(
    os.path.expanduser("~"),
    "Library",
    "Application Support",
    "com.raychemcoa",
)
CONFIG_FILE = os.path.join(APP_SUPPORT_DIR, "export_config.json")
PRODUCT_SPECS_FILE = generator.resource_path("product_specs.json")
DEFAULT_EXPORT_PATH = os.path.expanduser("~")
export_paths = {"etacom": DEFAULT_EXPORT_PATH, "busway": DEFAULT_EXPORT_PATH, "yuasa": DEFAULT_EXPORT_PATH}


class UserInputError(ValueError):
    """Raised when user-entered GUI values cannot be used to generate a report."""


def create_export_config_file():
    """Create config file at APP_SUPPORT_DIR/CONFIG_FILE to store export paths if it doesn't exist."""
    if DEBUG:
        print(f"Creating export config file at {CONFIG_FILE} with default paths: {export_paths}")
    os.makedirs(APP_SUPPORT_DIR, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(obj=export_paths, fp=f, ensure_ascii=False)


def load_last_path(company: str):
    """Read last export path for a specific company from config file."""
    if not os.path.exists(CONFIG_FILE):
        create_export_config_file()

    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            saved_path = data.get(f"{company}_export_path", DEFAULT_EXPORT_PATH)
            if saved_path and os.path.isdir(saved_path):
                return saved_path
    except Exception:
        return DEFAULT_EXPORT_PATH
    return DEFAULT_EXPORT_PATH


def save_all_paths(paths: dict):
    """Save export paths for all companies to config file."""
    if not os.path.exists(CONFIG_FILE):
        create_export_config_file()

    data = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            pass
    for company, path in paths.items():
        if path and os.path.isdir(path):
            data[f"{company}_export_path"] = path
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        if DEBUG:
            print(f"Saved export config paths: {paths}")
    except Exception as e:
        if DEBUG:
            print(f"Failed to save export config paths: {str(e)}")


def load_product_specs(path: str = PRODUCT_SPECS_FILE) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


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
    if not re.fullmatch(r"\d{8}", text):
        raise UserInputError(f"{label}格式必須為 YYYYMMDD。")
    try:
        datetime.strptime(text, "%Y%m%d")
    except ValueError as exc:
        raise UserInputError(f"{label}不是有效日期。") from exc
    return text


def format_type_1_viscosity(viscosity: float) -> str | int:
    return f"{viscosity:#.4g}" if viscosity < 1000 else round(viscosity)


def mousewheel_scroll_units(delta: int) -> int:
    if delta == 0:
        return 0
    if abs(delta) >= 120:
        return int(-delta / 120)
    return -1 if delta > 0 else 1


def build_type_1_context(company: str, values: dict, product_specs: dict) -> dict:
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
    gel_time = parse_positive_int(values.get("gel_time"), "凝膠時間 sec")
    spec = product_specs[company][product_name]

    return {
        "product_name": product_name,
        "date": test_date,
        "lot_no": lot_no,
        "weight": spec["weight"],
        "viscosity_range": spec["viscosity_range"],
        "appearance": spec["appearance"],
        "obs_appearance": spec["appearance"],
        "couple": COUPLE[product_name],
        "hardness": spec["hardness"],
        "gel_time_range": spec["gel_time_range"],
        "viscosity": format_type_1_viscosity(viscosity),
        "gel_time": gel_time,
    }


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
        "ay8000r_quant": parse_positive_int(values.get("ay8000r_quantity"), "AY8000R數量"),
        "ay8000b_quant": parse_positive_int(values.get("ay8000b_quantity"), "AY8000B數量"),
        "hy8000_quant": parse_positive_int(values.get("hy8000_quantity"), "HY8000數量"),
        "due_date": time.strftime("%Y-%m-%d", due_date.timetuple()),
        "ay8000r_viscosity": parse_positive_int(values.get("ay8000r_viscosity"), "AY8000R 黏度 cPs"),
        "ay8000b_viscosity": parse_positive_int(values.get("ay8000b_viscosity"), "AY8000B 黏度 cPs"),
        "hy8000_viscosity": "{:.1f}".format(parse_positive_float(values.get("hy8000_viscosity"), "HY8000 黏度 cPs")),
        "ay8000r_gel_time": parse_positive_int(values.get("ay8000r_gel_time"), "AY8000R 凝膠時間 sec"),
        "ay8000b_gel_time": parse_positive_int(values.get("ay8000b_gel_time"), "AY8000B 凝膠時間 sec"),
        "before_tensile_strength": before_tensile_strength,
        "after_tensile_strength": after_tensile_strength,
        "tensile_strength_diff": tensile_strength_diff,
        "acid_resistance": "{:.2f}".format(parse_positive_float(values.get("acid_resistance"), "耐酸性 %")),
    }


class ScrollableFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.body = ttk.Frame(self.canvas, padding=10)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.body, anchor="nw")
        self.body.columnconfigure(0, weight=1)

        self.body.bind("<Configure>", self.update_scroll_region)
        self.canvas.bind("<Configure>", self.update_canvas_width)

    def update_scroll_region(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def update_canvas_width(self, event):
        self.canvas.itemconfigure(self.canvas_window, width=event.width)

    def scroll_units(self, units: int):
        self.canvas.yview_scroll(units, "units")


class COAApp:
    def __init__(
        self,
        root: Tk,
        product_specs: dict | None = None,
        report_generator=generator.generate_coa_report,
        dialog=messagebox,
        file_dialog=filedialog,
    ):
        self.root = root
        self.product_specs = product_specs if product_specs is not None else load_product_specs()
        self.report_generator = report_generator
        self.dialog = dialog
        self.file_dialog = file_dialog
        self.vars: dict[str, tk.StringVar] = {}
        self.listboxes: dict[str, tk.Listbox] = {}
        self.company_frames: dict[str, ttk.Frame] = {}
        self.company_buttons: dict[str, ttk.Button] = {}
        self.active_company: str | None = None

        self.root.title("瑞肯材料品檢報告產生器")
        self.root.geometry("900x600")
        self.root.minsize(760, 520)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.report_callback_exception = self.report_callback_exception

        self.configure_style()
        self.build_layout()

    def configure_style(self):
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
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        outer = ttk.Frame(self.root, padding=10)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        self.tab_bar = ttk.Frame(outer)
        self.tab_bar.grid(row=0, column=0, sticky="w", pady=(0, 8))

        self.scroll_frame = ScrollableFrame(outer)
        self.scroll_frame.grid(row=1, column=0, sticky="nsew")

        etacom = self.add_company_tab("etacom", "安達康", 0)
        self.build_type_1_section(
            etacom,
            company="etacom",
            products=ETACOM_PRODUCT_NAME,
            default_product="樹脂CY2536L",
            visible_products=3,
            template=ETACOM_TEMPLATE_FILE,
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

        yuasa = self.add_company_tab("yuasa", "湯淺", 2)
        self.build_yuasa_section(yuasa)
        self.switch_company("etacom")
        self.root.bind_all("<MouseWheel>", self.on_mousewheel)
        self.root.bind_all("<Button-4>", self.on_scroll_up)
        self.root.bind_all("<Button-5>", self.on_scroll_down)

    def add_company_tab(self, company: str, title: str, column: int) -> ttk.Frame:
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
        if company == self.active_company:
            return
        if self.active_company:
            self.company_frames[self.active_company].grid_remove()
            self.company_buttons[self.active_company].configure(style="CompanyTab.TButton")
        self.company_frames[company].grid(row=0, column=0, sticky="nw")
        self.company_buttons[company].configure(style="SelectedCompanyTab.TButton")
        self.active_company = company
        self.scroll_frame.canvas.yview_moveto(0)
        self.root.after_idle(self.scroll_frame.update_scroll_region)

    def on_mousewheel(self, event):
        units = mousewheel_scroll_units(event.delta)
        if units:
            self.scroll_frame.scroll_units(units)
        return "break"

    def on_scroll_up(self, _event):
        self.scroll_frame.scroll_units(-1)
        return "break"

    def on_scroll_down(self, _event):
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
    ):
        self.add_listbox(parent, 0, "品名", f"{company}_product_name", products, default_product, visible_products)
        self.add_entry(parent, 1, "批號", f"{company}_lot_no", "T")
        self.add_entry(parent, 2, "黏度 cPs", f"{company}_viscosity")
        self.add_entry(parent, 3, "凝膠時間 sec", f"{company}_gel_time")
        self.add_entry(parent, 4, "檢測日期(YYYYMMDD)", f"{company}_date", time.strftime("%Y%m%d"))
        ttk.Button(
            parent,
            text="輸出報告",
            command=self.safe_callback(lambda: self.export_type_1_report(company, template)),
        ).grid(row=5, column=1, sticky="w", pady=(8, 0))

    def build_yuasa_section(self, parent):
        self.add_entry(parent, 0, "批號", "yuasa_lot_no", "T")

        ay8000r = ttk.LabelFrame(parent, text="AY8000R", padding=(8, 6))
        ay8000r.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 2))
        self.add_entry(ay8000r, 0, "AY8000R數量", "ay8000r_quantity")
        self.add_entry(ay8000r, 1, "黏度 cPs", "ay8000r_viscosity")
        self.add_entry(ay8000r, 2, "凝膠時間 sec", "ay8000r_gel_time")

        ay8000b = ttk.LabelFrame(parent, text="AY8000B", padding=(8, 6))
        ay8000b.grid(row=2, column=0, columnspan=2, sticky="ew", pady=2)
        self.add_entry(ay8000b, 0, "AY8000B數量", "ay8000b_quantity")
        self.add_entry(ay8000b, 1, "黏度 cPs", "ay8000b_viscosity")
        self.add_entry(ay8000b, 2, "凝膠時間 sec", "ay8000b_gel_time")

        hy8000 = ttk.LabelFrame(parent, text="HY8000", padding=(8, 6))
        hy8000.grid(row=3, column=0, columnspan=2, sticky="ew", pady=2)
        self.add_entry(hy8000, 0, "HY8000數量", "hy8000_quantity")
        self.add_entry(hy8000, 1, "黏度 cPs", "hy8000_viscosity")

        self.add_entry(parent, 4, "浸酸前引張強度 Kgf/cm2", "before_tensile_strength")
        self.add_entry(parent, 5, "浸酸後引張強度 Kgf/cm2", "after_tensile_strength")
        self.add_entry(parent, 6, "耐酸性 %", "acid_resistance")
        self.add_entry(parent, 7, "檢測日期(YYYYMMDD)", "yuasa_date", time.strftime("%Y%m%d"))
        ttk.Button(
            parent,
            text="輸出報告",
            command=self.safe_callback(lambda: self.export_yuasa_report(YUASA_TEMPLATE_FILE)),
        ).grid(row=8, column=1, sticky="w", pady=(8, 0))

    def add_entry(self, parent, row: int, label: str, key: str, default: str = ""):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=2)
        var = tk.StringVar(value=default)
        entry = ttk.Entry(parent, textvariable=var, width=26)
        entry.grid(row=row, column=1, sticky="w", pady=2)
        self.vars[key] = var
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
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="nw", padx=(0, 8), pady=2)
        listbox = tk.Listbox(parent, height=visible_items, width=26, exportselection=False)
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
        selection = self.listboxes[key].curselection()
        if not selection:
            return ""
        return self.listboxes[key].get(selection[0])

    def get_type_1_values(self, company: str) -> dict:
        return {
            "product_name": self.get_listbox_value(f"{company}_product_name"),
            "date": self.vars[f"{company}_date"].get(),
            "lot_no": self.vars[f"{company}_lot_no"].get(),
            "viscosity": self.vars[f"{company}_viscosity"].get(),
            "gel_time": self.vars[f"{company}_gel_time"].get(),
        }

    def get_yuasa_values(self) -> dict:
        keys = [
            "yuasa_lot_no",
            "ay8000r_quantity",
            "ay8000r_viscosity",
            "ay8000r_gel_time",
            "ay8000b_quantity",
            "ay8000b_viscosity",
            "ay8000b_gel_time",
            "hy8000_quantity",
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
        output_dir = self.file_dialog.askdirectory(initialdir=export_paths[company], title=title, mustexist=True)
        if not output_dir:
            return ""
        export_paths[company] = output_dir
        save_all_paths(export_paths)
        return output_dir

    def export_type_1_report(self, company: str, template: str):
        context = build_type_1_context(company, self.get_type_1_values(company), self.product_specs)
        output_dir = self.ask_output_directory(company, "選擇報告輸出資料夾")
        if not output_dir:
            return
        filename = self.generate_report(template, context, output_dir)
        self.show_message("成功", f"報告 {os.path.basename(filename)} 已成功匯出至 {output_dir}！")

    def export_yuasa_report(self, template: str):
        context = build_yuasa_context(self.get_yuasa_values())
        output_dir = self.ask_output_directory("yuasa", "選擇報告輸出資料夾")
        if not output_dir:
            return
        filename = self.generate_report(template, context, output_dir)
        self.show_message("成功", f"報告 {os.path.basename(filename)} 已成功匯出至 {output_dir}！")

    def generate_report(self, template: str, context: dict, output_dir: str) -> str:
        try:
            return self.report_generator(template_file=template, context=context, output_path=output_dir)
        except Exception as exc:
            raise RuntimeError(f"匯出失敗：\n{exc}") from exc

    def show_message(self, title: str, message: str):
        self.dialog.showinfo(title, message)

    def show_error(self, message: str):
        self.dialog.showerror("錯誤", message)

    def safe_callback(self, callback):
        def wrapped():
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
        if DEBUG:
            traceback.print_exception(exc_type, exc_value, exc_traceback)
        self.show_error(f"發生未預期錯誤：\n{exc_value}")

    def on_close(self):
        try:
            save_all_paths(export_paths)
        except Exception as exc:
            self.show_error(f"儲存設定失敗：\n{exc}")
            return
        self.root.destroy()


def run():
    export_paths["etacom"] = load_last_path("etacom")
    export_paths["busway"] = load_last_path("busway")
    export_paths["yuasa"] = load_last_path("yuasa")
    root = Tk()
    COAApp(root)
    root.mainloop()


if __name__ == "__main__":
    run()
