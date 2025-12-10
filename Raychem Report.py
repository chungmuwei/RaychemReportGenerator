import dearpygui.dearpygui as dpg
import generator
import time
from datetime import date
from dateutil.relativedelta import relativedelta
import os  
import json

ETACOM_TEMPLATE_FILE = generator.resource_path("templates/COA_Etacom_template.docx")
BUSWAY_TEMPLATE_FILE = generator.resource_path("templates/COA_Busway_template.docx")
YUASA_TEMPLATE_FILE = generator.resource_path("templates/COA_Yuasa_template.docx")
ETACOM_PRODUCT_NAME = ["樹脂CY2536L", "硬化劑HY2536", "硬化劑HY2537"]
BUSWAY_PRODUCT_NAME = ["CY2533L7", "HY2533"]

# PATHS
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".raychem_report_config.json")
PRODUCT_SPECS_FILE = generator.resource_path("product_specs.json")
DEFAULT_EXPORT_PATH = os.path.expanduser("~")
EXPORT_PATH = "/Volumes/Business/steven_20200721/1 備份 20200410/工作/A_ISO續評/2022複評/表單/P003生產流程/瑞肯COA" 

def load_last_path():
    """Read last export path from config file"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                data = json.load(f)
                return data.get("last_export_path", DEFAULT_EXPORT_PATH)
        except:
            return DEFAULT_EXPORT_PATH
    return DEFAULT_EXPORT_PATH

def save_last_path(path: str):
    """Save last export path to config file"""
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump({"last_export_path": path}, f)
    except:
        pass

# 新增顯示訊息的函式
def show_message(title, message, window_width=450, window_height=150, centred=True):
    
    # 計算置中位置 (Viewport 寬高 / 2 - 視窗寬高 / 2)
    if centred:
        vp_width = dpg.get_viewport_width()
        vp_height = dpg.get_viewport_height()
        pos_x = (vp_width // 2) - (window_width // 2)
        pos_y = (vp_height // 2) - (window_height // 2)

    # 建立一個唯一的 tag，避免重複開啟時衝突
    window_tag = f"msg_win_{time.time()}"
    
    with dpg.window(label=title, modal=True, show=True, tag=window_tag, 
                    width=window_width, height=window_height,
                    pos=[pos_x, pos_y] if centred else [100, 100], popup=True):
        dpg.add_text(message, wrap=window_width - 20)
        dpg.add_separator()
        # 關閉視窗的按鈕
        dpg.add_button(label="確定", width=75, callback=lambda: dpg.delete_item(window_tag))

def export_type_1_coa_report(sender, app_data, user_data):
    """
    Export Certificate of Analysis report word document for Etacon and Busway
    """

    # print(f"sender is: {sender}")
    # print(f"app_data is: {app_data}")
    # print(f"user_data is: {user_data}")

    output_dir = app_data["file_path_name"]
    save_last_path(output_dir)
    
    with open(PRODUCT_SPECS_FILE, 'r') as f:
        product_specs = json.load(f)

    # get value from user
    company = user_data["company"]
    template = user_data["template"]

    product_name = dpg.get_value(company+"_product_name")
    lot_no = dpg.get_value(company+"_lot_no")
    viscosity = dpg.get_value(company+"_viscosity")
    gel_time = dpg.get_value(company+"_gel_time")

    context = {
        "product_name": product_name,
        "date": time.strftime("%Y/%m/%d"),
        "lot_no": lot_no,
        "weight": product_specs[company][product_name]["weight"],
        "viscosity_range": product_specs[company][product_name]["viscosity_range"],
        "appearance": product_specs[company][product_name]["appearance"],
        "obs_appearance": product_specs[company][product_name]["appearance"],
        "hardness": product_specs[company][product_name]["hardness"],
        "gel_time_range": product_specs[company][product_name]["gel_time_range"],
        "viscosity": viscosity,
        "gel_time": gel_time
    }

    try:
        filename = generator.generate_coa_report(template_file=template, context=context, output_path=output_dir)
        show_message("成功", f"報告 {os.path.basename(filename)} 已成功匯出至 {output_dir}！")
    except Exception as e:
        show_message("錯誤", f"匯出失敗：\n{str(e)}")

def export_yuasa_coa_report(sender, app_data, user_data):
    """
    Export Certificate of Analysis report word document for Yuasa
    """

    # get output path
    output_dir = app_data["file_path_name"]
    save_last_path(output_dir)

    # get value from user
    company = user_data["company"]
    template = user_data["template"]

    lot_no = dpg.get_value("yuasa_lot_no")
    year = 2000 + int(lot_no[1:3])
    month = int(lot_no[3:5])
    day = int(lot_no[5:7])
    due_date = date(year, month, day) + relativedelta(month=6)

    ay8000r_quantity = dpg.get_value("ay8000r_quantity")
    ay8000r_viscosity = dpg.get_value("ay8000r_viscosity")
    ay8000r_gel_time = dpg.get_value("ay8000r_gel_time")

    ay8000b_quantity = dpg.get_value("ay8000b_quantity")
    ay8000b_viscosity = dpg.get_value("ay8000b_viscosity")
    ay8000b_gel_time = dpg.get_value("ay8000b_gel_time")

    hy8000_quantity = dpg.get_value("hy8000_quantity")
    hy8000_viscosity = dpg.get_value("hy8000_viscosity")

    before_tensile_strength = dpg.get_value("before_tensile_strength")
    after_tensile_strength = dpg.get_value("after_tensile_strength")
    tensile_strength_diff = round((100 * (before_tensile_strength - after_tensile_strength) / before_tensile_strength), 2)
    acid_resistance = dpg.get_value("acid_resistance")

    context = {
        "product_name": "AY8000RB",
        "date": time.strftime("%Y-%m-%d"),
        "lot_no": lot_no,
        "ay8000r_quant": ay8000r_quantity,
        "ay8000b_quant": ay8000b_quantity,
        "hy8000_quant": hy8000_quantity,
        "due_date": time.strftime("%Y-%m-%d", due_date.timetuple()),
        "ay8000r_viscosity": ay8000r_viscosity,
        "ay8000b_viscosity": ay8000b_viscosity,
        "hy8000_viscosity": "{:.1f}".format(hy8000_viscosity),
        "ay8000r_gel_time": ay8000r_gel_time,
        "ay8000b_gel_time": ay8000b_gel_time,
        "before_tensile_strength": before_tensile_strength,
        "after_tensile_strength": after_tensile_strength,
        "tensile_strength_diff": tensile_strength_diff,
        "acid_resistance": "{:.2f}".format(acid_resistance)
    }
    print(context)
    
    try:
        filename = generator.generate_coa_report(template_file=template, context=context, output_path=output_dir)
        # get the file name from the full path
        show_message("成功", f"報告 {os.path.basename(filename)} 已成功匯出至 {output_dir}！")
    except Exception as e:
        print(str(e))
        show_message("錯誤", f"匯出失敗：\n{str(e)}")

def show_file_dialog(sender, app_data, user_data):
    """
    顯示檔案選擇對話框，並確保預設路徑是上次使用的路徑。
    
    Args:
        sender: 觸發此 callback 的元件 (按鈕)
        app_data: unused
        user_data: 要顯示的檔案對話框 tag (str)
    """
    dialog_tag = user_data
    
    # 讀取最新的路徑 (從 config.json)
    current_path = load_last_path()
    
    # 更新對話框的預設路徑
    dpg.configure_item(dialog_tag, default_path=current_path)
    
    # 顯示對話框
    dpg.show_item(dialog_tag) 

def run():

    dpg.create_context()

    ################################################################
    #                           Fonts                              #
    ################################################################
    
    # add chinese font
    with dpg.font_registry():
        # normal size font
        with dpg.font(generator.resource_path("fonts/Noto_Sans_TC/NotoSansTC-Regular.otf"), 24) as zh_font:
            dpg.add_font_range_hint(dpg.mvFontRangeHint_Default)
            dpg.add_font_range_hint(dpg.mvFontRangeHint_Chinese_Full)
        # normal size bold font 
        with dpg.font(generator.resource_path("fonts/Noto_Sans_TC/NotoSansTC-Bold.otf"), 24) as zh_bold_font:
            dpg.add_font_range_hint(dpg.mvFontRangeHint_Default)
            dpg.add_font_range_hint(dpg.mvFontRangeHint_Chinese_Full)
        # header size font
        with dpg.font(generator.resource_path("fonts/Noto_Sans_TC/NotoSansTC-Regular.otf"), 32) as zh_header_font:
            dpg.add_font_range_hint(dpg.mvFontRangeHint_Default)
            dpg.add_font_range_hint(dpg.mvFontRangeHint_Chinese_Full)

    # dpg.show_font_manager()

    # dpg.show_style_editor()
    current_export_path = load_last_path()
    ################################################################
    #                          Windows                             #
    ################################################################
    with dpg.window(label="Example Window", tag="Primary Window"):
        # 安達康
        with dpg.collapsing_header(label="安達康"):
            dpg.add_listbox(label="品名", tag="etacom_product_name", default_value="樹脂CY2536L", items=ETACOM_PRODUCT_NAME, num_items=3)
            dpg.add_input_text(label="批號", tag="etacom_lot_no", default_value="T")
            dpg.add_input_int(label="黏度 cPs", tag="etacom_viscosity")
            dpg.add_input_int(label="凝膠時間 sec", tag="etacom_gel_time")
            
            dpg.add_file_dialog(label="輸出安達康報告", tag="etacom_file_dialog", 
                directory_selector=True, show=False, default_path=current_export_path, 
                callback=export_type_1_coa_report, user_data={"company": "etacom", 
                "template": ETACOM_TEMPLATE_FILE}, height=500)
            dpg.add_button(label="輸出報告", tag="etacom_export_button", callback=lambda: dpg.show_item("etacom_file_dialog"), user_data="etacom_file_dialog") 
        
        # 巴斯威爾
        with dpg.collapsing_header(label="巴斯威爾"):
            dpg.add_listbox(label="品名", tag="busway_product_name", default_value="CY2533L7", items=BUSWAY_PRODUCT_NAME, num_items=2)
            dpg.add_input_text(label="批號", tag="busway_lot_no", default_value="T")
            dpg.add_input_int(label="黏度 cPs", tag="busway_viscosity")
            dpg.add_input_int(label="凝膠時間 sec", tag="busway_gel_time")

            dpg.add_file_dialog(label="輸出巴斯威爾報告", tag="busway_file_dialog", 
                directory_selector=True, show=False, default_path=current_export_path, 
                callback=export_type_1_coa_report, user_data={"company": "busway", 
                "template": BUSWAY_TEMPLATE_FILE}, height=500)
            dpg.add_button(label="輸出報告", tag="busway_export_button", callback=lambda: dpg.show_item("busway_file_dialog"), user_data="busway_file_dialog")
        
        # 湯淺
        with dpg.collapsing_header(label="湯淺"):
            dpg.add_input_text(label="批號", tag="yuasa_lot_no", default_value="T")
            with dpg.tree_node(label="AY8000R"):
                dpg.add_input_int(label="AY8000R數量", tag="ay8000r_quantity")
                dpg.add_input_int(label="黏度 cPs", tag="ay8000r_viscosity")
                dpg.add_input_int(label="凝膠時間 sec", tag="ay8000r_gel_time")
            with dpg.tree_node(label="AY8000B"):
                dpg.add_input_int(label="AY8000B數量", tag="ay8000b_quantity")
                dpg.add_input_int(label="黏度 cPs", tag="ay8000b_viscosity")
                dpg.add_input_int(label="凝膠時間 sec", tag="ay8000b_gel_time")
            with dpg.tree_node(label="HY8000"):        
                dpg.add_input_int(label="HY8000數量", tag="hy8000_quantity")
                dpg.add_input_float(label="黏度 cPs", tag="hy8000_viscosity")
            dpg.add_input_int(label="浸酸前引張強度 Kgf/cm2", tag="before_tensile_strength")
            dpg.add_input_int(label="浸酸後引張強度 Kgf/cm2", tag="after_tensile_strength")
            dpg.add_input_float(label="耐酸性 %", tag="acid_resistance")

            dpg.add_file_dialog(label="輸出湯淺報告", tag="yuasa_file_dialog", 
                directory_selector=True, show=False, default_path=current_export_path, 
                callback=export_yuasa_coa_report, height=500)
            dpg.add_button(label="輸出報告", tag="yuasa_export_button", callback=lambda: dpg.show_item("yuasa_file_dialog"), user_data="yuasa_file_dialog")

        dpg.bind_font(zh_header_font)
        dpg.bind_item_handler_registry(item="etacom_export_button", handler_registry="etacom_export_button_handler")
        dpg.bind_item_handler_registry(item="busway_export_button", handler_registry="busway_export_button_handler")
        dpg.bind_item_handler_registry(item="yuasa_export_button", handler_registry="yuasa_export_button_handler")

    ################################################################
    #                          Handlers                            #
    ################################################################
    # with dpg.item_handler_registry(tag="etacom_export_button_handler") as handler:
    #     dpg.add_item_clicked_handler(callback=export_type_1_coa_report, user_data={"company": "etacom_", "template": ETACOM_TEMPLATE_FILE})

    # with dpg.item_handler_registry(tag="busway_export_button_handler") as handler:
    #     dpg.add_item_clicked_handler(callback=export_type_1_coa_report, user_data={"company": "busway_", "template": BUSWAY_TEMPLATE_FILE})

    # with dpg.item_handler_registry(tag="yuasa_export_button_handler") as handler:
    #     dpg.add_item_clicked_handler(callback=export_yuasa_coa_report)

    dpg.create_viewport(title='瑞肯材料品檢報告產生器', width=900, height=600)
    dpg.setup_dearpygui()
    dpg.show_viewport()
    dpg.set_primary_window("Primary Window", True)
    dpg.start_dearpygui()
    dpg.destroy_context()

if __name__ == "__main__":
    run()
