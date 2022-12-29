import dearpygui.dearpygui as dpg
import generator
import time
from datetime import date
from dateutil.relativedelta import relativedelta    

ETACON_TEMPLATE_FILE = "templates/COA_Etacom_template.docx"
BUSWAY_TEMPLATE_FILE = "templates/COA_Busway_template.docx"
YUASA_TEMPLATE_FILE = "templates/COA_Yuasa_template.docx"

ETACON_PRODUCT_NAME = ["樹脂CY2536", "硬化劑HY2536"]
BUSWAY_PRODUCT_NAME = ["CY2533L7"]

def export_type_1_coa_report(sender, app_data, user_data):
    """
    Export Certificate of Analysis report word document for Etacon and Busway
    """

    # print(f"sender is: {sender}")
    # print(f"app_data is: {app_data}")
    # print(f"user_data is: {user_data}")

    # get value from user
    company = user_data["company"]
    template = user_data["template"]

    product_name = dpg.get_value(company+"product_name")
    lot_no = dpg.get_value(company+"lot_no")
    viscosity = dpg.get_value(company+"viscosity")
    gel_time = dpg.get_value(company+"gel_time")

    context = {
        "product_name": product_name,
        "date": time.strftime("%Y/%m/%d"),
        "lot_no": lot_no,
        "viscosity": viscosity,
        "gel_time": gel_time
    }

    generator.generate_coa_report(template_file=template, context=context)

def export_yuasa_coa_report():
    """
    Export Certificate of Analysis report word document for Yuasa
    """

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
    generator.generate_coa_report(template_file=YUASA_TEMPLATE_FILE, context=context)

def run():

    dpg.create_context()

    ################################################################
    #                           Fonts                              #
    ################################################################
    
    # add chinese font
    with dpg.font_registry():
        with dpg.font("fonts/Noto_Sans_TC/NotoSansTC-Regular.otf", 24) as zh_font:
            dpg.add_font_range_hint(dpg.mvFontRangeHint_Default)
            dpg.add_font_range_hint(dpg.mvFontRangeHint_Chinese_Full)
    # dpg.show_font_manager()

    # dpg.show_style_editor()

    ################################################################
    #                          Handlers                            #
    ################################################################
    with dpg.item_handler_registry(tag="etacon_export_button_handler") as handler:
        dpg.add_item_clicked_handler(callback=export_type_1_coa_report, user_data={"company": "etacon_", "template": ETACON_TEMPLATE_FILE})

    with dpg.item_handler_registry(tag="busway_export_button_handler") as handler:
        dpg.add_item_clicked_handler(callback=export_type_1_coa_report, user_data={"company": "busway_", "template": BUSWAY_TEMPLATE_FILE})

    with dpg.item_handler_registry(tag="yuasa_export_button_handler") as handler:
        dpg.add_item_clicked_handler(callback=export_yuasa_coa_report)

    ################################################################
    #                          Windows                             #
    ################################################################
    with dpg.window(label="Example Window", tag="Primary Window"):
        # 安達康
        with dpg.collapsing_header(label="安達康"):
            dpg.add_listbox(label="品名", tag="etacon_product_name", default_value="樹脂CY2536", items=ETACON_PRODUCT_NAME, num_items=2)
            dpg.add_input_text(label="批號", tag="etacon_lot_no", default_value="T")
            dpg.add_input_int(label="黏度 cPs", tag="etacon_viscosity")
            dpg.add_input_int(label="凝膠時間 sec", tag="etacon_gel_time")
            dpg.add_button(label="輸出報告", tag="etacon_export_button")
        
        # 巴斯威爾
        with dpg.collapsing_header(label="巴斯威爾"):
            dpg.add_listbox(label="品名", tag="busway_product_name", default_value="CY2533L7", items=BUSWAY_PRODUCT_NAME, num_items=2)
            dpg.add_input_text(label="批號", tag="busway_lot_no", default_value="T")
            dpg.add_input_int(label="黏度 cPs", tag="busway_viscosity")
            dpg.add_input_int(label="凝膠時間 sec", tag="busway_gel_time")
            dpg.add_button(label="輸出報告", tag="busway_export_button")
        
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
            dpg.add_button(label="輸出報告", tag="yuasa_export_button")

        dpg.bind_font(zh_font)
        dpg.bind_item_handler_registry(item="etacon_export_button", handler_registry="etacon_export_button_handler")
        dpg.bind_item_handler_registry(item="busway_export_button", handler_registry="busway_export_button_handler")
        dpg.bind_item_handler_registry(item="yuasa_export_button", handler_registry="yuasa_export_button_handler")

    dpg.create_viewport(title='瑞肯材料品檢報告產生器', width=900, height=600)
    dpg.setup_dearpygui()
    dpg.show_viewport()
    dpg.set_primary_window("Primary Window", True)
    dpg.start_dearpygui()
    dpg.destroy_context()

if __name__ == "__main__":
    run()
