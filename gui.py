import dearpygui.dearpygui as dpg
import generator

ETACON_TEMPLATE_FILE = "templates/COA_Etacom_template.docx"
BUSWAY_TEMPLATE_FILE = "templates/COA_Busway_template.docx"
ETACON_PRODUCT_NAME = ["樹脂CY2536", "硬化劑HY2536"]
BUSWAY_PRODUCT_NAME = ["CY2533L7"]

def export_coa_report(sender, app_data, user_data):
    """
    Export Certificate of Analysis report word document
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

    generator.generate_coa_report(
        template_file=template, 
        product_name=product_name, 
        lot_no=lot_no, 
        viscosity=viscosity, 
        gel_time=gel_time
    )


def run():
    dpg.create_context()

    # add chinese font
    with dpg.font_registry():
        with dpg.font("fonts/Noto_Sans_TC/NotoSansTC-Regular.otf", 30) as zh_font:
            dpg.add_font_range_hint(dpg.mvFontRangeHint_Default)
            dpg.add_font_range_hint(dpg.mvFontRangeHint_Chinese_Full)
    # dpg.show_font_manager()

    ################################################################
    #                          Handler                             #
    ################################################################
    with dpg.item_handler_registry(tag="etacon_export_button_handler") as handler:
        dpg.add_item_clicked_handler(callback=export_coa_report, user_data={"company": "etacon_", "template": ETACON_TEMPLATE_FILE})

    with dpg.item_handler_registry(tag="busway_export_button_handler") as handler:
        dpg.add_item_clicked_handler(callback=export_coa_report, user_data={"company": "busway_", "template": BUSWAY_TEMPLATE_FILE})

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
            dpg.add_input_float(label="浸酸前引張強度%", tag="before_tensile_strength")
            dpg.add_input_float(label="浸酸後引張強度%", tag="after_tensile_strength")
            dpg.add_input_fload(label="耐酸性%", tag="acid_resistance")
            dpg.add_button(label="輸出報告", tag="yuasa_export_button")

        dpg.bind_font(zh_font)
        dpg.bind_item_handler_registry(item="etacon_export_button", handler_registry="etacon_export_button_handler")
        dpg.bind_item_handler_registry(item="busway_export_button", handler_registry="busway_export_button_handler")

    dpg.create_viewport(title='瑞肯材料品檢報告產生器', width=900, height=600)
    dpg.setup_dearpygui()
    dpg.show_viewport()
    dpg.set_primary_window("Primary Window", True)
    dpg.start_dearpygui()
    dpg.destroy_context()

if __name__ == "__main__":
    run()
