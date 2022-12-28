import dearpygui.dearpygui as dpg
import generator

ETACON_TEMPLATE_FILE = "templates/COA_Etacom_template.docx"
ETACON_PRODUCT_NAME = ["樹脂CY2536", "硬化劑HY2536"]


def export_coa_report(sender, app_data, user_data):
    """
    Export Certificate of Analysis report word document
    """

    # print(f"sender is: {sender}")
    # print(f"app_data is: {app_data}")
    # print(f"user_data is: {user_data}")

    # get value from user
    product_name = dpg.get_value("product_name")
    lot_no = dpg.get_value("lot_no")
    viscosity = dpg.get_value("viscosity")
    gel_time = dpg.get_value("gel_time")

    generator.generate_coa_report(
        template_file=user_data, 
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
        dpg.add_item_clicked_handler(callback=export_etacon_coa)

    with dpg.window(label="Example Window", tag="Primary Window"):
        dpg.add_text("安達康")
        dpg.add_listbox(label="品名", tag="product_name", default_value="樹脂CY2536", items=ETACON_PRODUCT_NAME, num_items=2)
        dpg.add_input_text(label="批號", tag="lot_no", default_value="T")
        dpg.add_input_int(label="黏度 cPs", tag="viscosity")
        dpg.add_input_int(label="凝膠時間 sec", tag="gel_time")
        dpg.add_button(label="Export", tag="etacon_export_button")

        dpg.bind_font(zh_font)

        dpg.bind_item_handler_registry(item="etacon_export_button", handler_registry="etacon_export_button_handler")

    dpg.create_viewport(title='COA Generator', width=900, height=600)
    dpg.setup_dearpygui()
    dpg.show_viewport()
    dpg.set_primary_window("Primary Window", True)
    dpg.start_dearpygui()
    dpg.destroy_context()

if __name__ == "__main__":
    run()
