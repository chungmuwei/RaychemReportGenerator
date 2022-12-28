from docxtpl import DocxTemplate
import time

def generate_etacom_coa(product_name: str, lot_no: str, viscosity: int, gel_time: int):
    template = DocxTemplate(template_file="templates/COA_Etacom_template.docx")
    context = {
        "product_name": product_name,
        "date": time.strftime("%Y/%m/%d"),
        "lot_no": lot_no,
        "viscosity": viscosity,
        "gel_time": gel_time
    }
    print(context)
    template.render(context=context)
    
    # production path
    # template.save(filename=f"/Volumes/Business/steven_20200721/1 備份 20200410/工作/A_ISO續評/2022複評/表單/P003生產流程/瑞肯COA/COA_Etacom_2536_2537/COA_{product_name}_{time.strftime('%Y%m%d')}.docx")
    
    # testing path
    template.save(filename=f"output/COA_{product_name}_{time.strftime('%Y%m%d')}.docx")

# /Volumes/Business/steven_20200721/1\ 備份\ 20200410/工作/A_ISO續評/2022複評/表單/P003生產流程/瑞肯COA/COA_Etacom_2536_2537
