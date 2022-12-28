from docxtpl import DocxTemplate
import time
import os
import re

ETACON_PATH = "/Volumes/Business/steven_20200721/1 備份 20200410/工作/A_ISO續評/2022複評/表單/P003生產流程/瑞肯COA/COA_Etacom_2536_2537" 
DOCX_FILE_EXTENSION = ".docx"

def generate_coa_report(template_file: str, product_name: str, lot_no: str, viscosity: int, gel_time: int):
    template = DocxTemplate(template_file=template_file)
    context = {
        "product_name": product_name,
        "date": time.strftime("%Y/%m/%d"),
        "lot_no": lot_no,
        "viscosity": viscosity,
        "gel_time": gel_time
    }

    template.render(context=context)
    
    # production path
    # filename = ETACON_PATH 

    # testing path
    filename = "output"
    filename += "/COA_" + re.sub(r'[^a-zA-Z0-9]', '', product_name) + "_" + time.strftime('%Y%m%d')
    filename = sequence_filename(filename)
    template.save(filename=filename)

def sequence_filename(path: str) -> str:
    
    order = 2

    # Same type of report was generated once
    if os.path.exists(path+DOCX_FILE_EXTENSION):
        # append suffix "-1" to the existed filename
        os.rename(path+DOCX_FILE_EXTENSION, path+"-1"+DOCX_FILE_EXTENSION)
        # append suffix "-1" to the current filename
        return path+"-2"+DOCX_FILE_EXTENSION
    
    # Same type of report was generated twice or more,
    # which means they has been sequenced
    elif os.path.exists(path+"-"+str(order)+DOCX_FILE_EXTENSION):
        while (os.path.exists(path+"-"+str(order)+DOCX_FILE_EXTENSION)):
            order += 1
        return path+"-"+str(order)+DOCX_FILE_EXTENSION
    
    # Same type of report was not generated today
    return path+DOCX_FILE_EXTENSION
    

    