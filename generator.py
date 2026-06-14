"""
This file includes functions to generate COA report docx file
"""

import os
import re
import sys


from docxtpl import DocxTemplate

TEST_EXPORT_PATH = "/Users/raymond/Desktop/code/python/RaychemReportGenerator/output"
ETACON_PATH = "/Volumes/Business/steven_20200721/1 備份 20200410/工作/A_ISO續評/2022複評/表單/P003生產流程/瑞肯COA/COA_Etacom_2536_2537" 
DOCX_FILE_EXTENSION = ".docx"

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def generate_coa_report(template_file: str, context: dict[str, str], output_path: str | None = None):
    """
    Generate COA report based on corresponding template docx file and context (test results), 
    then save the report to provided output path or TEST_EXPORT_PATH if output path is not provided.
    """
    print(f"""Generate COA report with template file: {template_file} and context: {context}""")
    
    template = DocxTemplate(template_file=resource_path(template_file))
    template.render(context=context)
    
    product_name = context["product_name"]
    lot_no = context["lot_no"]

    # production path
    # filename = ETACON_PATH 

    # testing path
    # Use provided output path or fallback to TEST_EXPORT_PATH
    target_directory = output_path if output_path else TEST_EXPORT_PATH
    filepath = os.path.join(target_directory, "COA_" + re.sub(r'[^a-zA-Z0-9]', '', product_name) + "_" + lot_no)
    filepath = sequence_filename(filepath)
    print(f"Export docx at {filepath}")
    template.save(filename=filepath)
    return filepath

def sequence_filename(path: str) -> str:
    """Sequence filename to avoid filename conflicts if same type of report was generated more than once on the same day."""

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
    

    
