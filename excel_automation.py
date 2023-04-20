import os
import pandas as pd
from openpyxl import load_workbook, Workbook
import aspose.cells
import xlwings as xw
from vba_modules import vba_code


def xls_to_xlsm(file_path):
    """"
    https://products.aspose.com/cells/python-net/conversion/xls-to-xlsm/

    Converts XLS to XLSM - XLS is old excel and not compatible with Openpyxl
    """
    workbook = aspose.cells.Workbook(file_path)
    # Split the file path into its directory and file name components
    directory, filename = os.path.split(file_path)

    # # Change the file extension from ".xls" to ".xlsm"
    new_filename = os.path.splitext(filename)[0] + ".xlsm"

    # # Combine the directory and new filename components into a new file path
    new_folder_path = os.path.join(directory, new_filename)
    new_folder_path = new_folder_path.replace("\\", "/")
    workbook.save(new_folder_path)
    os.remove(file_path)


def mass_convert_xls_to_xlsm(folder_path):
    """
    Loops through all XLS files and converts them to XLSM (macro enabled workbook)

    """

    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xls"):
            file_path = folder_path + "/" + file_name
            xls_to_xlsm(file_path)


def mass_analysis(folder_path):

    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsm"):
            file_path = folder_path + "/" + file_name
            inject_vba_code(file_path, vba_code)


def extract_macro_names(vba_code):
    """
    Extracts macro name from str which is used to execute code
    """
    macro_names = []
    for line in vba_code.splitlines():
        if line.startswith("Sub "):
            macro_names.append(line[4:].strip("()"))
    return macro_names


def inject_vba_code(file_path, vba_code):
    """
    Adds new sheet for analysis and injects vba macro to preform analysis.

    """
    # open the workbook in excel
    wb = xw.Book(file_path)
    analysis_sheet_name = "COMPCO AND BENCHMARKING"
    add_before = "Financial Data"
    # add analysis sheet
    try:
        wb.sheets.add(analysis_sheet_name, before=add_before)
    except ValueError:
        pass
    # open VBA editor
    vb = wb.app.api.VBE
    # add new module
    analysis_module = vb.VBProjects(1).VBComponents.Add(1)
    analysis_module.CodeModule.AddFromString(vba_code)
    # Execute all macros in module
    module_name = "Module1."
    macro_names = extract_macro_names(vba_code)
    print(macro_names)
    for macro in macro_names:
        macro = wb.macro(module_name + macro)
        macro()

    # close and save workbook
    # wb.save()
    # wb.close()


def main():
    # user = input("Enter your Babson Username -> ")
    # folder_path = f"C:/Users/{user}/Documents/GitHub/Team-Project/excel_files"
    # mass_convert_xls_to_xlsm(folder_path)
    file_path = "C:/Users/savilabermudez1/Documents/GitHub/Team-Project/excel_files/Company Comparable Analysis Apple Inc  (1).xlsm"
    print(file_path)
    inject_vba_code(file_path, vba_code)
    print(f"-"*50, "\nDone")


if __name__ == '__main__':
    main()
