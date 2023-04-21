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
            compco_and_benchmarking_analysis(file_path)


def extract_macro_names(vba_code):
    """
    Extracts macro name from str which is used to execute code
    """
    macro_names = []
    for line in vba_code.splitlines():
        if line.startswith("Sub "):
            macro_names.append(line[4:].strip("()"))
    return macro_names


def compco_and_benchmarking_analysis(file_path):
    """
    Adds new sheet for analysis and injects vba macro to preform analysis.

    """
    # open the workbook in excel and set references names
    wb = xw.Book(file_path)

    analysis_sheet_name = "COMPCO AND BENCHMARKING"
    financial_sheet = "Financial Data"
    trading_multiples = "Trading Multiples"
    operating_sheet = "Operating Statistics"

    # add analysis sheet
    try:
        wb.sheets.add(analysis_sheet_name, before=financial_sheet)
    except ValueError:
        pass

    # open VBA editor
    vb = wb.app.api.VBE
    # add new module
    analysis_module = vb.VBProjects(1).VBComponents.Add(1)
    analysis_module.CodeModule.AddFromString(vba_code)

    # loop through each sheet and unmerge cells
    for sheet in wb.sheets:
        sheet.api.Cells.UnMerge()

    # select sheets as objects to navigate
    analysis_sheet = wb.sheets[analysis_sheet_name]
    fd_sheet = wb.sheets[financial_sheet]
    tm_sheet = wb.sheets[trading_multiples]
    op_sheet = wb.sheets[operating_sheet]

    # company names and stats set up
    names_and_metrics = op_sheet.range("A14").expand('table')
    dest_cells = analysis_sheet.range("C6")
    empty_rows = names_and_metrics.offset(
        names_and_metrics.rows.count, 0).resize(2)
    empty_rows.delete()
    names_and_metrics = op_sheet.range("A14").expand('table')
    names_and_metrics.copy(destination=dest_cells)

    # Set up Percentile table
    analysis_sheet.range("C28").value = "Company vs Peers"
    analysis_sheet.range("C30").value = "Percentiles"
    percentile_range = analysis_sheet.range("C31:C51")
    percentile_range.value = [[i/100] for i in range(0, 101, 5)]
    percentile_range.number_format = "0%"

    # Create Percentiles
    data_range = analysis_sheet.range("C6").expand("table").offset(1, 1)
    percentile_equations = analysis_sheet.range("D31:O51")

    for i, col in enumerate(percentile_equations.columns):
        data_col_letter = chr(ord('C') + i + 1)
        for j, cell in enumerate(col):
            row_num = cell.row
            percentile = f"C{row_num}"
            data_row_num = j + 7
            formula_str = f"=PERCENTILE({data_col_letter}7:{data_col_letter}16,{percentile})"
            try:
                cell.formula = formula_str
                cell.number_format = "0.00"
            except Exception as e:
                print(
                    f"Error: Failed to set formula '{formula_str} in column {col.column}, row {cell.row}")
                print(f"Exception: {e}")
    format_percentile = analysis_sheet.range("D31").expand('table')
    format_percentile.number_format = "0%"

    # Paste Index Match Formula
    row = analysis_sheet.range("C6").expand('table').last_cell.row
    comp_range = analysis_sheet.range("D29:O29")
    for i, col in enumerate(comp_range):
        try:
            column = chr(ord("C")+i+1)
            if column in ["M", "O"]:
                formula_str = f"=1-INDEX(C31:C51,MATCH({column}{row},{column}31:{column}51))"
            else:
                formula_str = f"=INDEX(C31:C51,MATCH({column}{row},{column}31:{column}51))"
            analysis_sheet.range(f"{column}29").formula = formula_str
            analysis_sheet.range(f"{column}29").number_format = "0%"
        except Exception as e:
            print(
                f"Error: Failed to set formula '{formula_str} in column {col}, row {row}")
            print(f"Exception: {e}")

    # Find average comp % of specified company -- used in comparative company analysis
    avg_str = f"=AVERAGE({'D29:O29'})"
    print(avg_str)
    analysis_sheet.range("D28").formula = avg_str
    analysis_sheet.range("D28").number_format = "0%"

    name = extract_macro_names(vba_code)[0]
    print(name)
    run_macro(wb, name)

    # close and save workbook
    wb.save()
    wb.close()


def run_all_macros(wb, vba_code):
    # Execute all macros in module
    module_name = "Module1."
    macro_names = extract_macro_names(vba_code)
    print(macro_names)
    for macro in macro_names:
        macro = wb.macro(module_name + macro)
        macro()


def run_macro(wb, name):
    module_name = "Module1."
    macro_names = name
    macro = wb.macro(module_name + macro)
    print(macro)
    macro()


def main():
    # user = input("Enter your Babson Username -> ")
    # folder_path = f"C:/Users/{user}/Documents/GitHub/Team-Project/excel_files"
    # mass_convert_xls_to_xlsm(folder_path)

    file_path = "C:/Users/savilabermudez1/Documents/GitHub/Team-Project/excel_files/Company Comparable Analysis Microsoft Corporation (1).xlsm"
    compco_and_benchmarking_analysis(file_path)
    print(f"-"*50, "\n COMPCO and Benchmarking was Successfully Executed")


if __name__ == '__main__':
    main()
