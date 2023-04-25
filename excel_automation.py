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
            benchmarking_and_compco(file_path)


def extract_macro_names(vba_code):
    """
    Extracts macro name from str which is used to execute code
    """
    macro_names = []
    for line in vba_code.splitlines():
        if line.startswith("Sub "):
            macro_names.append(line[4:].strip("()"))
    return macro_names


def benchmarking_and_compco(file_path):
    """
    Adds new sheet for analysis and injects vba macro to preform analysis.
    https://www.geeksforgeeks.org/working-with-excel-files-in-python-using-xlwings/
    https://docs.xlwings.org/en/stable/quickstart.html#macros-call-python-from-excel
    http://www.et.byu.edu/~treedoug/_pages/teaching/ChEn263/Lectures/Lec23-XLWings_handout.pdf
    https://www.dataquest.io/blog/python-excel-xlwings-tutorial/
    """
    # open the workbook in excel and set references names
    wb = xw.Book(file_path)

    benchmarking_sheet = "BENCHMARKING"
    compco_sheet = "COMPCO"
    financial_sheet = "Financial Data"
    trading_multiples = "Trading Multiples"
    operating_sheet = "Operating Statistics"

    # add analysis sheet
    try:
        wb.sheets.add(benchmarking_sheet, before=financial_sheet)
        wb.sheets.add(compco_sheet, before=benchmarking_sheet)

    except ValueError:
        pass

    # open VBA editor
    vb = wb.app.api.VBE
    # add new module & add vba_code str
    module_name = "AnalysisModule"
    for comp in vb.VBProjects(1).VBComponents:
        if comp.Name == module_name:
            # If it exists, remove it
            vb.VBProjects(1).VBComponents.Remove(comp)

    # Add a new module with the specified code
    analysis_module = vb.VBProjects(1).VBComponents.Add(1)
    analysis_module.Name = module_name
    analysis_module.CodeModule.AddFromString(vba_code)

    # loop through each sheet and unmerge cells
    for sheet in wb.sheets:
        sheet.api.Cells.UnMerge()

    # select sheets as objects to navigate
    ben_sh = wb.sheets[benchmarking_sheet]
    fd_sh = wb.sheets[financial_sheet]
    tm_sh = wb.sheets[trading_multiples]
    op_sh = wb.sheets[operating_sheet]
    comp_sh = wb.sheets[compco_sheet]

    # company names and stats set up
    names_and_metrics = op_sh.range("A14").expand('table')
    dest_cells = ben_sh.range("C6")
    empty_rows = names_and_metrics.offset(
        names_and_metrics.rows.count, 0).resize(2)
    empty_rows.delete()
    names_and_metrics = op_sh.range("A14").expand('table')
    names_and_metrics.copy(destination=dest_cells)

    # Set up Percentile table
    ben_sh.range("C28").value = "Percentile Average"
    ben_sh.range("C29").value = "Company vs Peers"
    ben_sh.range("C30").value = "Percentiles"
    percentile_range = ben_sh.range("C31:C51")
    percentile_range.value = [[i/100] for i in range(0, 101, 5)]
    percentile_range.number_format = "0%"

    # Create Percentiles
    data_range = ben_sh.range("C6").expand("table").offset(1, 1)
    percentile_equations = ben_sh.range("D31:O51")

    for i, col in enumerate(percentile_equations.columns):
        data_col_letter = chr(ord('C') + i + 1)
        for j, cell in enumerate(col):
            row_num = cell.row
            percentile = f"C{row_num}"
            data_row_num = j + 7
            # MAKE DYNAMIC
            formula_str = f"=PERCENTILE({data_col_letter}7:{data_col_letter}16,{percentile})"
            try:
                cell.formula = formula_str
                cell.number_format = "0.00"
            except Exception as e:
                print(
                    f"Error: Failed to set formula '{formula_str} in column {col.column}, row {cell.row}")
                print(f"Exception: {e}")
    format_percentile = ben_sh.range("D31").expand('table')
    format_percentile.number_format = "0%"

    # Paste Index Match Formula
    row = ben_sh.range("C6").expand('table').last_cell.row
    comp_range = ben_sh.range("D29:O29")
    for i, col in enumerate(comp_range):
        try:
            column = chr(ord("C")+i+1)
            if column in ["M", "O"]:
                formula_str = f"=IFERROR(1-INDEX(C31:C51,MATCH({column}{row},{column}31:{column}51)),"")"
            else:
                formula_str = f"=IFERROR(INDEX(C31:C51,MATCH({column}{row},{column}31:{column}51)),"")"
            ben_sh.range(f"{column}29").formula = formula_str
            ben_sh.range(f"{column}29").number_format = "0%"
        except Exception as e:
            print(
                f"Error: Failed to set formula '{formula_str} in column {col}, row {row}")
            print(f"Exception: {e}")

    # Find average comp % of specified company -- used in comparative company analysis
    avg_str = f"=AVERAGE({'D29:O29'})"
    ben_sh.range("D28").formula = avg_str
    ben_sh.range("D28").number_format = "0%"

    # Run formatting macro
    run_macro(wb, vba_code, 0)
    '''
    '''
    # Copy Names Over ()
    names_range = ben_sh.range("C6").expand('down')
    names_range.copy(destination=comp_sh.range("C6"))


    #Identify all desired header pages
    shrout = fd_sh.range("C14").value
    mcap = fd_sh.range("D14").value
    net_debt = fd_sh.range("E14").value
    total_ev = fd_sh.range("H14").value
    ntm_rev = fd_sh.range("O14").value
    ntm_ebitda = fd_sh.range("P14").value
    ntm_eps = fd_sh.range("Q14").value

    compco_headers = [shrout, mcap, net_debt, total_ev, ntm_rev, ntm_ebitda, ntm_eps]

    for i, value in enumerate(compco_headers):
        cell = comp_sh.range("D6").offset(0, i)
        cell.value = value


    stat_range = comp_sh.range("D7:J17")
    for i, col in enumerate(stat_range):
        try:
            column = chr(ord("D")+i)
            formula_str = f"=IFERROR(INDEX(C31:C51,MATCH({column}{row},{column}31:{column}51)),"")"
            ben_sh.range(f"{column}29").formula = formula_str
            ben_sh.range(f"{column}29").number_format = "0%"
        except Exception as e:
            print(
                f"Error: Failed to set formula '{formula_str} in column {col}, row {row}")
            print(f"Exception: {e}")

    # close and save workbook
    # wb.save()
    # wb.close()


def test(file_path):
    wb = xw.Book(file_path)
    benchmarking_sheet = "BENCHMARKING"
    compco_sheet = "COMPCO"
    financial_sheet = "Financial Data"
    trading_multiples = "Trading Multiples"
    operating_sheet = "Operating Statistics"

    ben_sh = wb.sheets[benchmarking_sheet]
    fd_sh = wb.sheets[financial_sheet]
    tm_sh = wb.sheets[trading_multiples]
    op_sh = wb.sheets[operating_sheet]
    comp_sh = wb.sheets[compco_sheet]

    shrout = fd_sh.range("C14").value
    mcap = fd_sh.range("D14").value
    net_debt = fd_sh.range("E14").value
    total_ev = fd_sh.range("H14").value
    ntm_rev = fd_sh.range("O14").value
    ntm_ebitda = fd_sh.range("P14").value
    ntm_eps = fd_sh.range("Q14").value

    fd_stats = [shrout, mcap, net_debt, total_ev, ntm_rev, ntm_ebitda, ntm_eps]

    for i, value in enumerate(fd_stats):
        print(i, value)


def run_all_macros(wb, vba_code):
    # Execute all macros in module
    module_name = "Module1."
    macro_names = extract_macro_names(vba_code)
    print(macro_names)
    for macro in macro_names:
        macro = wb.macro(module_name + macro)
        macro()


def run_macro(wb, vba_code, location):
    module_name = "AnalysisModule."
    macro_names = extract_macro_names(vba_code)[location]
    macro_names = wb.macro(module_name + macro_names)
    print(macro_names)
    macro_names()


def main():
    # user = input("Enter your Babson Username -> ")
    user = "savilabermudez1"
    folder_path = f"C:/Users/{user}/Documents/GitHub/Team-Project/excel_files"
    mass_convert_xls_to_xlsm(folder_path)
    file_path = "C:/Users/savilabermudez1/Documents/GitHub/Team-Project/excel_files/Company Comparable Analysis Apple Inc  (2).xlsm"
    # mass_analysis(folder_path)
    # print(f"-"*50, "\n Benchmarking was Successfully Executed \n", "-"*50)
    # """Individual Testing Below """

    benchmarking_and_compco(file_path)
    # test(file_path)
    print(f"-"*50, "\n Benchmarking was Successfully Executed \n", "-"*50)


if __name__ == '__main__':
    main()
