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
            macro_names.append(line[4:].strip().rstrip('()'))
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

    # select sheets as objects to navigate
    ben_sh = wb.sheets[benchmarking_sheet]
    fd_sh = wb.sheets[financial_sheet]
    tm_sh = wb.sheets[trading_multiples]
    op_sh = wb.sheets[operating_sheet]
    comp_sh = wb.sheets[compco_sheet]

    sheet_names = [fd_sh, tm_sh, op_sh, comp_sh, ben_sh]

    # ADD vba Module
    add_vba_macros(wb)

    # loop through each sheet and unmerge cells
    for sheet in wb.sheets:
        sheet.api.Cells.UnMerge()

    # delete rows in main raw data sheets
    target_sheets = sheet_names[:3]
    delete_rows(target_sheets)

    # Paste newly formatted data
    dest_cells = ben_sh.range("C6")
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
    starting_data_cell = "D7"
    percentile_1_column = "C"
    percentile_equations = ben_sh.range("D31:O51")

    create_percentile_equation(
        percentile_equations, starting_data_cell, percentile_1_column)

    # Paste Index Match Formula
    row = ben_sh.range("C6").expand('table').last_cell.row
    comp_range = ben_sh.range("D29:O29")
    for i, col in enumerate(comp_range):
        try:
            column = chr(ord("C")+i+1)
            if column in ["M", "O"]:
                formula_str = f'=IFERROR((1-INDEX(C31:C51,MATCH({column}{row},{column}31:{column}51))),"NA")'

            else:
                formula_str = f'=IFERROR(INDEX(C31:C51,MATCH({column}{row},{column}31:{column}51)),"NA")'

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
    all_macros = extract_macro_names(vba_code)
    formatting_macro = all_macros[0]
    run_macro(wb, formatting_macro)

    # set up compco data
    compco_set_up(sheet_names)

    # Set up Percentile table

    comp_sh.range("K22").value = "Applied Percentile"
    comp_sh.range("L22").formula = "=BENCHMARKING!D28"
    comp_sh.range("L22").number_format = "0%"
    comp_sh.range("K23").value = "Percentiles"
    percentile_range = comp_sh.range("K24:K44")
    percentile_range.value = [[i/100] for i in range(0, 101, 5)]
    percentile_range.number_format = "0%"

    # run percentile analysis on trading multiples
    trading_percentile_range = comp_sh.range("L24:N44")
    starting_cell = "L7"
    percentile_2_column = "K"

    create_percentile_equation(
        trading_percentile_range, starting_cell, percentile_2_column)

    # set Relative Value Analysis
    relative_valuation(sheet_names)

    # Run formatting macro
    all_macros = extract_macro_names(vba_code)
    compco_format1 = all_macros[1]
    compco_format2 = all_macros[2]
    run_macro(wb, compco_format1)
    run_macro(wb, compco_format2)

    wb.save()
    wb.close()


def compco_set_up(sheet_names):
    ''''
    sheet_names = [fd_sh, tm_sh, op_sh, comp_sh, ben_sh]

    '''
    financial_data = sheet_names[0]
    benchmarking_sheet = sheet_names[-1]
    trading_multiples = sheet_names[1]
    compco_sheet = sheet_names[3]

    # Copy Names Over ()
    names_range = benchmarking_sheet.range("C6").expand('down')
    names_range.copy(destination=compco_sheet.range("C6"))

    # Identify all desired header pages

    current_price = financial_data.range("B14").expand('down')
    shrout = financial_data.range("C14").expand('down')
    mcap = financial_data.range("D14").expand('down')
    net_debt = financial_data.range("E14").expand('down')
    total_ev = financial_data.range("H14").expand('down')
    ntm_rev = financial_data.range("O14").expand('down')
    ntm_ebitda = financial_data.range("P14").expand('down')
    ntm_eps = financial_data.range("Q14").expand('down')

    ev_to_rev = trading_multiples.range("G14").expand('down')
    ev_to_ebitda = trading_multiples.range("H14").expand('down')
    price_to_earnings = trading_multiples.range("I14").expand('down')

    compco_headers = [current_price, shrout, mcap, net_debt,
                      total_ev, ntm_rev, ntm_ebitda, ntm_eps, ev_to_rev, ev_to_ebitda, price_to_earnings]
    starting_column = "D"
    for item in compco_headers:
        item.copy(destination=compco_sheet.range(f"{starting_column}6"))
        starting_column = chr(ord(starting_column)+1)


def relative_valuation(sheet_names):
    """sheet_names = [fd_sh, tm_sh, op_sh, comp_sh, ben_sh]
    Pull sheet list and m,ake varibles with relevant ones

    """

    compco_sheet = sheet_names[3]
    compco_sheet.range("P6").value = "NTM EV/REV"
    compco_sheet.range("Q6").value = "NTM EV/EBITDA"
    compco_sheet.range("R6").value = "NTM P/E"

    # Paste Variables
    valuation_labels = ["Weighting %", "Applied Multiple", "Company Metric", "EV", "Net Debt", "Equity Value",
                        "Shares Outstanding", "Implied Price / Share", "Weighted Price", "Current Price", "Implied Market Upside(Downside)"]
    starting_row = 7
    for label in valuation_labels:
        compco_sheet.range(f"O{starting_row}").value = label
        starting_row += 1

    # set company metrics to variable
    ntm_rev = compco_sheet.range("I6").end('down').value
    ntm_ebitda = compco_sheet.range("J6").end('down').value
    ntm_pe = compco_sheet.range("K6").end('down').value

    # set default weighting for blended price target
    default_weighting = 1/3
    compco_sheet.range("P7:R7").value = default_weighting
    compco_sheet.range("P7:R7").number_format = "0%"

    # set index match that uses avg from benchmarking analysis
    starting_column1 = "L"
    starting_column2 = "P"

    for i in range(3):
        # always sent to these cells
        formula_str = f"=INDEX({starting_column1}24:{starting_column1}44,MATCH($L$22,$K$24:$K$44))"
        compco_sheet.range(f"{starting_column2}8").formula = formula_str
        starting_column1 = chr(ord(starting_column1)+1)
        starting_column2 = chr(ord(starting_column2)+1)

    # Set Company Metric
    compco_sheet.range("P9").value = ntm_rev
    compco_sheet.range("Q9").value = ntm_ebitda
    compco_sheet.range("R9").value = ntm_pe

    # Set EV equations
    compco_sheet.range("P10").formula = "=P9*P8"
    compco_sheet.range("Q10").formula = "=Q9*Q8"
    compco_sheet.range("R10").formula = "=R12+R11"

    # Set Net Debt
    compco_sheet.range("P11:R11").value = compco_sheet.range(
        "G14").end('down').value

    # Set Equity Value
    compco_sheet.range("P12").formula = "=P10-P11"
    compco_sheet.range("Q12").formula = "=Q10-Q11"
    compco_sheet.range("R12").formula = "=R13*R14"

    # Set Shares Outstanding
    compco_sheet.range("P13:R13").value = compco_sheet.range(
        "E6").end('down').value

    # Set Implied Share Price
    compco_sheet.range("P14").formula = "=P12/P13"
    compco_sheet.range("Q14").formula = "=Q12/Q13"
    compco_sheet.range("R14").formula = "=R8*R9"

    # Set Weighted Price
    compco_sheet.range("P15").formula = "=SUMPRODUCT(P14:R14,P7:R7)"

    # Set Current Price
    compco_sheet.range("P16").value = compco_sheet.range(
        "D6").end('down').value

    # Set Implied over/under price
    compco_sheet.range("P17").formula = "=P15/P16-1"
    compco_sheet.range("P17").number_format = "0%"


def add_vba_macros(wb):
    '''
    Add vba module and copy paste vba_code str from py file 
    '''
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


def delete_rows(sheet_names):
    '''
    Delete rows in specified sheet names - needed to correct downloaded format into useable form
    '''
    for sheet_name in sheet_names:
        # company names and stats set up
        names_and_metrics = sheet_name.range("A14").expand('table')

        empty_rows = names_and_metrics.offset(
            names_and_metrics.rows.count, 0).resize(2)
        empty_rows.delete()


def create_percentile_equation(range, starting_data_cell, percentile_column):
    for i, col in enumerate(range.columns):
        data_col_letter = chr(ord(percentile_column) + i + 1)
        for j, cell in enumerate(col):
            row_num = cell.row
            percentile = f"{percentile_column}{row_num}"
            last_row = col.sheet.range(starting_data_cell).end('down').row - 1
            formula_str = f"=PERCENTILE({data_col_letter}7:{data_col_letter}{last_row},{percentile})"
            try:
                cell.formula = formula_str
                cell.number_format = "0.00"
            except Exception as e:
                print(
                    f"Error: Failed to set formula '{formula_str} in column {col.column}, row {cell.row}")
                print(f"Exception: {e}")


def run_macro(wb, macro_name):
    module_name = "AnalysisModule."
    macro = wb.macro(module_name + macro_name)
    print(macro)
    macro()


def main():
    user = input("Enter your Babson Username -> ")
    folder_path = f"C:/Users/{user}/Documents/GitHub/Team-Project/excel_files"
    mass_convert_xls_to_xlsm(folder_path)
    mass_analysis(folder_path)
    print(f"-"*50, "\n Benchmarking was Successfully Executed \n", "-"*50)


if __name__ == '__main__':
    main()
