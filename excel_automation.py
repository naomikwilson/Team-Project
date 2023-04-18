import os
import win32com.client as client
from openpyxl import load_workbook, Workbook
from openpyxl.utils import column_index_from_string


def convert_xls_to_xlsm(folder_path):
    excel = client.Dispatch("excel.application")
    for file_name in os.listdir(folder_path):
        xls_file_path = os.path.join(folder_path, file_name)
        print(xls_file_path)
        wb = excel.Workbooks.Open(xls_file_path)
        wb.SaveAs(xls_file_path, 52)
        wb.Close
    excel.Quit()


def add_sheet_and_vba(filepath):
    # open the existing workbook
    wb = load_workbook(filepath)
    analysis_sheet = wb.create_sheet("Compco and Benchmarking Analysis")

    wb.save(filepath)


def main():
    user = "savilabermudez1"
    folder_path = f"C://Users/{user}//Documents//GitHub//Team-Project//excel_files//"
    file_name = "Company Comparable Analysis Apple Inc .xlsm"
    sample_file_path = os.path.join(folder_path, file_name)
    print(sample_file_path)

    add_sheet_and_vba(sample_file_path)

    text_code = '''
    Sub HelloWorld()
        MsgBox "Hello, World
    End Sub    

    '''


if __name__ == '__main__':
    main()
