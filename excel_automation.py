import os
import xlrd
from openpyxl import load_workbook, Workbook


def convert_xls_to_xlsm(folder_path):
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xls"):
            # open xls file using xlrd
            xls_book = xlrd.open_workbook(os.path.join(folder_path, file_name))

            # create new XLSM book using openpyxl
            xlsm_new_book = Workbook()

            for sheet_name in xls_book.sheet_name():
                sheet = xls_book.sheet_by_name(sheet_name)
                new_sheet = xlsm_new_book.create_sheet(title=sheet_name)
                for row_index in range(sheet.nrows()):
                    new_sheet.append(sheet.row_value(row_index))
        name, _ = os.path.splitext(file_name)
        xlsm_new_book.save(os.path.join(folder_path, name + "xlsm"))


def add_sheet_and_vba(filepath):
    pass


def main():
    user = "savilabermudez1"
    folder_path = f"C://Users/{user}//Documents//GitHub//Team-Project//excel_files//"
    file_name = "Company Comparable Analysis Apple Inc .xlsm"

    sample_file_path = os.path.join(folder_path, file_name)
    print(sample_file_path)

    convert_xls_to_xlsm(folder_path)

    # add_sheet_and_vba(sample_file_path)

    # text_code = '''
    # Sub HelloWorld()
    #     MsgBox "Hello, World
    # End Sub

    # '''


if __name__ == '__main__':
    main()
