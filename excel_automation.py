import os
import win32com.client as client


def convert_xls_to_xlsm(folder_path):
    excel = client.Dispatch("excel.application")
    for file_name in os.listdir(folder_path):
        xls_file_path = os.path.join(folder_path, file_name)
        print(xls_file_path)
        wb = excel.Workbooks.Open(xls_file_path)
        wb.SaveAs(xls_file_path, 52)
        wb.Close
    excel.Quit()


def main():
    user = "savilabermudez1"
    folder_path = f"C:/Users/{user}/Documents/GitHub/Team-Project/excel_files/"
    print(folder_path)
    convert_xls_to_xlsm(folder_path)


if __name__ == '__main__':
    main()
