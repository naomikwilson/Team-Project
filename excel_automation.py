import os
import pandas as pd
from openpyxl import workbook, load_workbook


def convert_xls_to_xlsm(folder_path):

    # Loop through all files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".xls"):
            # Open the XLS file
            xls_file_path = os.path.join(folder_path, filename)
            df = pd.read_excel(xls_file_path)

            # Save the XLS file as XLSM
            xlsm_file_path = os.path.join(
                folder_path, os.path.splitext(filename)[0] + ".xlsm")
            writer = pd.ExcelWriter(xlsm_file_path, engine='openpyxl')
            df.to_excel(writer, index=False)
            writer.close()


def main():
    user = "savilabermudez1"
    folder_path = f"C:/Users/{user}/Documents/GitHub/Team-Project/excel_files"
    convert_xls_to_xlsm(folder_path)


if __name__ == '__main__':
    main()
