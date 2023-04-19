import os
import xlrd
from openpyxl import load_workbook, Workbook
import aspose.cells
import xlwings as xw


def xls_to_xlsm(file_path):
    """"
    https://products.aspose.com/cells/python-net/conversion/xls-to-xlsm/
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
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xls"):
            file_path = folder_path + "/" + file_name
            xls_to_xlsm(file_path)


def inject_vba_code(file_path):
    """

    """
    # open the workbook in excel
    wb = xw.Book(file_path)

    # open the VBA editor
    wb.app.api.VBE.MainWindow.Visible = True

    # get a reference to the VBA Project
    vba_project = wb.app.api.VBE.ActiveVBProject

    # Create New Module and inject code
    module = vba_project.VBComponents.Add(1)
    module.CodeModule.AddFromString("My Function() \n ")

    # save changes and close the VBA editor
    wb.app.api.VBE.MainWindow.Visible = False
    wb.save()


def main():

    user = input("Enter your Babson Username -> ")
    folder_path = f"C:/Users/{user}/Documents/GitHub/Team-Project/excel_files"
    mass_convert_xls_to_xlsm(folder_path)

    file_path = "C:/Users/savilabermudez1/Documents/GitHub/Team-Project/excel_files/Company Comparable Analysis Apple Inc  (1).xlsm"
    print(file_path)
    inject_vba_code(file_path)
    print("Done")


    # add_sheet_all(folder_path)
    # text_code = '''
    # Sub HelloWorld()
    #     MsgBox "Hello, World
    # End Sub
if __name__ == '__main__':
    main()
