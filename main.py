import os.path
import xlwings as xw
import os
import shutil
from pathlib import Path


# Path to find files
def input_files():
    source = input('Please, select a path to find .xlsx files: ')
    return source


# Add all .xls files to a master workbook
def consolidation(source):
    try:
        excel_files = list(Path(source).glob('*.xlsx'))
        combined_wb = xw.Book()

        for excel_file in excel_files:
            wb = xw.Book(excel_file)
            for sheet in wb.sheets:
                sheet.api.Copy(After=combined_wb.sheets[0].api)
            wb.close()

        combined_wb.sheets[0].delete()
        combined_wb.save(f'master_workbook.xlsx')
        if len(combined_wb.app.books) == 1:
            combined_wb.app.quit()
        else:
            combined_wb.close()
    except:
        print("An exception occurred trying to find files in the path selected.")


# Moving files to different folder
def move_files(source):
    try:
        all_files = list(Path(source).glob('*'))

        if Path("Processed").is_dir():
            print("")
        else:
            os.mkdir("Processed")
            os.mkdir("Not applicable")

        for all_file in all_files:
            if str(all_file).endswith(".xlsx"):
                shutil.move(all_file, "Processed")
            else:
                shutil.move(all_file, "Not applicable")

    except:
        print("An exception occurred trying to reassign directory to the files previous selected.")


def main():
    cycle_main = True

    while cycle_main:
        source_files = input_files()
        consolidation(source_files)
        move_files(source_files)

        cycle = input('Would you like to select another file? (y/n): ')

        if cycle.lower() == "n":
            cycle_main = False


main()
