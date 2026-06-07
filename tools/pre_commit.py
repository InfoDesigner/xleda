import xlwings as xw
from pathlib import Path
import pywintypes
import time

xlsm_file = Path(__file__).parent.parent / "src/xleda/xleda_template.xlsm"
xlsx_file = Path(__file__).parent.parent / "src/xleda/xleda_template.xlsx"



def convert_template_to_xlsx():
        
    
    with xw.App(visible=False, add_book=False) as app:
        
        app.api.DisplayAlerts = False

        wb = app.books.open(xlsm_file, read_only=True)
        ws = wb.sheets('Field Analysis')
        ws.activate()

        # Loop through all shapes and clear their triggers
        for shp in ws.shapes:
            shp.api.OnAction = ""

        # Save as .xlsx and reenable alerts
        try:
            wb.api.SaveAs(str(xlsx_file), FileFormat=51)
            time.sleep(3)
            print("xlsx file updated")

        except pywintypes.com_error:
            print("xlsx file is open, and can't be updated")

        

        app.api.DisplayAlerts = True


if __name__ == '__main__':
    
    convert_template_to_xlsx()