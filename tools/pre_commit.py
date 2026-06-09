import xlwings as xw
from pathlib import Path
import time
import platform


os = platform.system()
win = os == 'Windows'
mac = os == 'Darwin'

xlsm_file = Path(__file__).parent.parent / "src/xleda/xleda_template.xlsm"
xlsx_file = Path(__file__).parent.parent / "src/xleda/xleda_template.xlsx"


shape_collection = {'debug': ['Config'],
                    'Field Analysis': ['Field Analysis', 'Data Description', 'Composition', 'Summary Stats', 'Percentiles', 'Field Lists', 'Compiled Lists']}



def convert_template_to_xlsx():
        
    
    with xw.App(visible=False, add_book=False) as app:
        
        app.display_alerts = False
        book = app.books.open(xlsm_file, read_only=True)

        # ---------------------------------------------------------------
        # Loop through sheets and clear out VBA related UI elements

        for sheet in shape_collection.keys():

            ws = book.sheets(sheet)
            ws.activate()
            
            # Loop through all shapes and delete them
            for shp in shape_collection[sheet]:
                ws.shapes(shp).delete()

            ws.range("A2").select()

        
        # ---------------------------------------------------------------
        # Save/catch errors

        try:
            if win:

                # Save as .xlsx and reenable alerts
                book.api.SaveAs(str(xlsx_file), FileFormat=51)

            
            if mac:

                # Convert path to a Mac HFS path (with colons)
                hfs_target_path = xw._xlmac.posix_to_hfs_path(xlsx_file)

                # Call save_workbook_as via .api WITHOUT providing a file_format flag
                book.api.save_workbook_as(filename=hfs_target_path, overwrite=True)

        except Exception as e:
            print(f"Error: {e}")


        # Add a pause to let the file save
        time.sleep(3)

        print("xlsx file updated")

        app.api.DisplayAlerts = True


if __name__ == '__main__':
    
    convert_template_to_xlsx()