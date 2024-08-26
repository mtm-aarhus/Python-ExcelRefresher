import win32com.client as win32
import pythoncom

def open_close_as_excel(file_path):
    try:
        pythoncom.CoInitialize()
        Xlsx = win32.DispatchEx('Excel.Application')
        Xlsx.DisplayAlerts = True
        Xlsx.Visible = True
        book = Xlsx.Workbooks.Open(file_path)
        book.RefreshAll()
        Xlsx.CalculateUntilAsyncQueriesDone()
        book.Save()
        book.Close(SaveChanges=True)
        Xlsx.Quit()
        pythoncom.CoUninitialize()

        book = None
        Xlsx = None
        del book
        del Xlsx
        print("-- Opened/Closed as Excel --")

    except Exception as e:
        print(e)

    finally:
        # RELEASES RESOURCES
        book = None
        Xlsx = None

open_close_as_excel(r"C:\Users\az60026\Downloads\Aktiviteter.xlsx")