"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import win32com.client as win32
import pythoncom


def process(orchestrator_connection: OrchestratorConnection) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")
    queue_element = orchestrator_connection.get_next_queue_element('ExcelRefresher','',True)
    values = queue_element.data.split('|')
    
    # Assign values to variables
    if len(values) >= 2:
        first_value = values[0]
        second_value = values[1]
        print(second_value)
        open_close_as_excel(first_value)
    else:
        raise ValueError("Queue element does not contain enough values")


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
