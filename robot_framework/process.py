"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import win32com.client as win32
import pythoncom


def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")
    print(queue_element.__dict__)
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
        xlsx = win32.DispatchEx('Excel.Application')
        xlsx.DisplayAlerts = True
        xlsx.Visible = True
        book = xlsx.Workbooks.Open(file_path)
        book.RefreshAll()
        xlsx.CalculateUntilAsyncQueriesDone()
        book.Save()
        book.Close(SaveChanges=True)
        xlsx.Quit()
        pythoncom.CoUninitialize()
        book = None
        xlsx = None
        del book
        del xlsx
        print("-- Opened/Closed as Excel --")

    except Exception as e:
        print(e)

    finally:
        # RELEASES RESOURCES
        book = None
        xlsx = None
