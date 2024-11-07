"""This module is the primary module of the robot framework. It collects the functionality of the rest of the framework."""

# This module is not meant to exist next to linear_framework.py in production:
# pylint: disable=duplicate-code

import sys

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueStatus

from robot_framework import initialize
from robot_framework import reset
from robot_framework.exceptions import handle_error, BusinessError, log_exception
from robot_framework import process
from robot_framework import config
from datetime import datetime, timedelta, timezone
import pyodbc
import json


def main():
    """The entry point for the framework. Should be called as the first thing when running the robot."""
    orchestrator_connection = OrchestratorConnection.create_connection_from_args()
    sys.excepthook = log_exception(orchestrator_connection)
    orchestrator_connection.log_trace("Robot Framework started.")
    initialize.initialize(orchestrator_connection)
    sql_server = orchestrator_connection.get_constant("SqlServer")

    # Gets queue from db
    conn = pyodbc.connect(
   "DRIVER={SQL Server};"+f"SERVER={sql_server};DATABASE=PYORCHESTRATOR;Trusted_Connection=yes;")

    # Get the current UTC time and 24-hour threshold
    current_time = datetime.now(timezone.utc)  # Timezone-aware UTC time
    time_threshold = current_time - timedelta(hours=24)

    # Step 1: Fetch rows where the timestamp is more than 24 hours old
    query = """
    SELECT SharePointSite, FolderPath, CustomFunction
    FROM [PyOrchestrator].[dbo].[QueueExcelRefresher]
    WHERE TimeStamp < ? OR TimeStamp IS NULL
    """
    cursor = conn.cursor()
    cursor.execute(query, time_threshold)

    # Retrieve data and prepare `references` and `data`
    rows = cursor.fetchall()
    if rows:
        references = tuple(row[1] for row in rows)  # Using FolderPath as the reference

        # Convert each row to a JSON string for structured data storage
        data = tuple(json.dumps({
            "SharePointSite": row[0],
            "FolderPath": row[1],
            "CustomFunction": row[2]
        }) for row in rows)

        # Call bulk_create_queue_elements with JSON-formatted data
        orchestrator_connection.bulk_create_queue_elements("ExcelRefresher", references=references, data=data)
        update_query = """
        UPDATE [PyOrchestrator].[dbo].[QueueExcelRefresher]
        SET TimeStamp = ? WHERE TimeStamp < ? OR TimeStamp IS NULL
        """
        cursor.execute(update_query, (current_time, time_threshold))
        conn.commit()

    queue_element = None
    error_count = 0
    task_count = 0
    # Retry loop
    for _ in range(config.MAX_RETRY_COUNT):
        try:
            reset.reset(orchestrator_connection)

            # Queue loop
            while task_count < config.MAX_TASK_COUNT:
                task_count += 1
                queue_element = orchestrator_connection.get_next_queue_element(config.QUEUE_NAME)

                if not queue_element:
                    orchestrator_connection.log_info("Queue empty.")
                    break  # Break queue loop

                try:
                    process.process(orchestrator_connection, queue_element)
                    orchestrator_connection.set_queue_element_status(queue_element.id, QueueStatus.DONE)

                except BusinessError as error:
                    handle_error("Business Error", error, queue_element, orchestrator_connection)

            break  # Break retry loop

        # We actually want to catch all exceptions possible here.
        # pylint: disable-next = broad-exception-caught
        except Exception as error:
            error_count += 1
            handle_error(f"Process Error #{error_count}", error, queue_element, orchestrator_connection)

    reset.clean_up(orchestrator_connection)
    reset.close_all(orchestrator_connection)
    reset.kill_all(orchestrator_connection)

    if config.FAIL_ROBOT_ON_TOO_MANY_ERRORS and error_count == config.MAX_RETRY_COUNT:
        raise RuntimeError("Process failed too many times.")

main()