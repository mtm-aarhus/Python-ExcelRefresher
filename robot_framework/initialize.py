"""This module defines any initial processes to run when the robot starts."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection


def initialize(orchestrator_connection: OrchestratorConnection) -> None:
    """Do all custom startup initializations of the robot."""
    orchestrator_connection.log_trace("Initializing.")
    orchestrator_connection.create_queue_element('ExcelRefresher',None,r"C:\Users\az60026\Downloads\Aktiviteter.xlsx|Test")
