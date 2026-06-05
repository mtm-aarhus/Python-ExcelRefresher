from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import os
from robot_framework.process import process
from OpenOrchestrator.database.queues import QueueElement, QueueStatus
import json
from multiprocessing import freeze_support
from typing import Optional

def make_queue_element_with_payload(
    payload: dict | list,
    queue_name: str,
    reference: Optional[str] = None,
    created_by: Optional[str] = None,
    status: QueueStatus = QueueStatus.NEW, 
) -> QueueElement:
    # Validate & serialize
    data_str = json.dumps(payload, ensure_ascii=False)
    if len(data_str) > 2000:
        raise ValueError("data exceeds 2000 chars (column limit)")

    return QueueElement(
        queue_name=queue_name,
        status=status,
        data=data_str,
        reference=reference,
        created_by=created_by,
    )

def main():
    raw_json = """{
        "KopierKøelementFraOpenOrchestrator": "OgIndsætDetSomRaw_Json",
    }"""

    payload = json.loads(raw_json)

    qe = make_queue_element_with_payload(
        payload=payload,
        queue_name="ExcelRefresher",
        reference="Sandbox",
        status=QueueStatus.NEW, 
    )

    orchestrator_connection = OrchestratorConnection(
        "ExcelRefresher",
        os.getenv("OpenOrchestratorSQL"),
        os.getenv("OpenOrchestratorKey"),
        None,
        None,
        None
    )

    process(orchestrator_connection, qe)


if __name__ == "__main__":
    freeze_support()
    main()
