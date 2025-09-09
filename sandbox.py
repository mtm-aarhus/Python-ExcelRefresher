from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import os
import win32com.client
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import time

def sharepoint_client(tenant: str, client_id: str, thumbprint: str, cert_path: str, sharepoint_site_url: str, orchestrator_connection: OrchestratorConnection) -> ClientContext:
    """
    Creates and returns a SharePoint client context.
    """
    # Authenticate to SharePoint
    cert_credentials = {
        "tenant": tenant,
        "client_id": client_id,
        "thumbprint": thumbprint,
        "cert_path": cert_path
    }
    ctx = ClientContext(sharepoint_site_url).with_client_certificate(**cert_credentials)

    # Load and verify connection
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()

    orchestrator_connection.log_info(f"Authenticated successfully. Site Title: {web.properties['Title']}")
    return ctx

def download_file_from_sharepoint(client: ClientContext, sharepoint_file_url: str) -> str:
    """
    Downloads a file from SharePoint and returns the local file path.
    Handles both cases where subfolders exist or only the root folder is used.
    """
    # Extract the root folder, folder path, and file name
    path_parts = sharepoint_file_url.split('/')
    DOCUMENT_LIBRARY = path_parts[0]  # Root folder name (document library)
    FOLDER_PATH = '/'.join(path_parts[1:-1]) if len(path_parts) > 2 else ''  # Subfolders inside root, or empty if none
    file_name = path_parts[-1]  # File name

    # Construct the local folder path inside the Documents folder
    documents_folder = os.path.join(os.path.expanduser("~"), "Documents", FOLDER_PATH) if FOLDER_PATH else os.path.join(os.path.expanduser("~"), "Documents", DOCUMENT_LIBRARY)

    # Ensure the folder exists
    if not os.path.exists(documents_folder):
        os.makedirs(documents_folder)

    # Define the download path inside the folder
    download_path = os.path.join(os.getcwd(), file_name)

    # Download the file from SharePoint
    with open(download_path, "wb") as local_file:
        file = (
            client.web.get_file_by_server_relative_path(sharepoint_file_url)
            .download(local_file)
            .execute_query()
        )
    # Define the maximum wait time (60 seconds) and check interval (1 second)
    wait_time = 60  # 60 seconds
    elapsed_time = 0
    check_interval = 1  # Check every 1 second


    # While loop to check if the file exists at `file_path`
    while not os.path.exists(download_path) and elapsed_time < wait_time:
        time.sleep(check_interval)  # Wait 1 second
        elapsed_time += check_interval

    # After the loop, check if the file still doesn't exist and raise an error
    if not os.path.exists(download_path):
        raise FileNotFoundError(f"File not found at {download_path} after waiting for {wait_time} seconds.")

    print(f"[Ok] file has been downloaded into: {download_path}")
    return download_path


def refresh_excel_file(file_path: str):
    """
    Refreshes an Excel file at the specified file path.
    """

    # Open an Instance of Application
    xlapp = win32com.client.DispatchEx("Excel.Application")

    # Optional, e.g., if you want to debug
    xlapp.Visible = True

    # Open File
    Workbook = xlapp.Workbooks.Open(file_path)

    # Refresh all  
    Workbook.RefreshAll()

    # Wait until Refresh is complete
    xlapp.CalculateUntilAsyncQueriesDone()

    # Save File  
    Workbook.Save()
    Workbook.Close(SaveChanges=True)

    # Quit Instance of Application
    xlapp.Quit()

    # Delete Instance of Application
    del Workbook
    del xlapp

    print(f"[Ok] Excel file at {file_path} has been refreshed and saved.")

def upload_file_to_sharepoint(client: ClientContext, sharepoint_file_url: str, local_file_path: str):
    """
    Uploads the specified local file back to SharePoint at the given URL.
    Uses the folder path directly to upload files.
    """
    # Extract the root folder, folder path, and file name
    path_parts = sharepoint_file_url.split('/')
    DOCUMENT_LIBRARY = path_parts[0]  # Root folder name (document library)
    FOLDER_PATH = '/'.join(path_parts[1:-1]) if len(path_parts) > 2 else ''  # Subfolders inside root, or empty if none
    file_name = path_parts[-1]  # File name

    # Construct the server-relative folder path (starting with the document library)
    if FOLDER_PATH:
        folder_path = f"{DOCUMENT_LIBRARY}/{FOLDER_PATH}"
    else:
        folder_path = f"{DOCUMENT_LIBRARY}"

    # Get the folder where the file should be uploaded
    target_folder = client.web.get_folder_by_server_relative_url(folder_path)
    client.load(target_folder)
    client.execute_query()

    # Upload the file to the correct folder in SharePoint
    with open(local_file_path, "rb") as file_content:
        uploaded_file = target_folder.upload_file(file_name, file_content).execute_query()

    print(f"[Ok] file has been uploaded to: {sharepoint_file_url} on SharePoint")
    os.remove(local_file_path)

# Usage Example:

# Get credentials from Orchestrator
orchestrator_connection = OrchestratorConnection("ExcelRefresher", os.getenv('OpenOrchestratorSQL'), os.getenv('OpenOrchestratorKey'), None)
RobotCredentials = orchestrator_connection.get_credential("Robot365User")
username = RobotCredentials.username
password = RobotCredentials.password

# SharePoint site URL
SHAREPOINT_SITE_URL = "https://aarhuskommune.sharepoint.com/Teams/tea-teamsite12345678"

# SharePoint file URL (full path including root folder and subfolder)
sharepoint_file_url = "Delte dokumenter/filename.xlsx"

# 1. Create the SharePoint client


client = sharepoint_client(username, password, SHAREPOINT_SITE_URL)

# 2. Download the file from SharePoint
local_file_path = download_file_from_sharepoint(client, sharepoint_file_url)

# 3. Refresh the Excel file
refresh_excel_file(local_file_path)

# 4. Upload the refreshed file back to SharePoint
upload_file_to_sharepoint(client, sharepoint_file_url, local_file_path)
