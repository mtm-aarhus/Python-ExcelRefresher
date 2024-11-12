# Excel Data Refresher Process

This project automates the process of refreshing data in Excel files and uploading them to SharePoint using OpenOrchestrator and Robot Framework. The automation is set up to run periodically, leveraging queues and structured error handling.

## Dependencies

- `OpenOrchestrator == 1.*`
- `Pillow == 10.*`
- `Office365-REST-Python-Client == 2.5.13`
- `pywin32 == 306`


## Configuration

1. **Queue Configuration:**
   - The automation fetches queue elements from a SQL Server database. The query retrieves entries where the `TimeStamp` is older than 24 hours or is null, ensuring that the data is refreshed periodically.


## How It Works

1. **Fetching Queue Elements:**
   - Establishes a connection to SQL Server using `pyodbc`.
   - Retrieves rows from `QueueExcelRefresher` where `TimeStamp` is outdated.
   - Adds the data to an orchestrator queue for further processing.

2. **Processing Each Queue Element:**
   - Downloads the specified Excel file from SharePoint.
   - Refreshes the data using `win32com.client`.
   - Saves and uploads the refreshed file back to SharePoint.
   - Optionally, creates monthly folders named in Danish if the custom function is specified.


## Code Walkthrough

### Fetch Queue Elements
This is added to the queue_framework.py file to fetch and dispatch the queue from the PyOrchestrator db in the OpenOrchestrator SQL server. To avoid the same queue elements being run again if it runs every queue element multiple times a day for testing new queue elements, it only fetches timestamp from today and then writes the current timestamp for all the rows it just added to the queue. It dispatches the queuelement as a json file

```python
sql_server = orchestrator_connection.get_constant("SqlServer")
conn_string = "DRIVER={SQL Server};" + f"SERVER={sql_server.value};DATABASE=PYORCHESTRATOR;Trusted_Connection=yes;"
conn = pyodbc.connect(conn_string)
ac
current_time = datetime.now(timezone.utc)
time_threshold = current_time - timedelta(hours=20)

query = """
SELECT SharePointSite, FolderPath, CustomFunction
FROM [PyOrchestrator].[dbo].[QueueExcelRefresher]
WHERE TimeStamp < ? OR TimeStamp IS NULL
"""
cursor = conn.cursor()
cursor.execute(query, time_threshold)
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
```


### Custom Functionality
Certain queue elements deviates a little from the main process, so instead of having individual processes for small deviations they are incorporated into the process.

- **Monthly Folder Creation:**
  - Creates folders named in Danish, such as "Januar" or "Februar".
  - Stores the files in these organized monthly folders.

---

## Contact

For any questions or issues, feel free to reach out to the project maintainers.


# Robot-Framework V2

This repo is meant to be used as a template for robots made for [OpenOrchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator).

## Quick start

1. To use this template simply use this repo as a template (see [Creating a repository from a template](https://docs.github.com/en/repositories/creating-and-managing-repositories/creating-a-repository-from-a-template)).
__Don't__ include all branches.

2. Go to `robot_framework/__main__.py` and choose between the linear framework or queue based framework.

3. Implement all functions in the files:
    * `robot_framework/initialize.py`
    * `robot_framework/reset.py`
    * `robot_framework/process.py`

4. Change `config.py` to your needs.

5. Fill out the dependencies in the `pyproject.toml` file with all packages needed by the robot.

6. Feel free to add more files as needed. Remember that any additional python files must
be located in the folder `robot_framework` or a subfolder of it.

When the robot is run from OpenOrchestrator the `main.py` file is run which results
in the following:
1. The working directory is changed to where `main.py` is located.
2. A virtual environment is automatically setup with the required packages.
3. The framework is called passing on all arguments needed by [OpenOrchestrator](https://github.com/itk-dev-rpa/OpenOrchestrator).

## Requirements
Minimum python version 3.10

## Flow

This framework contains two different flows: A linear and a queue based.
You should only ever use one at a time. You choose which one by going into `robot_framework/__main__.py`
and uncommenting the framework you want. They are both disabled by default and an error will be
raised to remind you if you don't choose.

### Linear Flow

The linear framework is used when a robot is just going from A to Z without fetching jobs from an
OpenOrchestrator queue.
The flow of the linear framework is sketched up in the following illustration:

![Linear Flow diagram](Robot-Framework.svg)

### Queue Flow

The queue framework is used when the robot is doing multiple bite-sized tasks defined in an
OpenOrchestrator queue.
The flow of the queue framework is sketched up in the following illustration:

![Queue Flow diagram](Robot-Queue-Framework.svg)

## Linting and Github Actions

This template is also setup with flake8 and pylint linting in Github Actions.
This workflow will trigger whenever you push your code to Github.
The workflow is defined under `.github/workflows/Linting.yml`.

