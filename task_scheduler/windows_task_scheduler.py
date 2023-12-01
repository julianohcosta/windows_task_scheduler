from datetime import datetime
import subprocess
from dataclasses import dataclass
from subprocess import CompletedProcess
from typing import List, Dict, Optional

# queries
QUERY_BY_TASK_FOLDER_OR_TASK_NAME:str = "query_by_taskFolder_or_taskName"
DISABLE_TASK:str = "disable_task"
ENABLE_TASK:str = "enable_task"

# Command dictionary
CMD:Dict[str, str] = {
    QUERY_BY_TASK_FOLDER_OR_TASK_NAME: "schtasks, /query, /fo, CSV, /nh, /tn, {}{}", # taskfolder ; task name
    DISABLE_TASK: "schtasks, /change, /tn, {}{}, /disable", # taskfolder ; task name"
    ENABLE_TASK: "schtasks, /change, /tn, {}{}, /enable" # taskfolder ; task name
}


@dataclass
class Task():
    """Task object"""
    task_name: str
    task_folder: str
    next_run_time: Optional[datetime]
    status: str


def task_exists(task:Task=None, task_name:str=None, task_folder:str=None) -> bool:
    """Check if a task exists"""

    if task is not None:
        task_name = task.task_name
        task_folder = task.task_folder
    elif task_name is None or task_folder is None:
        raise ValueError("Task name and task folder must be provided")

    result:CompletedProcess = exec_cmd(CMD[QUERY_BY_TASK_FOLDER_OR_TASK_NAME], task_name, task_folder)
    return result.returncode == 0


def query_folder(task_folder:str, datetime_format:str='%d/%m/%Y %H:%M:%S') -> List[Task]:
    """Query all tasks in a given folder"""

    result = exec_cmd(CMD[QUERY_BY_TASK_FOLDER_OR_TASK_NAME], "", task_folder)
    if task_folder in result.stdout:
        return parse_tasks(result.stdout, datetime_format)
    return []


def parse_tasks(text: str, datetime_format: str) -> List[Task]:
    ''' Parse string to Task object '''

    lines = text.strip().split('\n')
    tasks = []
    for line in lines:
        parts = [part.strip('"') for part in line.split(',')]
        task_folder, task_name = parts[0].rsplit('\\', 1)
        task_folder += '\\'
        next_run_time = datetime.strptime(parts[1], datetime_format) if parts[1] != "N/A" else None
        tasks.append(Task(task_name, task_folder, next_run_time, parts[2]))
    return tasks


def exec_cmd(query:str, task_name:str, task_folder:str) -> CompletedProcess:

    if task_folder is not None:
        if not task_folder.endswith("\\"):
            task_folder += "\\"
        cmd:List[str] = query.format(task_folder, task_name).split(", ")
    else:
        # Task is in the root folder
        cmd:List[str] = query.format(task_name, "").split(", ")

    return subprocess.run(cmd, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

# test
print(query_folder("\\Microsoft\\Office\\"))