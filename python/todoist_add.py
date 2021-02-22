from todoist.api import TodoistAPI
import clipboard
import argparse
import os
from pathlib import Path
import json
import sys

# Parse input parameters
parser = argparse.ArgumentParser(description='Add task to Todoist Inbox.')
parser.add_argument('-cb', '--clipboard', help='Read note contents from clipboard', action='store_true')
parser.add_argument('taskName', help='the actual text for the task')
inputArgs = parser.parse_args()

# Load ~/.cmd_tricks.json config file
config = None
with open(os.path.join(Path.home(), ".cmd_tricks.json")) as f:
    config = json.load(f)

# Exit if no config is found
if config is None:
    print("Could not load config file (at ~/.cmd_tricks.json)")
    sys.exit(1)

# Read Todoist API key form .cmd_tricks.json
api = TodoistAPI(config["todoist"]["api_key"])

taskName = inputArgs.taskName

# Create task, if clipboard cmd is given, use this
if inputArgs.clipboard:
    taskNote = clipboard.paste()
    api.quick.add(taskName, note=taskNote)
else:
    api.quick.add(taskName)
