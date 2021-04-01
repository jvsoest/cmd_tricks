from todoist.api import TodoistAPI
import clipboard
import argparse
import os
from pathlib import Path
import json
import sys
import tkinter as tk

# Parse input parameters
parser = argparse.ArgumentParser(description='Add task to Todoist Inbox.')
parser.add_argument('-cb', '--clipboard', help='Read note contents from clipboard', action='store_true')
parser.add_argument('-ui', '--userInterface', help='Show user interface to save information', action='store_true')
parser.add_argument('taskName', nargs='?', default='', help='the actual text for the task')
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

# get contents from clipboard
cbContents = None
if inputArgs.clipboard:
    cbContents = clipboard.paste()

############################
# Add GUI functions
############################
class UserInterface:

    def __init__(self, todoistApi, taskName, taskNote):
        # super().__init__()
        self.todoistApi = todoistApi

        root = tk.Tk()
        root.title("Todoist quick entry")

        self.txtTaskName = tk.Entry(root, width=80)
        self.txtTaskName.pack(side=tk.TOP, anchor=tk.NW, pady=10, padx=10)
        self.txtTaskName.insert(tk.INSERT, taskName)

        self.txtNotes = tk.Text(root, width=60, height=12)
        self.txtNotes.pack(side=tk.TOP, anchor=tk.NW, pady=10, padx=10)
        if taskNote != None:
            self.txtNotes.insert(tk.INSERT, taskNote)

        btnSave = tk.Button(root, text="Save", command = self.saveActionButton)
        btnSave.pack(side=tk.TOP)

        # keybind control + enter to save action
        root.bind('<Control-Return>', self.saveActionShortcut)

        # keybind escape to exit
        root.bind('<Escape>', self.quitAction)
        self.txtTaskName.focus_set()

        root.mainloop()

    def saveActionButton(self):
        taskName = self.txtTaskName.get()
        taskNote = self.txtNotes.get("1.0",'end-1c')
        self.todoistApi.addItemToTodoist(taskName, taskNote)

    def saveActionShortcut(self, argument):
        taskName = self.txtTaskName.get()
        taskNote = self.txtNotes.get("1.0",'end-1c')
        self.todoistApi.addItemToTodoist(taskName, taskNote)

    def quitAction(self, argument):
        sys.exit(0)
        

############################
# Add actual todoist action
############################
class TodoistActions:
    def __init__(self, api):
        super().__init__()
        self.api = api
    def addItemToTodoist(self, taskName, taskNote):
        savePerformed = False
        if taskNote is not None:
            if taskNote != "":
                self.api.quick.add(taskName, note=taskNote)
                savePerformed = True

        if not savePerformed:
            self.api.quick.add(taskName)
        
        sys.exit(0)

# determine to show UI or not...
todoistApi = TodoistActions(api)
if inputArgs.userInterface:
    UserInterface(todoistApi, inputArgs.taskName, cbContents)
else:
    todoistApi.addItemToTodoist(inputArgs.taskName, cbContents)