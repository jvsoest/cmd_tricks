import requests
import os
from pathlib import Path
import json
import sys
from datetime import date

# Load ~/.cmd_tricks.json config file
config = None
with open(os.path.join(Path.home(), ".cmd_tricks.json")) as f:
    config = json.load(f)

# Exit if no config is found
if config is None:
    print("Could not load config file (at ~/.cmd_tricks.json)")
    sys.exit(1)

class TodoistHelper:
    def __init__(self, apiKey):
        self.__apiKey = apiKey
    
    def getDeadlineTasks(self):
        responseTasks = requests.get(
            "https://api.todoist.com/rest/v1/tasks",
            params={
                "filter": "@deadline"
            },
            headers={
                "Authorization": "Bearer %s" % self.__apiKey
            }).json()
        return responseTasks

tdHelper = TodoistHelper(config["todoist"]["api_key"])

from icalendar import Calendar, Event
cal = Calendar()
for task in tdHelper.getDeadlineTasks():
    event = Event()
    event.add('summary', task['content'])
    event.add('dtstart', date.fromisoformat(task['due']['date']))
    event.add('description', task['url'])
    event.add('location', task['url'])
    cal.add_component(event)

f = open(config["todoist"]["deadlines_storage_path"], 'wb')
f.write(cal.to_ical())
f.close()