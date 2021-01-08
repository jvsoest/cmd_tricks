#################################################################################
# How to use:
#  python keypress.py <arguments> [file or string to type]
#
#  Arguments:
#    -int   Optional; interval (in seconds) between keystrokes, decimals allowed
#    -s     Optional; do not search for file, but use input string as keystrokes
#################################################################################

import pynput
from pynput.keyboard import Key, Controller
import sys
import time

keyboard = Controller()

# Switch to previous focus window
keyboard.press(Key.alt_l)
keyboard.press(Key.tab)
keyboard.release(Key.tab)
keyboard.release(Key.alt_l)

# Wait one second to switch focus between windows
time.sleep(1)

def pressKey(myCharacter):
    myKey = myCharacter

    keyboard.press(myKey)
    keyboard.release(myKey)

def processString(characterString, keyPressInterval):
    print("Start keypress")
    for myKey in list(characterString):
        if keyPressInterval is not None:
            time.sleep(keyPressInterval)
        pressKey(myKey)

# Set interval and input string
keyPressInterval = None
inputString = sys.argv[len(sys.argv)-1]

# If -int argument given, set keyPressInterval
if "-int" in sys.argv:
    index = sys.argv.index("-int") + 1
    keyPressInterval = float(sys.argv[index])

# if -s argument given, treat input as string to parse
if "-s" in sys.argv:
    processString(inputString, keyPressInterval)
else:
    with open(inputString, 'r') as f:
        fileContent = f.read()
        processString(inputString, keyPressInterval)