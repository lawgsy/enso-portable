"""Enso system commands."""
import os
import sys
import time
import subprocess

import enso
import enso.config
from enso.messages import displayMessage


def cmd_enso(ensoapi, cmd):
    """
    Some commands related to Enso itself.

    <i>Control Enso with Enso.</i><br /><br />
    <b>Possible parameters:</b> quit, restart, about
    """
    if cmd == 'quit':
        title = "<p>Closing <command>Enso</command>...</p>"
        caption = "<caption>enso</caption>"
        displayMessage("%s%s" % (title, caption))
        time.sleep(1)
        PID = os.getpid()
        os.system("TASKKILL /F /PID %s" % PID)  # if run by .exe
        sys.exit(0)  # if through python normally
    if cmd == 'restart':
        subproc = [enso.enso_executable, "--restart " + str(os.getpid())]
        subprocess.Popen(subproc)
        title = "<p>Restarting <command>Enso</command>...</p>"
        caption = "<caption>enso</caption>"
        displayMessage("%s%s" % (title, caption))
        time.sleep(1)
        PID = os.getpid()
        os.system("TASKKILL /F /PID %s" % PID)  # if run by .exe
        sys.exit(0)  # if through python normally
    elif cmd == 'about':
        displayMessage(enso.config.ABOUT_BOX_XML)


cmd_enso.valid_args = ['about', 'quit', 'restart']
