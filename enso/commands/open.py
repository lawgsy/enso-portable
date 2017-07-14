"""Functionality to open applications (and memorize new shortcuts)."""

from win32com.shell import shell, shellcon
from enso.platform.win32.scriptfolder import get_script_folder_name
import os
import re
import win32api
import win32con
import pythoncom
import logging


unlearn_open_undo = []

my_documents_dir = shell.SHGetFolderPath(0, shellcon.CSIDL_PERSONAL, 0, 0)
dir_name = "Enso's Learn As Open Commands"
LEARN_AS_DIR = os.path.join(get_script_folder_name(), dir_name)

# Check if Learn-as dir exist and create it if not
if (not os.path.isdir(LEARN_AS_DIR)):
    os.makedirs(LEARN_AS_DIR)

SHORTCUT_TYPE_EXECUTABLE = 'x'
SHORTCUT_TYPE_FOLDER = 'f'
SHORTCUT_TYPE_URL = 'u'
SHORTCUT_TYPE_DOCUMENT = 'd'
SHORTCUT_TYPE_CONTROL_PANEL = 'c'


def _cpl_exists(cpl_name):
    path1 = os.path.expandvars("${WINDIR}\\%s.cpl" % cpl_name)
    path2 = os.path.expandvars("${WINDIR}\\system32\\%s.cpl" % cpl_name)
    return os.path.isfile(path1) or os.path.isfile(path2)


control_panel_applets = [i[:3] for i in (
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"control panel",
        "rundll32.exe shell32.dll,Control_RunDLL"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"accessibility options (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL access.cpl"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"add or remove programs (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"display properties (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL desk.cpl"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"regional and language options (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL intl.cpl"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"game controllers (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL joy.cpl"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"mouse properties (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL main.cpl @0"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"keyboard properties (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL main.cpl @1"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"microsoft exchange profiles (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL mlcfg32.cpl",
        _cpl_exists("mlcfg32")),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"sounds and audio devices (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"modem properties (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL modem.cpl",
        _cpl_exists("modem")),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"network connections (control panel)",
        "RUNDLL32.exe SHELL32.DLL,Control_RunDLL NCPA.CPL"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"system properties (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"device manager (control panel)",
        "devmgmt.msc"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"disk management (control panel)",
        "diskmgmt.msc"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"scanners and cameras (control panel)",
        "control.exe sticpl.cpl"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"removable storage (control panel)",
        "ntmsmgr.msc"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"hardware profiles (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,2"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"advanced system properties (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,3"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"date and time (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL timedate.cpl"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"add new hardware wizard (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL hdwwiz.cpl @1"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"add printer wizard (control panel)",
        "rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL AddPrinter"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"printers and faxes (control panel)",
        "rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL PrintersFolder"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"fonts (control panel)",
        "rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL FontsFolder"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"windows firewall (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL firewall.cpl"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"speech properties (control panel)",
        ("rundll32.exe shell32.dll,Control_RunDLL "
         "\"${COMMONPROGRAMFILES}\\Microsoft Shared\\Speech\\sapi.cpl\""),
        os.path.isfile(os.path.expandvars(
            "${COMMONPROGRAMFILES}\\Microsoft Shared\\Speech\\sapi.cpl"
        ))),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"internet options (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"odbc data source administrator (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL odbccp32.cpl"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"power options (control panel)",
        "rundll32.exe shell32.dll,Control_RunDLL powercfg.cpl"),
    (SHORTCUT_TYPE_CONTROL_PANEL,
        u"bluetooth properties (control panel)",
        "control.exe bhtprops.cpl",
        _cpl_exists("bhtprops")),
) if len(i) < 4 or i[3]]


class _PyShortcut():
    def __init__(self, base):
        self._base = base
        self._base_loaded = False
        self._shortcut_type = None

    def load(self, filename=None):
        if filename:
            self._filename = filename
        try:
            intf = self._base.QueryInterface(pythoncom.IID_IPersistFile)
            intf.Load(self._filename)
        except:
            err = "Error loading shell-link for file %s" % self._filename
            logging.error(err)
        self._base_loaded = True

    def save(self, filename=None):
        if filename:
            self._filename = filename
        intf = self._base.QueryInterface(pythoncom.IID_IPersistFile)
        intf.Save(self._filename, 0)

    def get_filename(self):
        return self._filename

    def get_type(self):
        if not self._base_loaded:
            raise Exception(("Shortcut data has not been loaded yet. "
                             "Use load(filename) before using get_type()"))

        name, ext = os.path.splitext(self._filename)
        if ext.lower() == '.lnk':
            file_path = self._base.GetPath(0)
            if file_path and file_path[0]:
                if os.path.isdir(file_path[0]):
                    self._shortcut_type = SHORTCUT_TYPE_FOLDER
                else:
                    extension = os.path.splitext(file_path[0])[1].lower()
                    if extension in ['.exe', '.com', '.cmd', '.bat']:
                        self._shortcut_type = SHORTCUT_TYPE_EXECUTABLE
                    else:
                        self._shortcut_type = SHORTCUT_TYPE_DOCUMENT
            else:
                self._shortcut_type = SHORTCUT_TYPE_DOCUMENT
        elif ext.lower() == '.url':
            self._shortcut_type = SHORTCUT_TYPE_URL
        else:
            self._shortcut_type = SHORTCUT_TYPE_DOCUMENT
        return self._shortcut_type

    def __getattr__(self, name):
        if name != "_base":
            return getattr(self._base, name)


class PyShellLink(_PyShortcut):
    """Shell shortcut."""

    def __init__(self):
        """Initialization method."""
        base = pythoncom.CoCreateInstance(
            shell.CLSID_ShellLink,
            None,
            pythoncom.CLSCTX_INPROC_SERVER,
            shell.IID_IShellLink
        )
        _PyShortcut.__init__(self, base)


class PyInternetShortcut(_PyShortcut):
    """URL shortcut."""

    def __init__(self):
        """Initialization method."""
        base = pythoncom.CoCreateInstance(
            shell.CLSID_InternetShortcut,
            None,
            pythoncom.CLSCTX_INPROC_SERVER,
            shell.IID_IUniformResourceLocator
        )
        _PyShortcut.__init__(self, base)


def expand_path_variables(file_path):
    """Expand system path variables in given path."""
    import re
    re_env = re.compile(r'%\w+%')

    def expander(mo):
        return os.environ.get(mo.group()[1:-1], 'UNKNOWN')

    return os.path.expandvars(re_env.sub(expander, file_path))


def displayMessage(msg):
    """Display given message in Enso."""
    import enso.messages
    enso.messages.displayMessage("<p>%s</p>" % msg)


ignored = re.compile("(uninstall|read ?me|faq|f.a.q|help)", re.IGNORECASE)


def get_shortcuts(directory):
    """Return all shortcuts in a given directory."""
    shortcuts = []
    sl = PyShellLink()
    for dirpath, dirnames, filenames in os.walk(directory):
        for filename in filenames:
            if ignored.search(filename):
                continue
            name, ext = os.path.splitext(filename)
            if not ext.lower() in (".lnk", ".url"):
                continue
            shortcut_type = SHORTCUT_TYPE_DOCUMENT
            if ext.lower() == ".lnk":
                sl.load(os.path.join(dirpath, filename))
                shortcut_type = sl.get_type()
            elif ext.lower() == ".url":
                shortcut_type = SHORTCUT_TYPE_URL
            shortcuts.append(
                (shortcut_type, name.lower(), os.path.join(dirpath, filename))
            )
    return shortcuts


def reload_shortcuts_map():
    """Reload shortcut mapping."""
    desk_dir = shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOPDIRECTORY, 0, 0)
    quick_launch_dir = os.path.join(
        shell.SHGetFolderPath(0, shellcon.CSIDL_APPDATA, 0, 0),
        "Microsoft",
        "Internet Explorer",
        "Quick Launch")
    start_menu_dir = shell.SHGetFolderPath(0, shellcon.CSIDL_STARTMENU, 0, 0)
    cmmn_dir = shellcon.CSIDL_COMMON_STARTMENU
    common_start_menu_dir = shell.SHGetFolderPath(0, cmmn_dir, 0, 0)

    shortcuts = get_shortcuts(LEARN_AS_DIR) + \
        get_shortcuts(desk_dir) + \
        get_shortcuts(quick_launch_dir) + \
        get_shortcuts(start_menu_dir) + \
        get_shortcuts(common_start_menu_dir) + \
        control_panel_applets
    return dict((s[1], s) for s in shortcuts)

shortcuts_map = reload_shortcuts_map()


def cmd_open(ensoapi, target):
    """Continue typing to open an application or document."""
    displayMessage(u"Opening <command>%s</command>..." % target)

    try:
        global shortcuts_map
        shortcut_type, shortuct_id, file_path = shortcuts_map[target]
        file_path = os.path.normpath(expand_path_variables(file_path))
        logging.info("Executing '%s'" % file_path)

        if shortcut_type == SHORTCUT_TYPE_CONTROL_PANEL:
            if " " in file_path:
                executable = file_path[0:file_path.index(' ')]
                params = file_path[file_path.index(' ')+1:]
            else:
                executable = file_path
                params = None
            try:
                win32api.ShellExecute(
                    0,
                    'open',
                    executable,
                    params,
                    None,
                    win32con.SW_SHOWDEFAULT
                )
                # rcode = win32api.ShellExecute(
                #     0,
                #     'open',
                #     executable,
                #     params,
                #     None,
                #     win32con.SW_SHOWDEFAULT)
            except Exception, e:
                logging.error(e)
        else:
            os.startfile(file_path)

        return True
    except Exception, e:
        logging.error(e)
        return False

cmd_open.valid_args = [s[1] for s in shortcuts_map.values()]


def cmd_open_with(ensoapi, application):
    """Open currently selected file(s)/folder with the specified application."""
    seldict = ensoapi.get_selection()
    if seldict.get('files'):
        file = seldict['files'][0]
    elif seldict.get('text'):
        file = seldict['text'].strip()
    else:
        file = None

    if not (file and (os.path.isfile(file) or os.path.isdir(file))):
        ensoapi.display_message(u"No file or folder is selected")
        return

    displayMessage(u"Opening <command>%s</command>..." % application)

    global shortcuts_map
    try:
        print shortcuts_map[application][2]
        print shortcuts_map[application]
        executable = expand_path_variables(shortcuts_map[application][2])
    except:
        print application
        print shortcuts_map.keys()
        print shortcuts_map.values()
    try:
        win32api.ShellExecute(
            0,
            'open',
            executable,
            '"%s"' % file,
            os.path.dirname(file),
            win32con.SW_SHOWDEFAULT)
    except Exception, e:
        logging.error(e)

cmd_open_with.valid_args = [
    s[1] for s in shortcuts_map.values() if s[0] == SHORTCUT_TYPE_EXECUTABLE
]


def is_url(text):
    """Check if given text contains a valid URL."""
    urlfinders = [
        re.compile((
            "([0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}|(((news|telnet|"
            "nttp|file|http|ftp|https)://)|(www|ftp)[-A-Za-z0-9]*\\.)[-A-Za-z0-"
            "9\\.]+)(:[0-9]*)?/[-A-Za-z0-9_\\$\\.\\+\\!\\*\\(\\),;:@&=\\?/~\\#"
            "\\%]*[^]'\\.}>\\),\\\"]"
        )),
        re.compile((
            "([0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}|(((news|telnet|"
            "nttp|file|http|ftp|https)://)|(www|ftp)[-A-Za-z0-9]*\\.)[-A-Za-z0-"
            "9\\.]+)(:[0-9]*)?"
        )),
        re.compile((
            "(~/|/|\\./)([-A-Za-z0-9_\\$\\.\\+\\!\\*\\(\\),;:@&=\\?/~\\#\\%]|"
            "\\\\)+"
        )),
        re.compile("'\\<((mailto:)|)[-A-Za-z0-9\\.]+@[-A-Za-z0-9\\.]+"),
    ]

    for urltest in urlfinders:
        if urltest.search(text, re.I):
            return True

    return False


def cmd_learn_as_open(ensoapi, name):
    """Learn to open a document or application."""  # as {name}
    if name is None:
        displayMessage(u"You must provide a name")
        return
    seldict = ensoapi.get_selection()
    if seldict.get('files'):
        file = seldict['files'][0]
    elif seldict.get('text'):
        file = seldict['text'].strip()
    else:
        ensoapi.display_message(u"No file is selected")
        return

    if not (os.path.isfile(file) or os.path.isdir(file) or is_url(file)):
        displayMessage(
            u"Selection represents no existing file, folder or URL.")
        return

    file_name = name.replace(":", "").replace("?", "").replace("\\", "")
    file_path = os.path.join(LEARN_AS_DIR, file_name)

    if os.path.isfile(file_path + ".url") or os.path.isfile(file_path + ".lnk"):
        displayMessage((
            "<command>open %s</command> already exists. "
            "Please choose another name."
        ) % name)
        return

    if is_url(file):
        shortcut = PyInternetShortcut()

        shortcut.SetURL(file)
        intf = shortcut.QueryInterface(pythoncom.IID_IPersistFile)
        intf.Save(file_path + ".url", 0)
    else:
        shortcut = PyShellLink()

        shortcut.SetPath(file)
        shortcut.SetWorkingDirectory(os.path.dirname(file))
        shortcut.SetIconLocation(file, 0)

        intf = shortcut.QueryInterface(pythoncom.IID_IPersistFile)
        intf.Save(file_path + ".lnk", 0)

    global shortcuts_map
    shortcuts_map = reload_shortcuts_map()
    cmd_open.valid_args = [s[1] for s in shortcuts_map.values()]
    cmd_open_with.valid_args = [
        s[1] for s in shortcuts_map.values() if s[0] == SHORTCUT_TYPE_EXECUTABLE
    ]
    cmd_unlearn_open.valid_args = [s[1] for s in shortcuts_map.values()]

    displayMessage(u"<command>open %s</command> is now a command" % name)


def cmd_unlearn_open(ensoapi, name):
    """Unlearn "open (name)" command."""  # {name}
    file_path = os.path.join(LEARN_AS_DIR, name)
    if os.path.isfile(file_path + ".lnk"):
        sl = PyShellLink()
        sl.load(file_path + ".lnk")
        unlearn_open_undo.append([name, sl])
        os.remove(file_path + ".lnk")
    elif os.path.isfile(file_path + ".url"):
        sl = PyInternetShortcut()
        sl.load(file_path + ".url")
        unlearn_open_undo.append([name, sl])
        os.remove(file_path + ".url")

    global shortcuts_map
    shortcuts_map = reload_shortcuts_map()
    cmd_open.valid_args = [s[1] for s in shortcuts_map.values()]
    cmd_open_with.valid_args = [
        s[1] for s in shortcuts_map.values() if s[0] == SHORTCUT_TYPE_EXECUTABLE
    ]
    cmd_unlearn_open.valid_args = [s[1] for s in shortcuts_map.values()]
    displayMessage(u"Unlearned <command>open %s</command>" % name)


cmd_unlearn_open.valid_args = [s[1] for s in shortcuts_map.values()]


def cmd_undo_unlearn(ensoapi):
    """Undo the last "unlearn open" command."""
    if len(unlearn_open_undo) > 0:
        name, sl = unlearn_open_undo.pop()
        sl.save()
        displayMessage(("Undo successful."
                        "<command>open %s</command> is now a command" % name))
    else:
        ensoapi.display_message(u"There is nothing to undo")

if __name__ == "__main__":
    import doctest
    doctest.testmod()
