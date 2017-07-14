"""Functionality to switch to the specified window."""

import win32api
import win32gui
import win32process
import win32con
import os
import xml.sax.saxutils
import unicodedata

from enso.commands import CommandObject
import logging

from ctypes import c_int, c_wchar_p, WinDLL, create_unicode_buffer, GetLastError
from ctypes.wintypes import HWND  # , RECT, POINT


def _stdcall(dllname, restype, funcname, *argtypes):
    """
    A decorator for a generator.

    The decorator loads the specified dll, retrieves the
    function with the specified name, set its restype and argtypes,
    it then invokes the generator which must yield twice: the first
    time it should yield the argument tuple to be passed to the dll
    function (and the yield returns the result of the call).
    It should then yield the result to be returned from the
    function call.
    """
    def decorate(func):
        api = getattr(WinDLL(dllname), funcname)
        api.restype = restype
        api.argtypes = argtypes

        def decorated(*args, **kw):
            iterator = func(*args, **kw)
            nargs = iterator.next()
            if not isinstance(nargs, tuple):
                nargs = (nargs,)
            try:
                res = api(*nargs)
            except Exception, e:
                return iterator.throw(e)
            return iterator.send(res)
        return decorated
    return decorate


def nonzero(result):
    """
    Check if nonzero.

    If the result is zero, and GetLastError() returns a non-zero error code and
    raise a WindowsError
    """
    if result == 0 and GetLastError():
        raise WinError()  # NOQA - ignore flake8: WinError is defined elsewhere
    return result


@_stdcall("user32", c_int, "GetWindowTextLengthW", HWND)
def GetWindowTextLength(hwnd):
    """Get window title length."""
    yield nonzero((yield hwnd,))


@_stdcall("user32", c_int, "GetWindowTextW", HWND, c_wchar_p, c_int)
def GetWindowText(hwnd):
    """Get window title."""
    len = GetWindowTextLength(hwnd)+1
    buf = create_unicode_buffer(len)
    nonzero((yield hwnd, buf, len))
    yield buf.value


class GoCommand(CommandObject):
    """Go to specified window."""

    def __init__(self, parameter=None):
        """Initialization."""
        super(GoCommand, self).__init__()
        self.parameter = parameter

    def on_quasimode_start(self):
        """Initialization 2."""
        def callback(found_win, windows):
            # Determine if the window is application window
            if not win32gui.IsWindow(found_win):
                return True
            # Invisible windows are of no interest
            if not win32gui.IsWindowVisible(found_win):
                return True
            # Also disabled windows we do not want
            if not win32gui.IsWindowEnabled(found_win):
                return True
            exstyle = win32gui.GetWindowLong(found_win, win32con.GWL_EXSTYLE)
            # AppWindow flag would be good at this point
            if exstyle & win32con.WS_EX_APPWINDOW != win32con.WS_EX_APPWINDOW:
                style = win32gui.GetWindowLong(found_win, win32con.GWL_STYLE)
                # Child window is suspicious
                if style & win32con.WS_CHILD == win32con.WS_CHILD:
                    return True
                parent = win32gui.GetParent(found_win)
                owner = win32gui.GetWindow(found_win, win32con.GW_OWNER)
                # Also window which has parent or owner is probably not an
                # application window
                if parent > 0 or owner > 0:
                    return True
                # Tool windows we also avoid
                # TODO: Avoid tool windows? Make exceptions? Make configurable?
                toolwindow = win32con.WS_EX_TOOLWINDOW
                if exstyle & toolwindow == toolwindow:
                    return True
            # There are some specific windows we do not want to switch to
            win_class = win32gui.GetClassName(found_win)
            if win_class in ["WindowsScreensaverClass", "tooltips_class32"]:
                return True
            # Now we probably have application window

            # Get title
            # Using own GetWindowText, because win32gui.GetWindowText() doesn't
            # return unicode string.
            win_title = GetWindowText(found_win)
            # Removing all accents from characters
            win_title = unicodedata.normalize('NFKD', win_title)
            win_title = win_title.encode('ascii', 'ignore')

            # Get PID so we can get process name
            _, process_id = win32process.GetWindowThreadProcessId(found_win)
            process = ""
            try:
                # Get process name
                query_info = win32con.PROCESS_QUERY_INFORMATION
                phandle = win32api.OpenProcess(
                    query_info | win32con.PROCESS_VM_READ,
                    False,
                    process_id)
                pexe = win32process.GetModuleFileNameEx(phandle, 0)
                pexe = os.path.normcase(os.path.normpath(pexe))
                # Remove extension
                process, _ = os.path.splitext(os.path.basename(pexe))
            except Exception, e:  # NOQA - ignore flake8: exception too generic
                pass

            # Add hwnd and title to the list
            windows.append((found_win, "%s: %s" % (process, win_title)))
            return True

        # Compile list of application windows
        self.windows = []
        win32gui.EnumWindows(callback, self.windows)

        lowered = map(lambda x: str(x[1]).lower(), self.windows)
        self.valid_args = sorted(lowered)

    def __call__(self, ensoapi, window=None):
        """Executed when the class is called."""
        if window is None:
            return None
        logging.debug("Go to window '%s'" % window)
        for hwnd, title in self.windows:
            title = xml.sax.saxutils.escape(title).lower()
            if title == window:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

                nomove = win32con.SWP_NOMOVE
                notop = win32con.HWND_NOTOPMOST
                top = win32con.HWND_TOPMOST

                swp_sz = nomove | win32con.SWP_NOSIZE
                swp_sz2 = win32con.SWP_SHOWWINDOW | nomove | win32con.SWP_NOSIZE

                win32gui.SetWindowPos(hwnd, notop, 0, 0, 0, 0, swp_sz)
                win32gui.SetWindowPos(hwnd, top, 0, 0, 0, 0, swp_sz)
                win32gui.SetWindowPos(hwnd, notop, 0, 0, 0, 0, swp_sz2)
        return hwnd

cmd_go = GoCommand()
cmd_go.valid_args = []
