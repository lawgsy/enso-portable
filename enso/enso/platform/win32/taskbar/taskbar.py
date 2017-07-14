""" Handle taskbar icon, menu, functionality """
# import enso.config
import logging
import os

import win32api
import win32con
import win32gui
import winerror

from inspect import getargspec


class SysTrayIcon(object):
    """ Class representing the tray icon """
    MENU_ITEM_ID_ABOUT = 1023
    MENU_ITEM_ID_EXIT = 1024
    free_menu_id = 1025

    def __init__(self,
                 icon,
                 hover_text,
                 menu_options,
                 on_quit=None,
                 default_menu_index=None,
                 window_class_name="EnsoTrayWndClass",):

        self.default_icon = icon
        self.hover_text = hover_text
        self.notify_id = None
        if on_quit:
            self.on_quit = on_quit
        self.custom_menu_items = {}

        self.WM_ONLOAD = win32gui.RegisterWindowMessage("SystrayOnLoad")
        self.WM_CREATED = win32gui.RegisterWindowMessage("TaskbarCreated")
        message_map = {
            self.WM_CREATED : self._on_restart,
            self.WM_ONLOAD: self._on_load,
            win32con.WM_DESTROY: self._on_destroy,
            win32con.WM_COMMAND: self._on_command,
            win32con.WM_USER + 20 : self._on_taskbar_notify,
        }

        # Register the Window class.
        self._window_class = win32gui.WNDCLASS()
        self._window_class.hInstance = win32api.GetModuleHandle(None)
        self._window_class.lpszClassName = window_class_name
        self._window_class.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW
        self._window_class.hCursor = win32api.LoadCursor(0, win32con.IDC_ARROW)
        self._window_class.hbrBackground = win32con.COLOR_WINDOW
        self._window_class.lpfnWndProc = message_map # could also specify a wndproc.

        # Don't blow up if class already registered to make testing easier
        try:
            self.class_atom = win32gui.RegisterClass(self._window_class)
        except win32gui.error, err_info:
            if err_info.winerror != winerror.ERROR_CLASS_ALREADY_EXISTS:
                raise

        # Create the helper Window
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow(
            self.class_atom,
            window_class_name, #Title same as class-name
            style,
            0,
            0,
            win32con.CW_USEDEFAULT,
            win32con.CW_USEDEFAULT,
            0,
            0,
            self._window_class.hInstance,
            None)
        win32gui.UpdateWindow(self.hwnd)

        self._create_icons()


    def _create_icons(self):
        # Try and find a custom icon
        hinst = win32api.GetModuleHandle(None)

        hicon = None
        if os.path.isfile(self.default_icon):
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            # try:
            hicon = win32gui.LoadImage(hinst,
                                       self.default_icon,
                                       win32con.IMAGE_ICON,
                                       0,
                                       0,
                                       icon_flags)
            # except Exception:
            #     logging.error("Can't load icon file - using default.")

        if not hicon:
            logging.error("Can't find icon file - using default.")
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

        self.hicon = hicon

        flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
        nid = (self.hwnd, 0, flags, win32con.WM_USER + 20, hicon, self.hover_text)
        self.notify_id = nid
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
        except win32gui.error:
            # This is common when windows is starting, and this code is hit
            # before the taskbar has been created.
            logging.warning("Failed to add the taskbar icon - is explorer running?")
            # but keep running anyway - when explorer starts, we get the
            # TaskbarCreated message.

    def _on_load(self, hwnd, msg, wparam, lparam):
        pass

    def _on_restart(self, hwnd, msg, wparam, lparam):
        self._create_icons()

    def _on_destroy(self, hwnd, msg, wparam, lparam):
        # TODO: figure out why tray icon doesn't get removed?
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, self.notify_id)
        win32gui.PostQuitMessage(0) # Terminate the current thread.

    def _on_taskbar_notify(self, hwnd, msg, wparam, lparam):
        if lparam == win32con.WM_LBUTTONUP:
            if self.on_leftclick:
                self.on_leftclick(self)
        elif lparam == win32con.WM_LBUTTONDBLCLK:
            #self.change_tooltip("You double-clicked me - goodbye")
            #win32gui.DestroyWindow(self.hwnd)
            if self.on_doubleclick:
                self.on_doubleclick(self)
        elif lparam == win32con.WM_RBUTTONUP:
            if self.on_rightclick:
                self.on_rightclick(self)
        return 1

    def _on_command(self, hwnd, msg, wparam, lparam):
        w_id = win32api.LOWORD(wparam)
        if w_id == self.MENU_ITEM_ID_ABOUT:
            if self.on_about:
                self.on_about(self)
        elif w_id == self.MENU_ITEM_ID_EXIT:
            if self.on_quit:
                self.on_quit(self)
            win32gui.DestroyWindow(self.hwnd)
            win32gui.UnregisterClass(self.class_atom, self._window_class.hInstance)
        elif w_id in self.custom_menu_items.keys():
            if callable(self.custom_menu_items[w_id]['func']):
                try:
                    self.custom_menu_items[w_id]['func'](self)
                except Exception as error:
                    label = self.custom_menu_items[w_id]['text']
                    logging.error("Error executing menu item func: %s", label)
                    logging.error(error)
        else:
            print "Unknown command - %s" % w_id

    def on_quit(self, systray):
        """ Pass through """
        pass

    def on_about(self, systray):
        """ Pass through """
        pass

    def on_doubleclick(self, systray):
        """ Pass through """
        pass

    def on_leftclick(self, systray):
        """ Pass through """
        pass

    def on_rightclick(self, systray):
        """ Display menu """
        #self.change_tooltip("You right clicked me.")
        menu = win32gui.CreatePopupMenu()
        win32gui.AppendMenu(menu, win32con.MF_STRING, 1023, "About")
        if len(self.custom_menu_items) > 0:
            for menu_item in self.custom_menu_items.itervalues():
                if callable(menu_item['func']):
                    try:
                        is_checked = menu_item['func'](self, get_state=True)
                    except Exception as exception:
                        print exception
                else:
                    is_checked = False
                flags = win32con.MF_STRING | (win32con.MF_CHECKED if is_checked else 0)
                win32gui.AppendMenu(menu, flags, menu_item['id'], menu_item['text'])

        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, -1, "")
        win32gui.AppendMenu(menu, win32con.MF_STRING, 1024, "E&xit")
        win32gui.SetMenuDefaultItem(menu, 0, True)
        pos = win32gui.GetCursorPos()
        # See http://bit.ly/ms-api-ref
        win32gui.SetForegroundWindow(self.hwnd)
        win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN, pos[0], pos[1], 0, self.hwnd, None)
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)

    def add_menu_item(self, text, func):
        """ Adds an option to the menu """
        assert len(text) > 0
        assert callable(func)

        check1 = len(getargspec(func)[0]) == 2
        check2 = getargspec(func)[0][1] == 'get_state'
        assert check1 and check2, "Second function argument must be 'get_state'"

        menu_item = { 'id':self.free_menu_id, 'text':text, 'func':func }
        self.custom_menu_items[self.free_menu_id] = menu_item
        self.free_menu_id += 1

    def change_tooltip(self, text):
        """ Set tooltip """
        if self.notify_id:
            message = win32gui.NIM_MODIFY
        else:
            message = win32gui.NIM_ADD
        self.notify_id = (self.hwnd,
                          0,
                          win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
                          win32con.WM_USER + 20,
                          self.hicon,
                          text)
        win32gui.Shell_NotifyIcon(message, self.notify_id)

    # @classmethod
    def main_thread(self):
        """ Main thread """
        win32gui.PumpMessages()

# vim:set tabstop=4 shiftwidth=4 expandtab:
