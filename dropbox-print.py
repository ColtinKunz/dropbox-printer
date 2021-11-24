import re
import os
import dropbox
import win32con

from keyboard import press
from shutil import copyfileobj
from time import sleep
from pathlib import Path
from dotenv import load_dotenv
from contextlib import closing
from os import getenv, startfile, remove

from win32gui import (
    SetForegroundWindow,
    FindWindow,
    GetWindowText,
    EnumWindows,
    PostMessage,
)


class WindowMgr:
    """Encapsulates some calls to the winapi for window management"""

    def __init__(self):
        """Constructor"""
        self._handle = None

    def find_window(self, class_name, window_name=None):
        """find a window by its class_name"""
        self._handle = FindWindow(class_name, window_name)

    def close_window(self):
        PostMessage(self._handle, win32con.WM_CLOSE, 0, 0)

    def _window_enum_callback(self, hwnd, wildcard):
        """Pass to win32gui.EnumWindows() to check all the opened windows"""
        if re.match(wildcard, str(GetWindowText(hwnd))) is not None:
            self._handle = hwnd

    def find_window_wildcard(self, wildcard):
        """find a window whose title matches the wildcard regex"""
        self._handle = None
        EnumWindows(self._window_enum_callback, wildcard)

    def set_foreground(self):
        """put the window in the foreground"""
        SetForegroundWindow(self._handle)


w = WindowMgr()
# Load .env file
load_dotenv()
# Connect to Dropbox API
dbx = dropbox.Dropbox(getenv("DROPBOX_API_KEY"))

print_dir = None
print_dir_name = "Print"
for d in dbx.sharing_list_folders().entries:
    if d.name == print_dir_name:
        print_dir = d

if print_dir is None:
    raise FileNotFoundError(
        f"Can't find '{print_dir_name}' shared folder on Dropbox."
    )

# Temporary download path
path = Path("F:/Repos/dropbox-printer")
path = Path(".")

print("Running Dropbox Printer...")
while True:
    # For each entry in print directory on Dropbox
    for entry in dbx.files_list_folder(print_dir.path_lower).entries:
        # Grab response and make sure it is closed
        with closing(dbx.files_download(entry.path_lower)[1]) as r:
            # Make sure the status code is good
            if r.status_code == 200:
                # Save dropbox file temporarily to path
                temp_path = f"{path}{Path(entry.path_lower)}"
                with open(temp_path, "wb") as local_file:
                    r.raw.decode_content = True
                    copyfileobj(r.raw, local_file)
                try:
                    if os.path.isfile(temp_path):
                        print(f"Printing: {temp_path}")
                        startfile(temp_path, "print")
                        file_ending = temp_path.rsplit(".", 1)[1]
                        # pull up window
                        hwnd = None
                        sleep(2)
                        if (
                            file_ending == "png"
                            or file_ending == "jpg"
                            or file_ending == "jpeg"
                        ):
                            w.find_window_wildcard(".*Print.*")
                            w.set_foreground()
                        elif file_ending == "pdf":
                            w.find_window_wildcard(".*Acrobat Pro DC.*")
                            w.close_window()
                        print(f"Added to print queue: {temp_path}")
                        press("enter")
                        dbx.files_delete(entry.path_lower)
                    else:
                        raise FileNotFoundError(
                            f"Can't find '{temp_path}' on system."
                        )
                except Exception:
                    local_file.close()
                    remove(temp_path)
                    raise

                remove(temp_path)
        sleep(3)
    sleep(5)
