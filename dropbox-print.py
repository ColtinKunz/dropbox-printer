import os
import dropbox

from keyboard import press
from shutil import copyfileobj
from time import sleep
from pathlib import Path
from dotenv import load_dotenv
from contextlib import closing
from os import getenv, startfile, remove

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
                        print(f"Added to print queue: {temp_path}")
                        sleep(2)
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
