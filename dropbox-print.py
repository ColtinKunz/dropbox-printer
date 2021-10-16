import dropbox

from os import getenv
from dotenv import load_dotenv


# Load .env file
load_dotenv()
# Connect to Dropbox API
dbx = dropbox.Dropbox(getenv("DROPBOX_API_KEY"))
