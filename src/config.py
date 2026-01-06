import os
import sys
import logging
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Service Account Configuration and Google Sheets Configuration
SERVICE_ACCOUNT_KEYFILE = os.getenv("GOOGLE_CREDS_PATH", "credentials.json")
ASSETS_SHEET_NAME = os.getenv("ASSETS_SHEET_NAME", "")
ASSETS_SPREADSHEET_ID = os.getenv("ASSETS_SHEET_ID", "")

DEPARTMENTS_SPREADSHEET_ID = os.getenv("DEPARTMENTS_SHEET_ID", "")
DEPARTMENTS_SHEET_NAME = os.getenv("DEPARTMENTS_SHEET_NAME", "")

# Google Drive Configuration
SHARED_DRIVE_ID = os.getenv("SHARED_DRIVE_ID", "")

# Local Output Directories
OUTPUT_LOCAL_DIR_DOC = "docs"
OUTPUT_LOCAL_DIR_PDF = "pdfs"
OUTPUT_LOCAL_DIR_JPEG = "jpegs"

# Asset Sheet Columns (1-based indexing)
COL_ID = 1
COL_NAME = 3
COL_DATE = 4
COL_INVENTORY_NUMBER = 5
COL_UNIT = 6
COL_QUANTITY = 7
COL_PRICE = 9
COL_OWNERS = 10
COL_GENERATE_FLAG = 11

# Department Sheet Columns (1-based indexing)
DEPT_COL_CODE = 1
DEPT_COL_STATUS = 2
DEPT_COL_POSITION = 3
DEPT_COL_FULLNAME = 4
DEPT_COL_NORMALIZED = 5

# Document Generation Settings
FILE_NAME_PATTERN = "Акт. {deptname}"
TEMPLATE_PATH = "template.docx"

# Number Formatting Settings
THOUSAND_SEPARATOR = " "
DECIMAL_SEPARATOR = ","
CURRENCY_SUFFIX = ""
ALLOW_ROUNDING_ADJUST = True

# Google API Scopes
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]

# Logging Configuration
logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s: %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)

# Shared logger instance
log = logging.getLogger(__name__)
