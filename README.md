# Asset Docx Generator

Generate formatted DOCX reports from Google Sheets using a local Word template

---

## Quick Start

### 1. Clone & enter project

```bash
git clone https://github.com/nonGratis/asset-acts
cd asset-acts
```

### 2. Create Python virtual environment

```bash
python -m venv venv
source .venv/Scripts/activate   # Linux / macOS
                                # or
.venv\Scripts\activate          # Windows
```

### 3. Install requirements

```bash
pip install -r requirements.txt
```

### 4. Configure environment

Copy `.env.example` to `.env` and update with your values:

```bash
cp .env.example .env
```

Edit `.env` with your credentials:

```env
# Google Service Account Credentials
GOOGLE_CREDS_PATH=credentials.json

# Google Spreadsheet IDs
ASSETS_SHEET_ID=your_assets_spreadsheet_id_here
DEPARTMENTS_SHEET_ID=your_departments_spreadsheet_id_here

# Google Drive
SHARED_DRIVE_ID=your_shared_drive_id_here
```

### 5. Prepare files

* **`credentials.json`** → Google API service account file (path from `GOOGLE_CREDS_PATH`)
* **`template.docx`** → Word template with placeholders

### 6. Run

```bash
python main.py
```

