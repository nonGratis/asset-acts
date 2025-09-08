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

### 4. Prepare files

* **`credentials.json`** → Google API service account file.
* **`template.docx`** → Word template with placeholders

### 5. Run

```bash
python main.py
```

