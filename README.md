# BRANDO — PDF Bill Generator (Pro)

Features:
- Customer name & address fields
- Auto invoice number (manual override allowed)
- Optional logo upload (PNG/JPG)
- Save PDFs to ./invoices and keep a history list
- Viewer page with embedded PDF and a Print button
- Inline PDF to avoid IDM interception
- Clean blue theme

## Run
python -m venv .venv
. .venv/Scripts/Activate.ps1   # (Windows) or: source .venv/bin/activate
pip install -r requirements.txt
python app.py
Open http://127.0.0.1:5000

