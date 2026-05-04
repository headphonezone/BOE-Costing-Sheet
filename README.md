📦 BOE → Excel Auto-Filler
A Streamlit web app that automatically parses Bill of Entry (BOE) PDFs from ICEGATE and fills a pre-built Excel costing template — eliminating manual data entry for Focal shipments.

🚀 Features
Upload any Focal BOE PDF and get a filled Excel file in seconds

Auto-detects any number of items (not fixed at 10)

Extracts exchange rate, invoice number, date, freight, and insurance

Parses all item descriptions, unit prices (USD), and quantities

Reads BCD, SWS, and IGST duty amounts from Part-III duty pages

Licence-based BCD logic: when BCD = 0 for an item, automatically sums all DEBIT DUTY values (col 11) from the F. LICENCE DETAILS table for that item number and writes the total as BCD

Fills both C-SHEET and D-DETAILS tabs with correct formulas

Shows a live BCD Source preview tab before download so you can verify every item

🖥️ App Preview
Tab	What it shows
📦 Items	All extracted item descriptions, qty, and USD rate
🏦 Duties	BCD, SWS, IGST per item from Part-III
📜 Licences	All rows from F. LICENCE DETAILS (ITMSNO + LIC NO + DEBIT DUTY)
🔍 BCD Source	Per-item breakdown of whether BCD came from cash or licence
📁 Project Structure
text
boe-excel-filler/
├── boe_excel_filler.py   # Main Streamlit app
├── README.md             # This file
└── requirements.txt      # Python dependencies
⚙️ Setup & Installation
1. Clone the repository
bash
git clone https://github.com/your-username/boe-excel-filler.git
cd boe-excel-filler
2. Create a virtual environment (recommended)
bash
python -m venv venv
source venv/bin/activate        # macOS / Linux
venv\Scripts\activate           # Windows
3. Install dependencies
bash
pip install -r requirements.txt
4. Add your Excel template
Open boe_excel_filler.py and paste your base64-encoded Excel template into the TEMPLATE_B64 constant:

python
TEMPLATE_B64 = "your_base64_string_here"
To generate the base64 string from your .xlsx file:

python
import base64
with open("your_template.xlsx", "rb") as f:
    print(base64.b64encode(f.read()).decode())
5. Run the app
bash
streamlit run boe_excel_filler.py
The app opens at http://localhost:8501 in your browser.

📋 requirements.txt
text
streamlit
pdfplumber
openpyxl
pandas
🔧 How It Works
PDF Parsing Pipeline
text
BOE PDF
  │
  ├── Page 1        → Exchange rate, BE number, BE date
  ├── Part-II pages → Invoice number, date, freight, insurance, item list
  ├── Part-III pages→ BCD, SWS, IGST per item (x-position based extraction)
  └── Licence page  → F. LICENCE DETAILS table (ITMSNO + DEBIT DUTY col 11)
BCD Logic (Key Rule)
Condition	BCD Written to Excel
BCD > 0 from Part-III	Parsed cash BCD value
BCD = 0 AND licence rows exist for this item	Sum of all DEBIT DUTY (col 11) values for that ITMSNO
BCD = 0 AND no licence rows	0
Example (from a real BOE):

Item	Licence No	Debit Duty	BCD Written
1	2602025033	19.98	↘
1	2603045274	14622.52	14642.50
2	2603045274	2239.60	2239.60
3	2603045274	3284.70	3284.70
Excel Output
C-SHEET tab (costing sheet):

Exchange rate, invoice details, freight, insurance filled in

Item descriptions, QTY, USD rate populated

INR value, custom duty, cost per piece, +2% margin, total cost formulas written dynamically

D-DETAILS tab (duty breakdown):

BCD / SWS / IGST per item

Paid/Cyber Receipt formula (SWS + IGST when BCD is licence-based)

Full licence right-hand table with ITMSNO, Licence No, Debit Duty

Grand totals, SWS/IGST/Custom Duty summary rows, and breakdown rows

🛠️ Licence Extraction Method
The F. LICENCE DETAILS table is parsed using x-position based word extraction via pdfplumber — not a regex — making it robust to OCR noise and variable spacing:

Finds the header row containing ITMSNO or DEBIT DUTY

For each data row starting with 1 (INVSNO = 1):

ITMSNO = 2nd numeric token on the row

DEBIT DUTY = last numeric token on the row (col 11 is always rightmost)

LIC NO = first token with 9+ digits

Falls back to line-split text parsing if the header row is not found

📌 Notes
The Excel template must have two sheets named exactly C-SHEET and D-DETAILS

BOE PDF must be a standard ICEGATE-generated PDF

Multiple licence rows per item are fully supported and summed correctly

If a BOE has no licence table (all BCD values > 0), the app works normally using cash BCD values

🏢 Built For
Ferrari Video | Focal Shipments — Internal costing automation tool.

🐛 Troubleshooting
Problem	Fix
Licence rows not detected	Check the 📜 Licences tab — if empty, the licence page may be before page 6. Adjust range(5, ...) in the PDF loop
Wrong BCD values	Check 🔍 BCD Source tab to see what was parsed
Items not extracted	Part-II pattern may differ — enable st.exception(e) for full traceback
Template error	Ensure sheet names are exactly C-SHEET and D-DETAILS
