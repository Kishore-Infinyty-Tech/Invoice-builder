"""Creates the dummy-data workbook 'Invoice Generation Inputs.xlsx' (4 tabs) and the last-number file.
Run: python make_inputs.py [--force]   (writes into Inputs/ and a copy next to this script)"""
import shutil
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
OUT = BASE / "Inputs" / "Invoice Generation Inputs.xlsx"
if OUT.exists() and "--force" not in sys.argv:
    raise SystemExit(f"{OUT} already exists, not overwritten (use --force).")

HEAD = PatternFill("solid", fgColor="1F3864")


def style(ws, widths):
    for c in ws[1]:
        if c.value is not None:
            c.fill, c.font = HEAD, Font(bold=True, color="FFFFFF")
            c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


wb = Workbook()

# ---- tab 1: Invoices To Generate
ws = wb.active
ws.title = "Invoices To Generate"
ws.append(["S.No", "Client Name", "Currency", "Amount", "Item Description", "Notes to put in Invoice",
           "Include in Invoice Name", "HSN/SAC", "PO Number"])
rows = [
    ["Meridian Retail", "INR", 100000, "Data & development services for Oct'26 (1st-15th Oct)",
     "Data & development services for Oct'26 (1st-15th Oct)", "1stHalfOct2026", "998313", ""],
    ["Northwind Analytics", "USD", 2000, "ML/AI support for September 2026",
     "ML/AI support for September 2026 (20 days)", "Oct2026", "", ""],
    ["BlueOcean Logistics", "USD", 1500, "AI Engineer support for September 2026", "$1475 + Wire transfer $25",
     "SepOct2026", "", ""],
    ["Sahyadri Finance", "INR", 250000, "Staff Augmentation - Data Engineer - September 2026",
     "Staff Augmentation - Data Engineer - September 2026", "Sep2026", "998313", "SF/PO/2026-27/A4310/IS40770"],
    ["Kaveri Healthcare", "INR", 75000, "Healthcare analytics support for October 2026",
     "Healthcare analytics support for October 2026", "Oct2026", "998313", ""],
    ["Zenith Exports", "INR", 80000, "Supply chain analytics for October 2026",
     "Supply chain analytics for October 2026", "Oct2026", "998313", ""],
    ["Sahyadri Finance", "INR", 120000, "Staff Augmentation - Data Analyst - October 2026",
     "TEST: PO missing, should show red", "Oct2026", "998313", ""],
    ["Orion Textiles", "INR", 50000, "Consulting services for October 2026",
     "TEST: client not in Client Master, should show red", "Oct2026", "998313", ""],
    ["Greenfield Pharma", "INR", 60000, "Clinical data support for October 2026",
     "TEST: withholding % missing in master, red until you enter it", "Oct2026", "998313", ""],
    ["Pinnacle Advisory", "USD", 900, "Advisory support for October 2026",
     "TEST: inactive client, should show amber", "Oct2026", "998313", ""],
    ["Meridian Retail", "INR", 40000, "Additional data support for Oct'26 (16th-31st Oct)",
     "TEST: PO given but not needed, amber", "2ndHalfOct2026", "998313", "PO-0001"],
    ["Kaveri Healthcare", "INR", 30000, "Training sessions for October 2026",
     "TEST: HSN/SAC missing, should show red", "Oct2026-Training", "", ""],
]
good = [r for r in rows if not r[4].startswith("TEST:")]
bad = [r for r in rows if r[4].startswith("TEST:")]
for i, r in enumerate(good, 1):
    ws.append([i] + r)
style(ws, [7, 24, 10, 11, 48, 48, 22, 11, 34])
# rows that are meant to fail validation live on a separate tab (the tool reads only "Invoices To Generate")
wt = wb.create_sheet("Error Test Rows")
wt.append([c.value for c in ws[1]])
for i, r in enumerate(bad, 1):
    wt.append([i] + r)
style(wt, [7, 24, 10, 11, 48, 48, 22, 11, 34])

# ---- tab 2: Client Master (Service Provider now lives here)
ws = wb.create_sheet("Client Master")
ws.append(["S.No", "Client Name (Short)", "Client Name (Legal)", "Service Provider", "GSTIN", "Place of Supply",
           "GST Type", "GST %", "Payment Terms", "Amount Withheld?", "% Withheld", "PO Needed?", "Active?",
           "Address (optional)"])
for r in [
    [1, "Meridian Retail", "Meridian Retail Private Limited", "IAC", "29AABCM4821F1Z5", "Karnataka (29)", "IGST", 18, "Net15", "Yes", 10, "No", "Yes", "14 MG Road, Bengaluru, Karnataka 560001"],
    [2, "Northwind Analytics", "Northwind Analytics Inc.", "ITGS", "n/a", "Other Territory (97)", "n/a", "n/a", "Net7", "n/a", "n/a", "No", "Yes", "500 Market Street, Austin, TX, USA"],
    [3, "BlueOcean Logistics", "BlueOcean Logistics LLC", "ITGS", "n/a", "Other Territory (97)", "n/a", "n/a", "Net30", "n/a", "n/a", "No", "Yes", "210 Harbor Drive, Seattle, WA, USA"],
    [4, "Sahyadri Finance", "Sahyadri Finance Limited", "ITGS", "27AABCS5678K1Z2", "Maharashtra (27)", "CGST, SGST", "9, 9", "Net45", "No", "n/a", "Yes", "Yes", "Plot 12, BKC, Mumbai, Maharashtra 400051"],
    [5, "Kaveri Healthcare", "Kaveri Healthcare Private Limited", "IAC", "27AABCK9034L1Z8", "Maharashtra (27)", "CGST, SGST", "9, 9", "Net30", "No", "n/a", "No", "Yes", "8 Pune Station Road, Pune, Maharashtra 411001"],
    [6, "Zenith Exports", "Zenith Exports Private Limited", "ITGS", "24AABCZ7712P1Z3", "Gujarat (24); Rajasthan (08)", "IGST", 18, "Net60", "No", "n/a", "No", "Yes", "21 Ring Road, Surat, Gujarat 395002"],
    [7, "Greenfield Pharma", "Greenfield Pharma Private Limited", "IAC", "36AABCG3345R1Z1", "Telangana (36)", "IGST", 18, "Net30", "Yes", "X", "No", "Yes", "Plot 5, Genome Valley, Hyderabad, Telangana 500078"],
    [8, "Pinnacle Advisory", "Pinnacle Advisory Services LLP", "IAC", "n/a", "Other Territory (97)", "n/a", "n/a", "Net30", "n/a", "n/a", "No", "No", "1 Raffles Quay, Singapore"],
]:
    ws.append(r)
style(ws, [7, 22, 36, 16, 20, 28, 14, 8, 14, 16, 11, 11, 9, 46])

# ---- tab 3: Service Provider (dummy bank details)
ws = wb.create_sheet("Service Provider")
ws.append(["Company Name (short)", "Company Name (legal)", "Address", "GSTIN", "Phone", "Email", "Account Number",
           "IFSC", "Name on Account", "Website", "Bank Name", "Branch", "Swift Code", "Account Type"])
ws.append(["IAC", "Infinity Analytics Consulting", "Mumbai, Maharashtra, India", "27AABCI1234F1Z5", "9930420065",
           "accounts@iac-dummy.example", "000123456789", "HDFC0000123", "INFINITY ANALYTICS CONSULTING (primary holder)", "www.iac-dummy.example", "HDFC Bank Limited", "Fort, Mumbai", "HDFCINBB", "Current Account"])
ws.append(["ITGS", "Infinyty Tech Global Solutions Private Limited", "Chembur E, Mumbai - 400071, Maharashtra",
           "27AABCI9876G1Z4", "9930420066", "accounts@itgs-dummy.example", "000987654321", "HDFC0000013",
           "INFINYTY TECH GLOBAL SOLUTIONS PRIVATE LIMITED", "www.infinyty.tech", "HDFC Bank Limited", "Bandra Kurla Complex, Mumbai", "HDFCINBB", "Current Account"])
style(ws, [22, 46, 40, 20, 16, 34, 20, 16, 48, 24, 22, 28, 14, 16])

# ---- tab 4: Settings
ws = wb.create_sheet("Settings")
ws.append(["Setting", "Value", "Meaning"])
for r in [["GST Rate (INR)", 18, "Used when the client row has no GST % value"],
          ["Withholding line label", "Amount Withheld", "Default text of the withholding line (editable per invoice)"],
          ["Signature label", "Authorized Signature", "Text under the signature box"],
          ["Terms & Conditions 1", "Services are billed as per the agreed scope and commercial terms.", "Optional. Printed on every invoice. Add more rows (2, 3, ...) as needed"],
          ["Terms & Conditions 2", "Late payments may attract interest at 18% per annum.", "Optional"],
          ["Terms & Conditions 3", "Please notify us of any discrepancies within 7 days of receipt.", "Optional"]]:
    ws.append(r)
style(ws, [28, 24, 70])

OUT.parent.mkdir(exist_ok=True)
wb.save(OUT)
shutil.copy(OUT, BASE / "Invoice Generation Inputs (Dummy Data).xlsx")

last = BASE / "Inputs" / "last_invoice_numbers.txt"
if not last.exists():
    last.write_text("# Last invoice number issued for each service provider (one per line, PROVIDER=NUMBER).\n"
                    "# The tool updates this file every time invoices are generated.\n"
                    "# These are EXAMPLE values in each provider's own format: replace with your real last numbers.\n"
                    "IAC=INV-000151\nITGS=ITGS-1008\n", encoding="utf-8")
print("created", OUT)
