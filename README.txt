INVOICE TOOL  (everything is in this one folder)

START      Double-click "Start Invoice Tool.bat". The page opens in your browser and loads the invoices straight away.
STOP       Close the black window.

INPUT FILES (Inputs folder)
  Invoice Generation Inputs.xlsx   One Excel with these tabs:
      Invoices To Generate  Client Name, Currency, Amount, Item Description, Notes to put in Invoice,
                            Include in Invoice Name, HSN/SAC, PO Number
      Client Master         ... plus Service Provider (IAC or ITGS) for each client
      Service Provider      legal name, address, GSTIN, phone, email, account number, IFSC, name on account
      Settings              GST rate, withholding label, signature label
  last_invoice_numbers.txt         One line per provider, e.g.  IAC=INV-000151  and  ITGS=ITGS-1008
                                   Put the LAST number you really issued. The tool counts up from it, keeps the
                                   same format, and updates the file after every run.
  logo.png                         The one logo used on every invoice
  signature.png                    Optional digital signature image

The Excel file is read again every time you click "Load invoices" / "Reload" and every time you click Generate.
You can edit the Excel while the tool is running (Save it, then click Reload).

EACH RUN
  1. Open the tool. The tab "Invoices To Generate" is shown with checks (red = fix first, amber = warning).
  2. Adjust place / terms of payment / date / withholding on any row if needed, then click Generate all invoices.
  3. Review each invoice (PDF on the left, details on the right): Correct or Incorrect. "Edit this invoice" fixes the
     withholding line, terms, date or place and rebuilds that PDF with the same number.
  4. Finish: PDFs and Invoices-<Month>-<timestamp>.zip are saved in the Outputs folder.

FILE NAMES   ITGS-1009 ShortClientName CustomString.pdf   and   INV-000152 ShortClientName CustomString.pdf
             (CustomString = the "Include in Invoice Name" column)

ON A NEW COMPUTER (needed once)
  1. Install Python 3 from https://www.python.org/downloads/  (tick "Add python.exe to PATH").
  2. Have Microsoft Edge or Google Chrome installed (it makes the PDFs and shows the page).
  3. Double-click "Start Invoice Tool.bat". On the very first start it installs two small helper
     packages (openpyxl, Pillow) and needs the internet for that. After that it works without it.
  Windows: use the .bat file.   Mac / Linux: open a terminal in this folder and run:  python3 app.py
