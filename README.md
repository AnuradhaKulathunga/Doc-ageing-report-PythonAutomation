
# 📊 Document Ageing Report Automation (Python + Excel)

This project automates the creation of a **Document Ageing Report** from a daily `export.xls` file provided by the client.  
Instead of manually preparing reports in Excel (which takes ~30 minutes daily), this script generates a **Final Report** automatically in just a few seconds.

---

## 🚀 Features
- Reads `export.xls` (downloaded daily).
- Creates a **Summary sheet** with today's date.
- Generates a new sheet for **each account**.
- Adds an **extra "Doc Ageing" column** (days since document date).
- Formats tables with:
  - Bold headers
  - Light blue header fill
  - Autofit columns
  - Number formatting
  - Borders for better readability
- Adds **total rows** for amounts.
- Hides gridlines for a clean look.
- Saves the output as `outputfile/Final Report.xlsx`.

---


---

## 🛠️ Requirements
- Python 3.8+
- Libraries:
  - `pandas`
  - `numpy`
  - `xlwings`

Install them via:
```bash
pip install pandas numpy xlwings 

```
## 📂 Project Structure

```bash
Doc-ageing-report-PythonAutomation/
├─ export.xls # Input file (daily download)
├─ script.py # Main automation script
├─ outputfile/ # Contains Final Report.xlsx (auto-generated)
└─ README.md # Documentation (this file)
```


## ▶️ How to Run


1. Place the daily `export.xls` file in the project folder.

2. In Jupyter Notebook, go to `Cell → Run All` to execute all cells.


3. The final report will be created in:
```bash
outputfile/Final Report.xlsx
```
    

## 📖 Example Output

**Summary Sheet** → Overview of all companies & accounts.

**Account Sheets** → Each account gets its own sheet with transactions, Doc Ageing, totals, and formatting.

## 👨‍💻 Author

**Developed by Anuradha Kulathunga**
