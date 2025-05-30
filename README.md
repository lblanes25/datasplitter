# Audit Leader Workbook Splitter

This Python script splits a consolidated Excel workbook of audit data into separate Excel workbooks for each unique audit leader. Each resulting workbook contains only the relevant data for that leader, preserving formatting and applying tab colors based on QA results (e.g., red for DNC findings).

---

## 📦 Requirements

- Python 3.8+
- pandas
- numpy
- openpyxl

Install dependencies with:

```bash
pip install -r requirements.txt
