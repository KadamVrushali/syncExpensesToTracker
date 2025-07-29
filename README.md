# 🧾 Google Sheets: Expenses to Spending Tracker Sync

This Google Apps Script automates syncing categorized monthly expenses from the **"Expenses Incurred" (EO)** sheet to the corresponding columns in the **"Spending Tracker" (ST)** sheet.

## 📌 Use Case

The finance team maintains two sheets:
- **EO (Expenses Incurred):** Where individual entries of expenses are logged for each category.
- **ST (Spending Tracker):** A monthly summary of expenses categorized by type.

This script helps avoid manual copy-pasting of values every month by automatically syncing the total monthly spend per category from EO to ST.

---

## ⚙️ How It Works

1. **Detects Month**  
   Reads the current reporting month (e.g., `Apr-25`) from `A3` in the EO sheet.

2. **Normalizes Month Format**  
   Converts `Apr-25` ➝ `Apr 25` to match column headers in ST.

3. **Finds Matching Column**  
   Scans row 3 of ST to find the column where the month is listed.

4. **Maps Expense Categories**  
   Uses the labels in row 1 of EO and matches them with the row labels in column `B` of ST.

5. **Transfers Values**  
   Transfers corresponding expense values (from row 3 of EO) into the matching month column in ST.

---

## 🛠️ Challenges We Faced & Fixes

| Challenge | Fix |
|----------|-----|
| EO had no consistent row for data (some rows were just headers) | We updated the script to always read row `3` for expense values and row `1` for categories |
| ST had month names in a different format (`Apr 25` vs `Apr-25`) | Implemented a month-normalization logic |
| Expenses were not syncing when values were 0 or formatted as currency | Ensured script converts and reads numeric values including formatted ones |
| BE rows in EO caused misalignment | Script now reads only from the correct row and column structure |
| Month not found error | Added better logging and fallback handling when months aren't matched |

---

## 🧠 Assumptions

- EO:
  - A3 contains the month (`Apr-25`)
  - Row 1 contains categories like “Hardware & Fab”, “Services”, etc.
  - Row 3 contains the expense values
- ST:
  - Row 3 contains month names like “Apr 25”, “May 25”
  - Column B contains category names (rows 4 and below)

---

## 🪄 Setup Instructions

1. Open your Google Sheet.
2. Go to **Extensions > Apps Script**.
3. Create a new script file: `syncExpensesToTracker.gs`
4. Paste the code into this file.
5. Run `syncExpensesToTracker()` manually or set up a trigger.

---

## 🧪 Running the Script

You can either:
- Click ▶️ in the script editor to run it manually
- Set a time-driven trigger for monthly sync

---

## 👩‍💻 Credits

- **Vrushali Kadam** 

---
