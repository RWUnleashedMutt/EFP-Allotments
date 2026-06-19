# EFP Numbers Generator — Complete Tutorial

## Table of Contents

1. [Overview](#overview)
2. [What This Script Does](#what-this-script-does)
3. [Setup & Prerequisites](#setup--prerequisites)
4. [Configuration](#configuration)
5. [Input File Format](#input-file-format)
6. [Output Format](#output-format)
7. [How It Works: Step-by-Step](#how-it-works-step-by-step)
8. [Code Breakdown](#code-breakdown)
9. [Running the Script](#running-the-script)
10. [Customization Guide](#customization-guide)
11. [Troubleshooting](#troubleshooting)
12. [Example Scenarios](#example-scenarios)

---

## Overview

The **EFP Numbers Generator** is a Python script that transforms Square's category-level sales data into a budget allocation workbook. It takes raw sales figures by location and brand, calculates percentages, and automatically divides allocated budgets between "Coastal" and "Local" store groups.

**Purpose:** Help Southeast Pet management visualize brand performance and allocate vendor budgets proportionally across store clusters.

**Key Features:**

- Reads Square category-sales CSV exports
- Category Sales select display by channel
- Creates one Excel sheet per brand
- Calculates sales percentages and budget splits
- Color-codes Coastal vs. Local store groups
- Formats numbers with currency and percentage formatting
- Professional styling (fonts, fills, alignment)

---

## What This Script Does

### Problem It Solves

You have a total budget allocated to each brand (e.g., $1,000 for Nature's Logic). You need to:

1. See how much each store actually spent on that brand
2. Understand each store's share of total sales
3. Automatically calculate how much budget goes to Coastal vs. Local stores

**Manually doing this for 11 stores across 5+ brands is tedious and error-prone.**

### Solution

The script automates this by:

1. Reading a Square CSV export with sales by location and category
2. For each brand:
   - Calculating sales totals per store
   - Computing percentage of total brand sales
   - Allocating the brand's budget to Coastal and Local groups
3. Producing a professional Excel workbook with one sheet per brand

### Input

**File:** Square category-sales CSV export (e.g., `category-sales-2026-04-25-2026-05-25_2_.csv`)

**Columns expected:**

- `Channel` – Store name (e.g., "Lake Boone", "City Market: DTR")
- `Category` – Brand name (e.g., "Natures Logic", "Open Farm")
- `Gross Sales` – Sales amount (may include $ and commas)

### Output

**File:** `EFP_Numbers_All_Brands.xlsx`

**Contents:** One sheet per brand, formatted like:

| Location        | Brand                              | Total          | % of Total  |
| --------------- | ---------------------------------- | -------------- | ----------- |
| LB              | Natures Logic                      | $2,500.00      | 15.50%      |
| CM              | Natures Logic                      | $1,200.00      | 7.45%       |
| ...             | ...                                | ...            | ...         |
| **Grand Total** |                                    | **$16,129.03** | **100.00%** |
|                 |                                    |                |             |
| **Coastal**     | MF, SP, LF                         | $4,500.00      | 27.90%      |
|                 | Coastal allotment of $1,000 budget | $279.03        | 27.90%      |
|                 |                                    |                |             |
| **Local**       | CM, CVM, DTD, PP, SS, LB, SH, CC   | $11,629.03     | 72.10%      |
|                 | Local allotment of $1,000 budget   | $720.97        | 72.10%      |

---

## Setup & Prerequisites

### 1. Install Python & Dependencies

Ensure Python 3.7+ is installed. Install required packages:

```bash
pip install pandas openpyxl
```

Or with a virtual environment:

```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install pandas openpyxl
```

**Packages:**

- `pandas` – Read CSV and manipulate data
- `openpyxl` – Create and style Excel workbooks

### 2. Obtain a Square Export

1. Log in to Square Dashboard
2. Navigate to **Sales** → **Reports**
3. Select **Category Sales** (or similar)
4. Set date range (e.g., April 25 – May 25, 2026)
5. Export as CSV
6. Save the file in the same directory as the script

**Example filename:** `category-sales-2026-04-25-2026-05-25_2_.csv`

### 3. Verify File Placement

```
project_root/
├── generate_efp_numbers.py        (the script)
├── category-sales-2026-04-25-2026-05-25_2_.csv  (Square export)
└── (EFP_Numbers_All_Brands.xlsx will be created here)
```

---

## Configuration

All configuration is in the **CONFIG** section at the top of the script. Edit this to match your brand list, store layout, and budgets.

### BRANDS Dictionary

Maps brand names to allocated budgets:

```python
BRANDS = {
    "Natures Logic":   1000,
    "Open Farm":       2000,
    "Small Batch":     1100,
    "Vital Essentials": 2500,
    "Fromm":           1200,
}
```

**Rules:**

- **Key:** Brand name (must match exactly how it appears in the Square CSV)
- **Value:** Total budget allocated to this brand

**To add a brand:**

1. Identify the exact name from the Square CSV
2. Add a line: `"Brand Name": 5000,` (where 5000 is the budget)

**To remove a brand:**

1. Delete the line

**To change a budget:**

1. Update the value: `"Fromm": 1500,` (changed from 1200)

### CHANNEL_TO_STORE Dictionary

Maps Square's channel names (store full names) to Southeast Pet's store codes:

```python
CHANNEL_TO_STORE = {
    "City Market: DTR":          "CM",
    "Crabtree Valley Mall":      "CVM",
    "Downtown Durham":           "DTD",
    "Parkway Plaza":             "PP",
    "The Streets at Southpoint": "SS",
    "Southport - Tidewater":     "SP",
    "Front Street":              "MF",
    "Lake Boone":                "LB",
    "Stonehenge Market":         "SH",
    "Landfall Shopping Center":  "LF",
    "Crescent Commons":          "CC",
}
```

**Purpose:** Translates Square's verbose store names to short codes for display

**Rules:**

- **Key:** Channel name from Square CSV (must match exactly)
- **Value:** 2-letter store code

**If a channel is missing:**

- The script skips it and logs a warning
- Budget calculations will be incomplete

**To add a channel:**

1. Get the exact name from Square CSV
2. Get the store code (e.g., "WH", "DTD")
3. Add a line: `"Exact Square Channel Name": "XX",`

### COASTAL_STORES & LOCAL_STORES Sets

Partition stores into two geographic groups:

```python
COASTAL_STORES = {"MF", "SP", "LF"}
LOCAL_STORES = {"CM", "CVM", "DTD", "PP", "SS", "LB", "SH", "CC"}
```

**Purpose:**

- Coastal stores are grouped together (beach locations)
- Local stores are grouped together (Raleigh area)
- Budget is split between groups based on their % of total sales

**Rules:**

- Use the 2-letter store **codes** (not full names)
- Each store should be in exactly one set
- Together, they should cover all stores in `CHANNEL_TO_STORE`

**To reassign a store:**

1. Remove its code from one set
2. Add it to the other
3. Example: Move `"WH"` from Local to Coastal:
   ```python
   COASTAL_STORES = {"MF", "SP", "LF", "WH"}
   LOCAL_STORES = {"CM", "CVM", "DTD", "PP", "SS", "LB", "SH", "CC"}
   ```

### CSV_PATH & OUTPUT_PATH

File paths:

```python
CSV_PATH = "category-sales-2026-04-25-2026-05-25_2_.csv"
OUTPUT_PATH = "EFP_Numbers_All_Brands.xlsx"
```

**CSV_PATH:** Path to the Square export CSV (can be relative or absolute)
**OUTPUT_PATH:** Where the Excel workbook will be saved

---

## Input File Format

### Expected CSV Structure

The Square export CSV should contain at least these columns:

| Channel          | Category      | Gross Sales |
| ---------------- | ------------- | ----------- |
| Lake Boone       | Natures Logic | $2,500.00   |
| City Market: DTR | Natures Logic | $1,200.00   |
| Lake Boone       | Open Farm     | $1,800.00   |
| ...              | ...           | ...         |

### Column Details

**Channel**

- Store's full name as it appears in Square
- Example: "City Market: DTR", "Crabtree Valley Mall"
- Must match a key in `CHANNEL_TO_STORE` (or it will be ignored)

**Category**

- Brand name
- Example: "Natures Logic", "Open Farm"
- Must match a key in `BRANDS` (or it will be ignored)

**Gross Sales**

- Sales amount for that store + brand combination
- Format: Can include $ sign and commas (e.g., "$2,500.00")
- The script automatically cleans this to a numeric value

### Example CSV Content

```csv
Channel,Category,Gross Sales
Lake Boone,Natures Logic,$2,500.00
City Market: DTR,Natures Logic,$1,200.00
Crabtree Valley Mall,Natures Logic,$900.00
Lake Boone,Open Farm,$1,800.00
City Market: DTR,Open Farm,$2,100.00
...
```

### Notes on Data Quality

- **Missing data:** If a store-brand combination is missing from the CSV, it's treated as $0.00
- **Spelling:** Column names are case-sensitive; "Category" not "category"
- **Extra columns:** Additional columns are ignored (safe to export)
- **Multiple rows per store-brand:** If there are duplicates, they will be summed

---

## Output Format

### Workbook Structure

The script creates an Excel workbook with **one sheet per brand** (defined in `BRANDS`).

### Sheet Layout

Each sheet follows this structure:

```
Row 1:  [Header]        Location  Brand  Total          % of Total
Row 2:  [Data]          LB        Brand  $2,500.00      15.50%
Row 3:  [Data]          CM        Brand  $1,200.00      7.45%
...     [Data]          ...       ...    ...            ...
Row N:  [Grand Total]                   $16,129.03     100.00%
Row N+1:[Blank]
Row N+2:[Coastal Sum]   Coastal   MF, SP, LF  $4,500.00  27.90%
Row N+3:[Coastal Budget]           Coastal allotment... $279.03   27.90%
Row N+4:[Blank]
Row N+5:[Local Sum]     Local     CM, CVM...  $11,629.03  72.10%
Row N+6:[Local Budget]             Local allotment...  $720.97   72.10%
```

### Column Breakdown

| Column         | Content                   | Format               |
| -------------- | ------------------------- | -------------------- |
| A (Location)   | Store code or group name  | Text                 |
| B (Brand)      | Brand name                | Text                 |
| C (Total)      | Sales amount              | Currency ($X,XXX.00) |
| D (% of Total) | Percentage of grand total | Percentage (X.XX%)   |

### Styling

**Header Row (Row 1):**

- Dark blue background (#4F81BD)
- White bold text
- Centered alignment

**Data Rows:**

- No fill (white background)
- Regular black text
- Left-aligned text, right-aligned numbers

**Grand Total Row:**

- No fill
- Bold text
- Currency and percentage formatting

**Coastal Group (Rows N+2, N+3):**

- Light green fill (#E2EFDA)
- Bold italic text
- Budget row in darker gold text

**Local Group (Rows N+5, N+6):**

- Light blue fill (#DDEBF7)
- Bold italic text
- Budget row in darker gold text

**Column Widths:**

- Column A (Location): 12 characters
- Column B (Brand): 32 characters
- Column C (Total): 14 characters
- Column D (% of Total): 14 characters

---

## How It Works: Step-by-Step

### High-Level Flow

```
Load CSV
    ↓
For each brand in BRANDS:
    ├── Filter CSV for this brand
    ├── Calculate sales per store
    ├── Compute percentages
    ├── Split budget: Coastal vs. Local
    └── Create Excel sheet with results
    ↓
Save workbook
```

### Detailed Walkthrough

#### Step 1: Read CSV

```python
df = pd.read_csv(CSV_PATH)
```

- Loads the entire CSV into a pandas DataFrame
- All columns are read (extras are ignored later)

#### Step 2: Clean Gross Sales Column

```python
df["Gross Sales"] = (
    df["Gross Sales"].astype(str)
    .str.replace(r"[\$,]", "", regex=True)
    .astype(float)
)
```

**What it does:**

1. Convert `Gross Sales` to string (in case it's already numeric)
2. Remove $ and commas using regex: `[\$,]` matches $ or ,
3. Convert result back to float

**Example:**

- Input: `"$2,500.00"`
- After string replace: `"2500.00"`
- After float conversion: `2500.0`

#### Step 3: For Each Brand

```python
for brand, budget in BRANDS.items():
    brand_df = df[df["Category"] == brand]
```

- Loop through each brand in the `BRANDS` config
- Filter the CSV to only rows where `Category == brand`

#### Step 4: Calculate Per-Store Sales

```python
rows = []
for channel, code in CHANNEL_TO_STORE.items():
    match = brand_df[brand_df["Channel"] == channel]
    total = match["Gross Sales"].sum() if not match.empty else 0.0
    rows.append({"Location": code, "Total": total})
summary = pd.DataFrame(rows)
```

**For each store:**

1. Filter `brand_df` to rows where `Channel == channel`
2. Sum the `Gross Sales` (handles multiple rows per store)
3. If no rows found, use 0.0
4. Create a DataFrame with store codes and totals

**Result:**

```
Location  Total
LB        2500.0
CM        1200.0
CVM       900.0
...
```

#### Step 5: Calculate Percentages

```python
grand_total = summary["Total"].sum()
summary["pct"] = summary["Total"] / grand_total if grand_total else 0.0
```

- Sum all store totals to get brand's grand total
- Divide each store's total by grand total
- Result: `pct` column with values like 0.155 (15.5%)

**Safeguard:** If grand_total is 0 (no sales), pct defaults to 0.0

#### Step 6: Group by Coastal vs. Local

```python
coastal_total = summary[summary["Location"].isin(COASTAL_STORES)]["Total"].sum()
local_total = summary[summary["Location"].isin(LOCAL_STORES)]["Total"].sum()
coastal_pct = coastal_total / grand_total if grand_total else 0.0
local_pct = local_total / grand_total if grand_total else 0.0
```

- Filter summary to Coastal stores, sum their totals
- Filter summary to Local stores, sum their totals
- Calculate each group's percentage of grand total

**Example:**

- Coastal Total: $4,500 / $16,129.03 = 27.90%
- Local Total: $11,629.03 / $16,129.03 = 72.10%

#### Step 7: Split Budget

```python
coastal_budget = budget * coastal_pct
local_budget = budget * local_pct
```

- Multiply the brand's total budget by each group's percentage
- Coastal gets: $1,000 × 27.90% = $279.03
- Local gets: $1,000 × 72.10% = $720.97

#### Step 8: Create Excel Sheet

```python
ws = wb.create_sheet(title=brand[:31])
```

- Create a new sheet named after the brand
- `[:31]` limits sheet name to 31 chars (Excel limit)

#### Step 9: Write Headers & Data

```python
for ri, row in enumerate(summary.itertuples(index=False), 2):
    ws.cell(ri, 1, row.Location)
    ws.cell(ri, 2, brand)
    ws.cell(ri, 3, row.Total).number_format = DLR_FORMAT
    ws.cell(ri, 4, row.pct).number_format = PCT_FORMAT
```

- Row 1: Headers (handled separately)
- Rows 2+: Data from summary DataFrame
- Apply number formatting: DLR_FORMAT for currency, PCT_FORMAT for percentage

#### Step 10: Add Summary Rows

```python
# Grand total row
tr = len(summary) + 2
gt = ws.cell(tr, 3, grand_total)

# Coastal and Local summaries
cr = tr + 2
styled_row(ws, cr, COASTAL_FILL, GROUP_FONT, "Coastal", "MF, SP, LF", coastal_total, coastal_pct)
```

- Add a blank row separator
- Add Coastal summary row (with styling)
- Add Coastal budget allocation row
- Repeat for Local
- Styling includes color fills and fonts

#### Step 11: Save Workbook

```python
wb.save(OUTPUT_PATH)
```

- Write the workbook to disk
- File is ready for download/sharing

---

## Code Breakdown

### Section 0: Imports

```python
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
```

**pandas:** Data reading and manipulation
**openpyxl.Workbook:** Create Excel files from scratch
**openpyxl.styles:** Fonts, fills (colors), alignment
**openpyxl.utils.get_column_letter:** Convert column index to letter (1 → A, 2 → B, etc.)

### Section 1: CONFIG

```python
CSV_PATH = "category-sales-2026-04-25-2026-05-25_2_.csv"
OUTPUT_PATH = "EFP_Numbers_All_Brands.xlsx"

BRANDS = { ... }
CHANNEL_TO_STORE = { ... }
COASTAL_STORES = { ... }
LOCAL_STORES = { ... }
```

**All user-editable settings in one place.** Edit these to customize behavior.

### Section 2: STYLES

```python
HEADER_FILL = PatternFill("solid", fgColor="4F81BD")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FONT = Font(bold=True)
# ... more style definitions
```

**Predefined Excel styles used throughout:**

- Fills: Background colors (hex codes)
- Fonts: Bold, italic, color
- Formats: Currency (#,##0.00), percentage (0.00%)

**To change colors:**

1. Find the hex code online (e.g., "light green" = #E2EFDA)
2. Update the `PatternFill` or `Font` definition
3. Rerun the script

### Section 3: Helpers

```python
def styled_row(ws, row, fill, font, col1, col2, total, pct):
    for ci in range(1, 5):
        ws.cell(row, ci).fill = fill
    ws.cell(row, 1, col1).font = font
    # ... more cell assignments
```

**Purpose:** Reduce code duplication when writing styled rows

**Parameters:**

- `ws`: Worksheet object
- `row`: Row number
- `fill`: Background color (PatternFill object)
- `font`: Font style (Font object)
- `col1`: Value for column A (Location / Group name)
- `col2`: Value for column B (Brand / Description)
- `total`: Value for column C (Sales / Budget)
- `pct`: Value for column D (Percentage)

**What it does:**

1. Apply fill color to all 4 cells in the row
2. Apply font to column A and B
3. Set column C value and format as currency
4. Set column D value and format as percentage

**Example usage:**

```python
styled_row(ws, 5, COASTAL_FILL, GROUP_FONT, "Coastal", "MF, SP, LF", 4500, 0.279)
```

### Section 4: main() Function

The orchestrator that ties everything together.

**Main steps (see "How It Works" above for details):**

1. Read CSV
2. Clean Gross Sales column
3. Create empty workbook
4. For each brand:
   a. Filter data
   b. Calculate per-store sales
   c. Compute percentages
   d. Group by Coastal/Local
   e. Split budget
   f. Create sheet
   g. Write data
   h. Write summaries
5. Save workbook

**Print statements:**

- Provides user feedback ("Reading...", "✓ Brand Name")
- Final message: "Saved → filename"

### Section 5: Entry Point

```python
if __name__ == "__main__":
    main()
```

- Allows the script to be run directly: `python generate_efp_numbers.py`
- Also allows importing as a module without auto-executing `main()`

---

## Running the Script

### Prerequisites Checklist

- [ ] Python 3.7+ installed
- [ ] `pandas` and `openpyxl` installed (`pip install pandas openpyxl`)
- [ ] Square CSV export saved in script directory
- [ ] `BRANDS`, `CHANNEL_TO_STORE`, store mappings configured
- [ ] CSV column names are exactly: "Channel", "Category", "Gross Sales"

### Step 1: Verify Configuration

Open `generate_efp_numbers.py` and confirm:

```python
CSV_PATH = "your-square-export.csv"  # ← Correct filename?
OUTPUT_PATH = "EFP_Numbers_All_Brands.xlsx"

BRANDS = {
    "Natures Logic": 1000,
    # ... match your Square export?
}

CHANNEL_TO_STORE = {
    # ... match your store names in the CSV?
}
```

### Step 2: Run the Script

**From terminal:**

```bash
python generate_efp_numbers.py
```

**From IDE (VS Code, PyCharm, etc.):**

- Open the file
- Press Run / Play button

### Step 3: Monitor Output

**Console output example:**

```
Reading category-sales-2026-04-25-2026-05-25_2_.csv ...
  ✓  Natures Logic       budget $1,000  →  Coastal $279.03  |  Local $720.97
  ✓  Open Farm          budget $2,000  →  Coastal $587.23  |  Local $1,412.77
  ✓  Small Batch        budget $1,100  →  Coastal $308.92  |  Local $791.08
  ✓  Vital Essentials   budget $2,500  →  Coastal $698.11  |  Local $1,801.89
  ✓  Fromm              budget $1,200  →  Coastal $335.74  |  Local $864.26

Saved → EFP_Numbers_All_Brands.xlsx
```

**Each line shows:**

- Brand name (padded to 20 chars)
- Total budget
- Coastal allocation
- Local allocation

### Step 4: Verify Output File

1. Check current directory for `EFP_Numbers_All_Brands.xlsx`
2. Open it in Excel
3. Verify:
   - One sheet per brand ✓
   - Header row is blue with white text ✓
   - Data rows have proper formatting ✓
   - Numbers are formatted correctly ($ and %) ✓
   - Coastal and Local summaries are present ✓
   - Grand totals match expected budget splits ✓

---

## Customization Guide

### Scenario 1: Add a New Brand

**You want to include "Wellness Core" with a $1,500 budget.**

**Steps:**

1. Open `generate_efp_numbers.py`
2. Find the `BRANDS` dictionary
3. Add: `"Wellness Core": 1500,`
4. Save
5. Rerun the script

**Before:**

```python
BRANDS = {
    "Natures Logic":   1000,
    "Open Farm":       2000,
    "Small Batch":     1100,
    "Vital Essentials": 2500,
    "Fromm":           1200,
}
```

**After:**

```python
BRANDS = {
    "Natures Logic":   1000,
    "Open Farm":       2000,
    "Small Batch":     1100,
    "Vital Essentials": 2500,
    "Fromm":           1200,
    "Wellness Core":   1500,
}
```

**Important:** Brand name must match exactly as it appears in the Square CSV export.

### Scenario 2: Remove a Brand

**You no longer need "Small Batch" in the output.**

**Steps:**

1. Open `generate_efp_numbers.py`
2. Find the `BRANDS` dictionary
3. Delete the line: `"Small Batch": 1100,`
4. Save
5. Rerun the script

**Result:** "Small Batch" sheet won't be created. (The CSV can still have this data; it just won't be processed.)

### Scenario 3: Add a New Store

**Southeast Pet opens "Eastgate Shopping" with code "EG", a Local store.**

**Steps:**

1. Get the exact channel name from the Square CSV export (e.g., "Eastgate Shopping")
2. Open `generate_efp_numbers.py`
3. Add to `CHANNEL_TO_STORE`: `"Eastgate Shopping": "EG",`
4. Add "EG" to `LOCAL_STORES`: `LOCAL_STORES = {"CM", "CVM", "DTD", "PP", "SS", "LB", "SH", "CC", "EG"}`
5. Save and rerun

**Before:**

```python
CHANNEL_TO_STORE = {
    # ... existing stores ...
    "Crescent Commons": "CC",
}

LOCAL_STORES = {"CM", "CVM", "DTD", "PP", "SS", "LB", "SH", "CC"}
```

**After:**

```python
CHANNEL_TO_STORE = {
    # ... existing stores ...
    "Crescent Commons": "CC",
    "Eastgate Shopping": "EG",
}

LOCAL_STORES = {"CM", "CVM", "DTD", "PP", "SS", "LB", "SH", "CC", "EG"}
```

### Scenario 4: Reassign a Store to Coastal

**"Front Street" (MF) is being reclassified as Local instead of Coastal.**

**Before:**

```python
COASTAL_STORES = {"MF", "SP", "LF"}
LOCAL_STORES = {"CM", "CVM", "DTD", "PP", "SS", "LB", "SH", "CC"}
```

**After:**

```python
COASTAL_STORES = {"SP", "LF"}
LOCAL_STORES = {"CM", "CVM", "DTD", "PP", "SS", "LB", "SH", "CC", "MF"}
```

**Effect:** All future runs will allocate MF's sales percentage to the Local group budget instead of Coastal.

### Scenario 5: Change a Budget

**The board increased Open Farm's budget from $2,000 to $2,500.**

**Before:**

```python
BRANDS = {
    "Open Farm": 2000,
}
```

**After:**

```python
BRANDS = {
    "Open Farm": 2500,
}
```

**Rerun.** All percentages stay the same; but the Coastal and Local allocations will be higher.

### Scenario 6: Change the Color Scheme

**You want Coastal to be light orange instead of light green.**

**Step 1: Find the hex code for light orange** (e.g., #FFE6CC)

**Step 2: Update the COASTAL_FILL:**

```python
COASTAL_FILL = PatternFill("solid", fgColor="FFE6CC")  # Changed from E2EFDA
```

**Rerun.** Coastal summary rows will now have an orange background.

**Common colors:**

- Light red: `#F4CCCC`
- Light orange: `#FFE6CC`
- Light yellow: `#FFF2CC`
- Light green: `#E2EFDA`
- Light blue: `#DDEBF7`
- Light purple: `#E4DFEC`

### Scenario 7: Use a Different CSV File

**You have a new export from a different date range.**

**Before:**

```python
CSV_PATH = "category-sales-2026-04-25-2026-05-25_2_.csv"
```

**After:**

```python
CSV_PATH = "category-sales-2026-06-01-2026-06-30_2_.csv"
```

**Rerun.** The script reads the new file and generates fresh allocations.

---

## Troubleshooting

### Issue: "FileNotFoundError: [Errno 2] No such file or directory: 'category-sales-...csv'"

**Cause:** CSV file not found at the specified path

**Solution:**

1. Verify the CSV file exists in the script directory
2. Check the exact filename (case-sensitive on Linux/Mac)
3. Update `CSV_PATH` to match the actual filename
4. Rerun

### Issue: "KeyError: 'Gross Sales'" or "KeyError: 'Channel'"

**Cause:** CSV columns don't match expected names

**Solution:**

1. Open the CSV in Excel or a text editor
2. Verify column names are exactly: "Channel", "Category", "Gross Sales"
3. If different, edit the code:
   ```python
   df["Channel"]  # change to your column name
   df["Category"]  # change to your column name
   df["Gross Sales"]  # change to your column name
   ```
4. Rerun

### Issue: No sheets appear in the output Excel file

**Cause:** None of the brands in `BRANDS` matched any categories in the CSV

**Solution:**

1. Open the CSV in Excel
2. Check the exact brand names in the "Category" column
3. Update `BRANDS` to match these exact names
4. Rerun

**Example:**

- CSV has: "Natures Logic" (no apostrophe)
- Config had: "Natures Logic" (different formatting)
- Change to match the CSV exactly

### Issue: Budget allocations don't add up to the total budget

**Cause:** This is normal due to rounding and percentage calculations

**Note:**

- Coastal: $1,000 × 27.90% = $279.00 (rounded)
- Local: $1,000 × 72.10% = $721.00 (rounded)
- Total: $1,000.00 ✓

Rounding differences of $0.01–$0.10 are expected and acceptable.

### Issue: A store appears in CSV but not in the output

**Cause:** Store channel name doesn't match `CHANNEL_TO_STORE`

**Solution:**

1. Find the exact channel name in the CSV
2. Add it to `CHANNEL_TO_STORE`:
   ```python
   "Exact Channel Name From CSV": "XX",
   ```
3. Add the store code to `COASTAL_STORES` or `LOCAL_STORES`
4. Rerun

### Issue: ModuleNotFoundError: No module named 'pandas'

**Cause:** Required packages not installed

**Solution:**

```bash
pip install pandas openpyxl
```

### Issue: Output file is empty or corrupted

**Cause:** Workbook creation failed partway through

**Solution:**

1. Delete the output file: `EFP_Numbers_All_Brands.xlsx`
2. Run the script again
3. If the issue persists, check for errors in console output

### Issue: Percentages don't sum to 100% on a sheet

**Cause:** Rounding errors across many stores

**Note:** This is normal. Percentages are rounded to 2 decimal places; slight variance is acceptable.

**To verify:** Add up all the percentages manually—they should be very close to 100% (within ±0.05%).

---

## Example Scenarios

### Scenario A: Monthly Budget Allocation

**You run this script at the start of each month to allocate vendor budgets based on prior month's sales.**

**Workflow:**

1. Export Square category sales for the prior month (e.g., May 2026)
2. Save as `category-sales-2026-05-01-2026-05-31.csv`
3. Update `CSV_PATH` in the script
4. Rerun `generate_efp_numbers.py`
5. New output: `EFP_Numbers_All_Brands.xlsx`
6. Share with the team
7. Orders are placed based on the budget allocations

### Scenario B: Quarterly Budget Review

**Scenario:** End of Q2 2026, you need to present performance to the board.

**Steps:**

1. Export Q2 sales data from Square (April–June)
2. Also export Q1 and Q2 combined for YTD analysis
3. Run the script twice:
   - `generate_efp_numbers.py` (uses most recent config)
   - Generates Q2 breakdowns
4. Create two workbooks:
   - `EFP_Q2_Only.xlsx`
   - `EFP_YTD_2026.xlsx`
5. Present to board with insights:
   - Coastal vs. Local performance
   - Budget allocation effectiveness

**Modification:** Save outputs to different filenames:

```python
OUTPUT_PATH = "EFP_Q2_Only.xlsx"
# ... run script ...

# Then change to:
OUTPUT_PATH = "EFP_YTD_2026.xlsx"
# ... run again with YTD CSV ...
```

### Scenario C: Multi-Year Comparison

**Scenario:** Compare brand performance across 2025 and 2026.

**Steps:**

1. Export 2025 data: `category-sales-2025-full-year.csv`
2. Update `CSV_PATH` and `OUTPUT_PATH`:
   ```python
   CSV_PATH = "category-sales-2025-full-year.csv"
   OUTPUT_PATH = "EFP_2025_Full_Year.xlsx"
   ```
3. Run the script → generates `EFP_2025_Full_Year.xlsx`
4. Repeat with 2026 data:
   ```python
   CSV_PATH = "category-sales-2026-full-year.csv"
   OUTPUT_PATH = "EFP_2026_Full_Year.xlsx"
   ```
5. Run the script → generates `EFP_2026_Full_Year.xlsx`
6. Open both files side-by-side to compare:
   - Which brands grew?
   - Did Coastal outperform Local in 2026 vs. 2025?
   - Are budgets aligned with performance trends?

### Scenario D: Seasonal Analysis

**Scenario:** Create separate reports for summer (high season) and winter (low season).

**Steps:**

1. Export summer sales (June–August): `category-sales-summer-2026.csv`
2. Generate report:
   ```python
   CSV_PATH = "category-sales-summer-2026.csv"
   OUTPUT_PATH = "EFP_Summer_2026.xlsx"
   ```
3. Export winter sales (Dec–Feb): `category-sales-winter-2025-26.csv`
4. Generate report:
   ```python
   CSV_PATH = "category-sales-winter-2025-26.csv"
   OUTPUT_PATH = "EFP_Winter_2025-26.xlsx"
   ```
5. Compare:
   - Does Coastal outperform in summer?
   - Do certain brands spike seasonally?
   - Adjust budget allocations accordingly

---

## Summary

The **EFP Numbers Generator** automates a routine but tedious task: translating Square sales data into budget allocation spreadsheets. By running this script monthly, you:

1. **Save time** – No manual spreadsheet building
2. **Reduce errors** – Automated calculations are accurate
3. **Enable data-driven decisions** – See where sales are happening and allocate budget proportionally
4. **Maintain consistency** – Same format every time
5. **Facilitate comparisons** – Easy to track trends across months/years

The script is highly configurable. Adjusting brands, stores, budgets, and colors takes just a few edits to the CONFIG section.

For questions or issues, check the **Troubleshooting** section or review the printed console output for clues.
