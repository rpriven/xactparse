# xactparse

Extract and total Xactimate estimate line items from PDF.

## Purpose

Construction contractors use this tool to quickly answer:
- **How much is my initial check?** (ACV - Actual Cash Value)
- **What's the total job value?** (RCV - Replacement Cost Value)
- **How much is being held back?** (Depreciation)

## Usage

```bash
xactparse estimate.pdf output.xlsx
```

## Output

### Console Summary
```
============================================================
CONTRACTOR SUMMARY
============================================================

💰 INITIAL CHECK (ACV):        $21,893.11
💵 TOTAL JOB VALUE (RCV):      $21,893.11
⏳ HELD BACK (Depreciation):   $0.00
   ✅ No depreciation - Full replacement cost coverage

📊 BUDGET (60% of RCV):        $13,135.87
============================================================
```

### Excel File
- **Totals sheet**: Trade-by-trade breakdown with RCV, ACV, Depreciation
- **Master sheet**: All line items with full details
- **Trade sheets**: Individual sheets per trade category
- **Pie chart**: Visual breakdown of costs by trade

## Features

- Automatic trade categorization (17 categories)
- Depreciation warnings (shows % held back if applicable)
- Budget calculation (60% of RCV estimate)
- Excel formatting with totals and charts
- Clear contractor-focused display

## Requirements

```bash
python3 -m pip install --user pdfplumber pandas openpyxl
```

## Installation

```bash
# Clone or download
cd ~/github/xactparse

# Create wrapper script (already done if using from ~/bin)
chmod +x 57-xactparse.py
sudo ln -s $(pwd)/57-xactparse.py /usr/local/bin/xactparse
```

## Trade Categories

Automatically categorizes line items into:
- Baseboards, Trim, Casing
- Cabinets
- Carpet
- Cleaning
- Content Manipulation
- Doors
- Drywall
- Electrical
- Floor Protection
- HVAC
- Insulation
- Labor Minimums
- Painting
- Plumbing, Toilets, Sinks
- Showers, Tubs, Tile
- Tile Flooring
- Vinyl Flooring
- Laminate
- Mitigation
- Other (unmatched items)

## Related Tools

- **xactdiff**: Compare two estimates to find missing line items

## Privacy Note

This tool processes PDFs locally. No data is uploaded or transmitted.
Keep PDF estimates secure - they contain customer names, addresses, and financial details.
