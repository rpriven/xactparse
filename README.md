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

- **Multi-pattern regex parser** with fallback logic for format variations
- **AGE/LIFE and CONDITION support** - handles depreciation columns
- **Automatic trade categorization** (17+ categories)
- **Multi-line descriptions** - combines continuation lines intelligently
- **Smart content filtering** - skips dimensions, totals, and notes
- **Depreciation warnings** - shows % held back if applicable
- **Budget calculation** - 60% of RCV estimate
- **Excel formatting** - totals, charts, and trade breakdowns
- **Clear contractor-focused display**

## Format Support

### Supported Formats (67% success rate on real-world estimates)

✅ **Text-based PDFs** - Estimates with extractable text layers:
- American Family (AmFam)
- RC Estimates (most formats)
- Most contractor-generated estimates
- Digital Xactimate exports

✅ **Format Variations Handled:**
- Full format with AGE/LIFE and CONDITION columns
- Simple format without age/condition
- Alternative depreciation notation (parentheses or angle brackets)
- Multi-line item descriptions
- Special characters (dashes, quotes, ampersands, parentheses)

### Limited Support (requires OCR - not yet implemented)

⚠️ **Image-based PDFs** - Scanned or image-rendered estimates:
- Some Allstate estimates
- Some Liberty Mutual estimates
- Some State Farm estimates
- Photocopied or faxed estimates

**Note:** OCR support coming soon using pytesseract/OCRmyPDF

## Requirements

```bash
python3 -m pip install --user pdfplumber pandas openpyxl
```

## Installation

```bash
# Clone or download
cd ~/github/xactparse

# Create wrapper script (already done if using from ~/bin)
chmod +x xactparse.py
sudo ln -s $(pwd)/xactparse.py /usr/local/bin/xactparse
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

## Technical Details

### Parser Architecture

The parser uses a **multi-pattern fallback system** to handle format variations:

1. **Pattern 1: Full Format with AGE/LIFE**
   ```
   NUMBER. DESCRIPTION QTY+UNIT UNIT_PRICE TAX O&P RCV AGE/LIFE [yrs] Text COND% (DEPREC) ACV
   Example: 1. Remove charge... 13.49SQ 7.27 0.00 9.80 107.87 9/NA Avg. 0% (0.00) 107.87
   ```

2. **Pattern 2: Simple Format**
   ```
   NUMBER. DESCRIPTION QTY+UNIT UNIT_PRICE TAX O&P RCV (DEPREC) ACV
   Example: 52. R&R Vinyl window... 3.00EA 895.87 195.90 288.36 3,171.87 (951.56) 2,220.31
   ```

3. **Pattern 3: Angle Brackets**
   ```
   NUMBER. DESCRIPTION QTY UNIT UNIT_PRICE TAX O&P RCV <DEPREC> ACV
   Example: 10. Paint door... 2.00 EA 45.00 1.50 5.00 103.00 <10.30> 92.70
   ```

### Content Filtering

The parser intelligently skips:
- Dimension headers and dimension lines
- Total/subtotal/grand total lines
- Section headers (**CONTENTS**, **SUMMARY**, etc.)
- Instruction text (receipts, documentation requirements)
- Non-line-item content

### Multi-line Handling

Descriptions can span multiple lines. The parser:
1. Detects numbered items (e.g., "52. Description...")
2. Combines continuation lines until next numbered item or skip pattern
3. Preserves special characters and formatting
4. Cleans excessive whitespace

## Testing

Test against sample estimates:

```bash
./test-all-samples.sh
```

Current test results on 12 real-world estimates:
- ✅ **8 successful** (67%): AmFam, RC Estimates, WEBER
- ❌ **4 require OCR** (33%): Some Allstate, Liberty Mutual, State Farm

## Privacy Note

This tool processes PDFs locally. No data is uploaded or transmitted.
Keep PDF estimates secure - they contain customer names, addresses, and financial details.
