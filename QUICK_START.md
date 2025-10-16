# Quick Start Guide

## Setup (First Time Only)

1. **Install Python dependencies:**
```bash
pip install pandas openpyxl --break-system-packages
```

2. **Download the script:**
   - `populate_specifications.py` - Main script
   - `README.md` - Full documentation

## Basic Workflow

### Step 1: Export from Shopify
- Use Matrixify to export your products to Excel
- Save as `products.xlsx` (or any name)

### Step 2: Run the Script
```bash
# Process the file
python populate_specifications.py products.xlsx
```

This will create `products_with_specs.xlsx` with populated specifications.

### Step 3: Review Output
Open `products_with_specs.xlsx` and check:
- Column: `Metafield: custom.spec_list [list.single_line_text_field]`
- Look for JSON arrays like: `["Material: Stainless Steel", "Warranty: 2-year"]`

### Step 4: Import to Shopify
- Use Matrixify to import `products_with_specs.xlsx` back to Shopify
- Check a few products on your store to verify specifications appear

## Common Commands

```bash
# Basic - skip products with existing specs
python populate_specifications.py input.xlsx

# Overwrite all specifications
python populate_specifications.py input.xlsx --overwrite

# Custom output filename
python populate_specifications.py input.xlsx -o my_products.xlsx

# Both custom output and overwrite
python populate_specifications.py input.xlsx -o output.xlsx --overwrite
```

## What Gets Extracted?

The script looks for these specifications:
- ✅ Warranty (e.g., "2-year warranty")
- ✅ Material (e.g., "stainless steel", "ceramic")
- ✅ Capacity (e.g., "12 cups", "5.7 L")
- ✅ Power (e.g., "1800 watts", "12V")
- ✅ Settings (e.g., "6 precision settings")
- ✅ Features (e.g., "BPA Free", "Dishwasher Safe")
- ✅ Dimensions (from dimension metafields)

## Expected Output

For each product, you'll see:
```
✓ product-handle
  Extracted 4 specifications:
    - Material: Stainless Steel
    - Capacity: 12 cup
    - Warranty: 2-year warranty
    - Feature: Quiet Operation
```

## Troubleshooting

**Problem:** No specifications extracted
- **Solution:** Check if product has description/bulletpoints in Matrixify export

**Problem:** Too few specifications
- **Solution:** Vendor descriptions may be incomplete. Consider enriching them manually or requesting better data from vendors

**Problem:** Wrong specifications
- **Solution:** Review the pattern matching logic and customize for your products

## Tips for Best Results

1. **Better Vendor Data = Better Results**
   - Request detailed descriptions from vendors
   - Ask for structured specifications when possible

2. **Review Before Import**
   - Always check a sample of products before importing to Shopify
   - Verify specifications are accurate and relevant

3. **Customize for Your Products**
   - Add patterns specific to your product category
   - See README.md for customization instructions

4. **Iterative Approach**
   - Process a small batch first
   - Review and adjust
   - Process full catalog

## Need Help?

See `README.md` for:
- Full documentation
- Customization guide
- Advanced usage examples
- Troubleshooting details
