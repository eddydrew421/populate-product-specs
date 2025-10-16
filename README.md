# Product Specification Auto-Populator for Matrixify

This Python script automatically extracts product specifications from descriptions, bulletpoints, and metadata fields in your Matrixify Excel files, then populates the `spec_list` and `variant_spec_list` metafields for Google Merchant Center compliance.

## Problem Solved

When uploading products to Google Merchant Center, empty product specifications can cause "Misrepresentation" complaints. This script automates the tedious manual process of going through each SKU to extract and format specifications from vendor-provided descriptions.

## Features

✅ **Intelligent Extraction** - Uses pattern matching and NLP techniques to identify:
- Warranty information (e.g., "2-year warranty")
- Materials (e.g., "stainless steel", "ceramic", "glass")
- Capacity/Volume (e.g., "12 cups", "5.7 L")
- Power specifications (e.g., "1800 watts", "12V")
- Settings/Modes (e.g., "6 precision settings")
- Features (e.g., "BPA Free", "Dishwasher Safe", "LED Display")
- Dimensions (from dimension metafields)

✅ **Multiple Data Sources** - Extracts from:
- Product Title
- Body HTML description
- Long Description (rich text JSON)
- Product Bulletpoints (rich text JSON)
- Material metafield
- Size metafield
- Dimension metafields (width, height, depth)

✅ **Smart Deduplication** - Prevents duplicate specifications with the same key

✅ **Proper Formatting** - Outputs in the correct Matrixify format:
```json
["Key: Value", "Key: Value", ...]
```

✅ **Selective Processing** - Options to skip or overwrite existing specifications

## Requirements

- Python 3.6+
- pandas library
- openpyxl library (for Excel support)

Install dependencies:
```bash
pip install pandas openpyxl --break-system-packages
```

## Usage

### Basic Usage
Process a file, skipping products that already have specifications:
```bash
python populate_specifications.py input.xlsx
```
This creates `input_with_specs.xlsx`

### Specify Output File
```bash
python populate_specifications.py input.xlsx -o output.xlsx
```

### Overwrite Existing Specifications
```bash
python populate_specifications.py input.xlsx --overwrite
```

### Combined Options
```bash
python populate_specifications.py input.xlsx -o output.xlsx --overwrite
```

## Command Line Arguments

| Argument | Description |
|----------|-------------|
| `input_file` | Path to input Excel file (required) |
| `-o`, `--output` | Path to output Excel file (optional) |
| `--overwrite` | Overwrite existing specifications (default: skip) |
| `-h`, `--help` | Show help message |

## How It Works

### 1. Data Collection
The script reads your Matrixify Excel file and collects data from multiple columns:
- Title
- Body HTML (cleans HTML tags)
- Long Description (parses Shopify rich text JSON)
- Product Bulletpoints (parses Shopify rich text JSON)
- Material and Size metafields
- Dimension metafields

### 2. Intelligent Extraction
Uses regex patterns and keyword matching to identify:
- **Warranty patterns**: "2-year warranty", "5 year warranty"
- **Material keywords**: "stainless steel", "ceramic", "glass", etc.
- **Capacity patterns**: "12 cups", "5.7 L", "80g"
- **Power patterns**: "1800 watts", "12V"
- **Settings patterns**: "6 settings", "5 speeds"
- **Feature indicators**: "BPA-free", "dishwasher safe", "LED display"

### 3. Deduplication
Ensures no duplicate specification keys (e.g., won't add "Material" twice)

### 4. Formatting
Converts extracted specifications to JSON array format required by Matrixify:
```json
["Material: Stainless Steel", "Warranty: 2-year warranty", "BPA Free: Yes"]
```

### 5. Output
Populates both:
- `Metafield: custom.spec_list [list.single_line_text_field]`
- `Variant Metafield: custom.variant_spec_list [list.single_line_text_field]`

## Example Output

### Input Data:
```
Title: Salton Stainless Steel Smart Coffee Grinder
Long Description: "Crafted with premium steel, the Salton electric grinder offers 
6 precision grind settings... With a 2-year warranty and the ability to grind up 
to 12 cups at a time..."
```

### Extracted Specifications:
```json
[
  "Material: Stainless Steel",
  "Capacity: 12 cup",
  "Warranty: 2-year warranty",
  "Feature: Quiet Operation"
]
```

## Customization

### Adding New Specification Patterns
Edit the `spec_patterns` dictionary in the `SpecificationExtractor` class:

```python
self.spec_patterns = {
    'warranty': r'(\d+[\s-]?(?:year|yr|month|day)[\s-]?warranty)',
    'your_pattern': r'(your regex pattern)',
    # Add more patterns here
}
```

### Adding New Feature Keywords
Edit the `feature_indicators` list in `extract_features_from_text()`:

```python
feature_indicators = [
    ('removable', 'Removable'),
    ('dishwasher safe', 'Dishwasher Safe'),
    ('your_keyword', 'Your Label'),
    # Add more feature indicators here
]
```

### Changing Data Sources
Modify the `source_columns` dictionary in `process_excel_file()`:

```python
source_columns = {
    'title': 'Title',
    'your_field': 'Your Column Name',
    # Add more source columns here
}
```

## Limitations & Considerations

### Current Limitations
1. **Pattern-based extraction** - May miss specifications in unusual formats
2. **English language** - Optimized for English descriptions
3. **Quality depends on input** - Better input descriptions = better extraction

### Recommendations
1. **Review output** - Always verify extracted specifications are accurate
2. **Enhance vendor data** - Request more structured data from vendors when possible
3. **Iterate patterns** - Add more patterns based on your product categories
4. **Test on samples** - Run on a small subset first to validate results

## Troubleshooting

### No specifications extracted
- Check if descriptions exist in source fields
- Review description format - script works best with descriptive text
- Add custom patterns for your specific product category

### Incorrect specifications
- Review the extracted text in console output
- Adjust regex patterns for your data format
- Consider adding exclusion rules

### Pandas warnings
- The script uses `dtype=str` to avoid type warnings
- Warnings are usually harmless but can be ignored

## Advanced Usage

### Integration with Workflow
```bash
#!/bin/bash
# Example workflow script

# 1. Export from Shopify via Matrixify
# 2. Process specifications
python populate_specifications.py export.xlsx -o processed.xlsx --overwrite

# 3. Import back to Shopify via Matrixify
# (Upload processed.xlsx to Matrixify)
```

### Batch Processing
```python
import glob
from populate_specifications import process_excel_file

# Process all Excel files in a directory
for file in glob.glob("exports/*.xlsx"):
    output = file.replace("exports/", "processed/")
    process_excel_file(file, output, overwrite=False)
```

## Support & Contributions

This script was created to automate specification extraction for Google Merchant Center compliance. Feel free to:
- Modify patterns for your product categories
- Add new extraction rules
- Enhance the feature detection logic
- Share improvements with your team

## Version History

**v1.0** (Current)
- Initial release
- Pattern-based extraction
- Shopify rich text JSON parsing
- Multi-source data collection
- Command line interface with options

## License

This script is provided as-is for your ecommerce operations. Modify and use as needed for your business.
