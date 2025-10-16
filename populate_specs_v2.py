#!/usr/bin/env python3
"""
Enhanced Product Specification Auto-Populator for Matrixify Excel Files

This script extracts product specifications from HTML descriptions and other fields,
with improved logic and optional AI-powered extraction.

Version 2.0 - Enhanced extraction with better validation
"""

import pandas as pd
import re
import json
from bs4 import BeautifulSoup
from typing import List, Dict, Optional, Tuple
import sys
import os

class EnhancedSpecExtractor:
    """Extract specifications with improved validation and cleaning"""
    
    def __init__(self, use_ai: bool = False, api_key: Optional[str] = None):
        self.use_ai = use_ai
        self.api_key = api_key
        
        # Improved patterns with better validation
        self.patterns = {
            'Dimensions': [
                (r'(\d+\.?\d*)\s*["\']?\s*[xX×]\s*(\d+\.?\d*)\s*["\']?\s*[xX×]\s*(\d+\.?\d*)\s*["\']?\s*(inches?|in|cm|mm|")?', 
                 lambda m: self._format_dimensions(m)),
                (r'(?:dimensions?|size):?\s*(\d+\.?\d*)\s*[xX×]\s*(\d+\.?\d*)\s*(?:[xX×]\s*(\d+\.?\d*))?',
                 lambda m: self._format_dimensions(m)),
            ],
            'Weight': [
                (r'(\d+\.?\d*)\s*(lbs?|pounds?|kg|kilograms?)\b',
                 lambda m: f"{m.group(1)} {m.group(2)}"),
            ],
            'Capacity': [
                (r'(?:capacity:?\s*)?(\d+\.?\d*)\s*(?:-\s*)?(cup|cups|oz|ounces?|ml|milliliters?|liter|liters?|l|gallon|gallons?|qt|quarts?)\b',
                 lambda m: f"{m.group(1)} {m.group(2)}"),
                (r'(?:up to|upto)\s+(\d+)\s+(cup|cups)\b',
                 lambda m: f"{m.group(1)} {m.group(2)}"),
            ],
            'Power': [
                (r'(\d+)\s*(watt|watts|w)\b(?!\s*max)',
                 lambda m: f"{m.group(1)} {m.group(2)}"),
            ],
            'Voltage': [
                (r'(\d+)\s*(volt|volts|v)\b',
                 lambda m: f"{m.group(1)} {m.group(2)}"),
            ],
            'Material': [
                (r'(?:made (?:of|from)|material:?)\s+([a-z\s]{3,30}?)(?:\.|,|$|\s+(?:with|and|for))',
                 lambda m: self._clean_material(m.group(1))),
                (r'\b(stainless\s+steel|carbon\s+steel|aluminum|plastic|glass|ceramic|wood|silicone|rubber)\b',
                 lambda m: m.group(1).title()),
            ],
            'Color': [
                (r'(?:color|colour):?\s+([a-z\s]{3,20}?)(?:\.|,|$)',
                 lambda m: m.group(1).strip().title()),
            ],
            'Temperature Range': [
                (r'(?:up to|max(?:imum)?)\s+(\d+)\s*(?:degrees?|°)?\s*(f|fahrenheit|c|celsius)?',
                 lambda m: f"{m.group(1)}° {m.group(2).upper() if m.group(2) else 'F'}"),
            ],
            'Settings': [
                (r'(\d+)\s+(?:precision\s+)?(?:speed|heat|temperature)\s+settings?',
                 lambda m: f"{m.group(1)} settings"),
            ],
            'Pieces': [
                (r'(\d+)\s*(?:-\s*)?piece',
                 lambda m: f"{m.group(1)} pieces"),
            ],
        }
    
    def _format_dimensions(self, match) -> str:
        """Format dimension matches consistently"""
        groups = [g for g in match.groups() if g and g.replace('.', '').isdigit()]
        if len(groups) >= 3:
            unit = match.group(4) if len(match.groups()) >= 4 and match.group(4) else "inches"
            if unit == '"':
                unit = "inches"
            return f"{groups[0]} x {groups[1]} x {groups[2]} {unit}"
        elif len(groups) == 2:
            return f"{groups[0]} x {groups[1]}"
        return ""
    
    def _clean_material(self, material: str) -> str:
        """Clean and format material string"""
        material = material.strip().lower()
        # Remove common filler words
        material = re.sub(r'\b(with|and|for|the|a|an)\b.*', '', material).strip()
        if 3 < len(material) < 30:
            return material.title()
        return ""
    
    def _is_valid_value(self, value: str, max_length: int = 100) -> bool:
        """Validate extracted value"""
        if not value or not value.strip():
            return False
        value = value.strip()
        
        # Check length
        if len(value) > max_length:
            return False
        
        # Check for too many special characters
        special_char_ratio = sum(1 for c in value if not c.isalnum() and c != ' ') / len(value)
        if special_char_ratio > 0.3:
            return False
        
        # Check for nonsensical patterns
        if value.count(',') > 5 or value.count('.') > 5:
            return False
        
        return True
    
    def clean_html(self, html_text: str) -> str:
        """Remove HTML tags and clean text"""
        if pd.isna(html_text):
            return ""
        soup = BeautifulSoup(str(html_text), 'html.parser')
        text = soup.get_text(separator=' ', strip=True)
        return ' '.join(text.split())
    
    def extract_specs_from_text(self, text: str, max_specs: int = 6) -> Dict[str, str]:
        """Extract specifications with improved validation"""
        specs = {}
        text_lower = text.lower()
        
        for spec_name, patterns in self.patterns.items():
            if len(specs) >= max_specs:
                break
                
            for pattern, formatter in patterns:
                matches = re.search(pattern, text_lower, re.IGNORECASE)
                if matches:
                    try:
                        value = formatter(matches)
                        if value and self._is_valid_value(value):
                            specs[spec_name] = value
                            break
                    except:
                        continue
        
        return specs
    
    def extract_key_features(self, text: str, max_features: int = 2) -> List[Tuple[str, str]]:
        """Extract concise key features"""
        features = []
        
        # Look for feature lists (with bullets or numbered)
        bullet_pattern = r'[•●○▪▫]\s*([^•●○▪▫\n]{10,80})'
        bullets = re.findall(bullet_pattern, text)
        for bullet in bullets[:max_features]:
            feature = bullet.strip().rstrip('.,')
            if 10 < len(feature) < 80 and self._is_valid_value(feature):
                features.append(('Feature', feature))
        
        if features:
            return features
        
        # Look for 'featuring' or 'includes' phrases
        feature_keywords = [
            (r'featuring:?\s+([^.]{15,80})\.', 'Feature'),
            (r'includes:?\s+([^.]{15,80})\.', 'Included'),
            (r'comes with:?\s+([^.]{15,80})\.', 'Included'),
        ]
        
        for pattern, label in feature_keywords:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches[:max_features - len(features)]:
                feature = match.strip().rstrip(',')
                if 15 < len(feature) < 80 and self._is_valid_value(feature):
                    features.append((label, feature))
        
        return features
    
    def extract_brand(self, title: str, vendor: str = None) -> Optional[str]:
        """Extract brand with better validation"""
        # Prefer vendor field if available
        if pd.notna(vendor):
            brand = str(vendor).strip()
            if 2 < len(brand) < 30:
                return brand
        
        # Extract from title
        if pd.notna(title):
            words = str(title).split()
            if words:
                brand = words[0]
                if brand[0].isupper() and 2 < len(brand) < 30:
                    return brand
        
        return None
    
    def process_row(self, row: pd.Series) -> str:
        """Process a single row with enhanced extraction"""
        all_specs = {}
        
        # 1. Extract structured specs from description
        if pd.notna(row.get('Body HTML')):
            clean_text = self.clean_html(row['Body HTML'])
            text_specs = self.extract_specs_from_text(clean_text, max_specs=5)
            all_specs.update(text_specs)
        
        # 2. Extract from title (for dimensions, capacity, etc.)
        if pd.notna(row.get('Title')):
            clean_title = self.clean_html(row['Title'])
            title_specs = self.extract_specs_from_text(clean_title, max_specs=2)
            for key, value in title_specs.items():
                if key not in all_specs:
                    all_specs[key] = value
        
        # 3. Add brand
        brand = self.extract_brand(row.get('Title'), row.get('Vendor'))
        if brand and 'Brand' not in all_specs:
            all_specs['Brand'] = brand
        
        # 4. Add product type
        if pd.notna(row.get('Type')):
            product_type = str(row['Type']).strip()
            if product_type and 3 < len(product_type) < 50:
                all_specs['Category'] = product_type
        
        # 5. Add material from metafield if available
        if pd.notna(row.get('Metafield: custom.product_material [single_line_text_field]')):
            material = str(row['Metafield: custom.product_material [single_line_text_field]']).strip()
            if 'Material' not in all_specs and 3 < len(material) < 30:
                all_specs['Material'] = material
        
        # 6. Add 1-2 key features only if we have less than 4 specs
        if len(all_specs) < 4 and pd.notna(row.get('Body HTML')):
            clean_text = self.clean_html(row['Body HTML'])
            features = self.extract_key_features(clean_text, max_features=min(2, 5 - len(all_specs)))
            for label, feature in features:
                all_specs[label] = feature
        
        return self.format_specs_for_shopify(all_specs)
    
    def format_specs_for_shopify(self, specs: Dict[str, str]) -> str:
        """Format specifications as JSON array"""
        if not specs:
            return ""
        
        spec_list = [f"{key}: {value}" for key, value in specs.items()]
        return json.dumps(spec_list)


def process_excel_file(
    input_file: str,
    output_file: str = None,
    overwrite_existing: bool = False,
    verbose: bool = True
):
    """
    Process Excel file and populate specifications
    
    Args:
        input_file: Path to input Excel file
        output_file: Path to output file (defaults to input_file_with_specs.xlsx)
        overwrite_existing: If True, overwrite existing specs
        verbose: If True, print detailed progress
    """
    
    if verbose:
        print(f"Reading Excel file: {input_file}")
    
    df = pd.read_excel(input_file, sheet_name=0)
    
    if verbose:
        print(f"Found {len(df)} products\n")
    
    # Initialize extractor
    extractor = EnhancedSpecExtractor()
    
    # Column names
    spec_col = 'Metafield: custom.spec_list [list.single_line_text_field]'
    
    # Track statistics
    stats = {
        'total': len(df),
        'already_populated': 0,
        'newly_populated': 0,
        'skipped': 0,
        'specs_extracted': 0
    }
    
    # Process each row
    if verbose:
        print("Processing products...")
        print("="*70)
    
    for idx, row in df.iterrows():
        # Check if already has specs
        has_existing = pd.notna(row[spec_col]) and str(row[spec_col]).strip()
        
        if has_existing and not overwrite_existing:
            stats['already_populated'] += 1
            continue
        
        # Extract and format specs
        specs_json = extractor.process_row(row)
        
        if specs_json:
            df.at[idx, spec_col] = specs_json
            stats['newly_populated'] += 1
            
            # Count specs
            try:
                specs_list = json.loads(specs_json)
                stats['specs_extracted'] += len(specs_list)
                
                # Print sample for first few
                if verbose and stats['newly_populated'] <= 3:
                    print(f"\n✓ {row.get('Title', 'Unknown')}")
                    for spec in specs_list:
                        print(f"  • {spec}")
            except:
                pass
        else:
            stats['skipped'] += 1
    
    # Determine output file
    if output_file is None:
        base_name = os.path.splitext(input_file)[0]
        output_file = f"{base_name}_with_specs.xlsx"
    
    # Save the updated Excel file
    if verbose:
        print(f"\n{'='*70}")
        print(f"Saving to: {output_file}")
    
    df.to_excel(output_file, index=False, engine='openpyxl')
    
    # Print summary
    if verbose:
        print(f"\n{'='*70}")
        print("PROCESSING COMPLETE")
        print("="*70)
        print(f"Total products: {stats['total']}")
        print(f"Already had specs: {stats['already_populated']}")
        print(f"Newly populated: {stats['newly_populated']}")
        print(f"Skipped (no data): {stats['skipped']}")
        print(f"Total specs extracted: {stats['specs_extracted']}")
        if stats['newly_populated'] > 0:
            print(f"Average specs per product: {stats['specs_extracted']/stats['newly_populated']:.1f}")
        print(f"\nOutput saved to: {output_file}")
        print("="*70)
    
    return output_file


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Auto-populate product specifications from descriptions',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python populate_specs_v2.py test-list.xlsx
  python populate_specs_v2.py test-list.xlsx -o output.xlsx
  python populate_specs_v2.py test-list.xlsx --overwrite
        """
    )
    
    parser.add_argument('input_file', help='Input Excel file path')
    parser.add_argument('-o', '--output', help='Output file path (optional)')
    parser.add_argument('--overwrite', action='store_true',
                       help='Overwrite existing specifications')
    parser.add_argument('-q', '--quiet', action='store_true',
                       help='Suppress progress output')
    
    args = parser.parse_args()
    
    try:
        process_excel_file(
            args.input_file,
            args.output,
            overwrite_existing=args.overwrite,
            verbose=not args.quiet
        )
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)
