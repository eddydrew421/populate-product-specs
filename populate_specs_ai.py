#!/usr/bin/env python3
"""
AI-Powered Product Specification Auto-Populator

Uses Claude API to intelligently extract and generate product specifications
even from minimal descriptions. More accurate than pattern matching.

Requires: ANTHROPIC_API_KEY environment variable
"""

import pandas as pd
import json
from typing import List, Dict, Optional
import sys
import os
from bs4 import BeautifulSoup
import time

try:
    import anthropic
except ImportError:
    print("Error: anthropic package not installed")
    print("Install with: pip install anthropic --break-system-packages")
    sys.exit(1)


class AISpecExtractor:
    """Extract specifications using Claude API"""
    
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.environ.get('ANTHROPIC_API_KEY')
        if not self.api_key:
            raise ValueError(
                "API key required. Set ANTHROPIC_API_KEY environment variable "
                "or pass api_key parameter"
            )
        
        self.client = anthropic.Anthropic(api_key=self.api_key)
        self.model = "claude-3-5-haiku-20241022"  # Fast and cost-effective
        
        # System prompt for spec extraction
        self.system_prompt = """You are a product specification expert. Extract and format product specifications from product information.

Rules:
1. Output ONLY a JSON array of specifications
2. Each spec must be in format "Label: Value"
3. Extract ONLY factual specifications that are explicitly stated or strongly implied
4. Common spec types: Dimensions, Weight, Capacity, Material, Color, Power, Brand, Features
5. Keep specs concise (under 100 characters each)
6. Include 3-6 specs maximum
7. Do not make up specifications that aren't supported by the information
8. If brand is in the title, always include it

Example output format:
["Brand: Salton", "Material: Stainless Steel", "Capacity: 12 cups", "Power: 1800 watts"]

Return ONLY the JSON array, no other text."""
    
    def clean_html(self, html_text: str) -> str:
        """Remove HTML tags"""
        if pd.isna(html_text):
            return ""
        soup = BeautifulSoup(str(html_text), 'html.parser')
        return soup.get_text(separator=' ', strip=True)
    
    def build_product_context(self, row: pd.Series) -> str:
        """Build context string from product data"""
        context_parts = []
        
        if pd.notna(row.get('Title')):
            context_parts.append(f"Product: {row['Title']}")
        
        if pd.notna(row.get('Body HTML')):
            description = self.clean_html(row['Body HTML'])
            context_parts.append(f"Description: {description}")
        
        if pd.notna(row.get('Type')):
            context_parts.append(f"Category: {row['Type']}")
        
        if pd.notna(row.get('Vendor')):
            context_parts.append(f"Vendor: {row['Vendor']}")
        
        if pd.notna(row.get('Metafield: custom.product_material [single_line_text_field]')):
            material = row['Metafield: custom.product_material [single_line_text_field]']
            context_parts.append(f"Material: {material}")
        
        return "\n".join(context_parts)
    
    def extract_specs_with_ai(self, product_context: str) -> Optional[str]:
        """Use Claude to extract specifications"""
        try:
            message = self.client.messages.create(
                model=self.model,
                max_tokens=300,
                system=self.system_prompt,
                messages=[
                    {
                        "role": "user",
                        "content": product_context
                    }
                ]
            )
            
            # Extract the JSON array from response
            response_text = message.content[0].text.strip()
            
            # Parse and validate JSON
            specs = json.loads(response_text)
            
            if isinstance(specs, list) and len(specs) > 0:
                # Validate format
                for spec in specs:
                    if not isinstance(spec, str) or ':' not in spec:
                        return None
                return json.dumps(specs)
            
            return None
            
        except json.JSONDecodeError:
            print(f"Warning: Invalid JSON response from API")
            return None
        except Exception as e:
            print(f"Warning: API error: {e}")
            return None
    
    def process_row(self, row: pd.Series) -> str:
        """Process a single row with AI"""
        product_context = self.build_product_context(row)
        
        if not product_context.strip():
            return ""
        
        return self.extract_specs_with_ai(product_context) or ""


def process_excel_file(
    input_file: str,
    output_file: str = None,
    api_key: Optional[str] = None,
    overwrite_existing: bool = False,
    verbose: bool = True,
    batch_delay: float = 0.5
):
    """
    Process Excel file using AI
    
    Args:
        input_file: Path to input Excel file
        output_file: Path to output file
        api_key: Anthropic API key (or use ANTHROPIC_API_KEY env var)
        overwrite_existing: If True, overwrite existing specs
        verbose: If True, print detailed progress
        batch_delay: Delay between API calls (seconds) to avoid rate limits
    """
    
    if verbose:
        print(f"Reading Excel file: {input_file}")
    
    df = pd.read_excel(input_file, sheet_name=0)
    
    if verbose:
        print(f"Found {len(df)} products")
        print(f"Using AI model: claude-3-5-haiku-20241022")
        print(f"Estimated cost: ${len(df) * 0.0003:.2f} (approximate)\n")
    
    # Initialize AI extractor
    try:
        extractor = AISpecExtractor(api_key=api_key)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)
    
    # Column name
    spec_col = 'Metafield: custom.spec_list [list.single_line_text_field]'
    
    # Track statistics
    stats = {
        'total': len(df),
        'already_populated': 0,
        'newly_populated': 0,
        'failed': 0,
        'specs_extracted': 0,
        'api_calls': 0
    }
    
    # Process each row
    if verbose:
        print("Processing products with AI...")
        print("="*70)
    
    for idx, row in df.iterrows():
        # Check if already has specs
        has_existing = pd.notna(row[spec_col]) and str(row[spec_col]).strip()
        
        if has_existing and not overwrite_existing:
            stats['already_populated'] += 1
            continue
        
        # Extract with AI
        specs_json = extractor.process_row(row)
        stats['api_calls'] += 1
        
        if specs_json:
            df.at[idx, spec_col] = specs_json
            stats['newly_populated'] += 1
            
            # Count specs
            try:
                specs_list = json.loads(specs_json)
                stats['specs_extracted'] += len(specs_list)
                
                # Print samples
                if verbose and stats['newly_populated'] <= 3:
                    print(f"\n✓ {row.get('Title', 'Unknown')}")
                    for spec in specs_list:
                        print(f"  • {spec}")
            except:
                pass
        else:
            stats['failed'] += 1
        
        # Progress indicator
        if verbose and (idx + 1) % 10 == 0:
            print(f"\nProcessed {idx + 1}/{len(df)} products...")
        
        # Rate limiting
        if batch_delay > 0:
            time.sleep(batch_delay)
    
    # Determine output file
    if output_file is None:
        base_name = os.path.splitext(input_file)[0]
        output_file = f"{base_name}_with_ai_specs.xlsx"
    
    # Save
    if verbose:
        print(f"\n{'='*70}")
        print(f"Saving to: {output_file}")
    
    df.to_excel(output_file, index=False, engine='openpyxl')
    
    # Summary
    if verbose:
        print(f"\n{'='*70}")
        print("PROCESSING COMPLETE")
        print("="*70)
        print(f"Total products: {stats['total']}")
        print(f"Already had specs: {stats['already_populated']}")
        print(f"Newly populated: {stats['newly_populated']}")
        print(f"Failed: {stats['failed']}")
        print(f"Total specs extracted: {stats['specs_extracted']}")
        if stats['newly_populated'] > 0:
            print(f"Average specs per product: {stats['specs_extracted']/stats['newly_populated']:.1f}")
        print(f"API calls made: {stats['api_calls']}")
        print(f"Estimated cost: ${stats['api_calls'] * 0.0003:.2f}")
        print(f"\nOutput saved to: {output_file}")
        print("="*70)
    
    return output_file


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(
        description='AI-powered product specification extraction',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Environment Variables:
  ANTHROPIC_API_KEY    Your Anthropic API key (required)

Examples:
  export ANTHROPIC_API_KEY="sk-ant-..."
  python populate_specs_ai.py test-list.xlsx
  python populate_specs_ai.py test-list.xlsx -o output.xlsx
  python populate_specs_ai.py test-list.xlsx --overwrite --delay 1.0
        """
    )
    
    parser.add_argument('input_file', help='Input Excel file path')
    parser.add_argument('-o', '--output', help='Output file path (optional)')
    parser.add_argument('--api-key', help='Anthropic API key (or use env var)')
    parser.add_argument('--overwrite', action='store_true',
                       help='Overwrite existing specifications')
    parser.add_argument('-q', '--quiet', action='store_true',
                       help='Suppress progress output')
    parser.add_argument('--delay', type=float, default=0.5,
                       help='Delay between API calls (seconds, default: 0.5)')
    
    args = parser.parse_args()
    
    try:
        process_excel_file(
            args.input_file,
            args.output,
            api_key=args.api_key,
            overwrite_existing=args.overwrite,
            verbose=not args.quiet,
            batch_delay=args.delay
        )
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)
