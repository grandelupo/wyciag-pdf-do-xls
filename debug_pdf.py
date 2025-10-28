#!/usr/bin/env python3
"""Debug script to see what text is extracted from the PDF"""

import sys
import pdfplumber

if len(sys.argv) < 2:
    print("Usage: python debug_pdf.py <pdf_file>")
    sys.exit(1)

pdf_path = sys.argv[1]

print(f"Opening: {pdf_path}")
print("=" * 80)

with pdfplumber.open(pdf_path) as pdf:
    for page_num, page in enumerate(pdf.pages, 1):
        print(f"\n{'='*80}")
        print(f"PAGE {page_num}")
        print(f"{'='*80}\n")
        
        text = page.extract_text()
        if text:
            print(text)
        else:
            print("(No text found)")
        
        # Also try to extract tables
        tables = page.extract_tables()
        if tables:
            print(f"\n{'-'*80}")
            print(f"TABLES FOUND: {len(tables)}")
            print(f"{'-'*80}\n")
            for i, table in enumerate(tables):
                print(f"Table {i+1}:")
                for row in table:
                    print(row)
                print()

