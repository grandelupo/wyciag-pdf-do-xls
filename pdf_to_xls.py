#!/usr/bin/env python3
"""
Bank Statement PDF to Excel Converter
Converts Polish bank statements to Excel format with columns:
- Data (Date)
- Kontahent / Numer rachunku (Counterparty / Account number)
- Opis / Typ transakcji (Description / Transaction type)
- Kwota (Amount)
"""

import re
import sys
from pathlib import Path
import pdfplumber
import pandas as pd
from decimal import Decimal


def extract_transactions_from_pdf(pdf_path):
    """
    Extract transaction data from a PDF bank statement.
    
    Args:
        pdf_path: Path to the PDF file
        
    Returns:
        List of dictionaries containing transaction data
    """
    transactions = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            text = page.extract_text()
            
            if not text:
                continue
            
            # Split text into lines
            lines = text.split('\n')
            
            # Pattern for transaction line number at start (e.g., "1 01.09.2025")
            # This helps us identify where a new transaction starts
            transaction_start_pattern = r'^(\d+)\s+(\d{2}\.\d{2}\.\d{4})\s+'
            
            # Pattern for amount with PLN
            # Matches: "1 579,00 PLN" or "579,00 PLN" or "18 000,00 PLN"
            # Amount can have spaces as thousand separators (every 3 digits from right)
            # Must be preceded by whitespace or start, followed by " PLN"
            amount_pattern = r'(?:^|\s)(-?\d{1,3}(?:\s\d{3})*,\d{2})\s+PLN'
            
            # Pattern for Polish account number (26 digits, sometimes with spaces)
            account_pattern = r'\b\d{2}\s?\d{4}\s?\d{4}\s?\d{4}\s?\d{4}\s?\d{4}\s?\d{4}\b'
            
            # Pattern to identify balance line (ends with "PLN" followed by number)
            balance_pattern = r'\d+,\d{2}\s+PLN\s+\d+,\d{2}\s+PLN'
            
            i = 0
            while i < len(lines):
                line = lines[i].strip()
                
                # Check if this line starts a new transaction
                match = re.match(transaction_start_pattern, line)
                
                if match:
                    lp = match.group(1)
                    date = match.group(2)
                    
                    # Extract rest of the line after the transaction number and date
                    rest_of_line = line[match.end():].strip()
                    
                    # Initialize transaction data
                    counterparty_name = ""
                    counterparty_address = ""
                    account_number = ""
                    description = ""
                    amount = ""
                    
                    # The first line contains counterparty name and transaction amount at the end
                    # Look for amount on this line (first PLN amount is the transaction amount)
                    amount_matches = list(re.finditer(amount_pattern, rest_of_line))
                    
                    if amount_matches:
                        # Take the FIRST amount (transaction amount, not balance)
                        first_amount = amount_matches[0]
                        amount = first_amount.group(1).replace(' ', '')
                        # Everything before the first amount is part of counterparty name
                        counterparty_name = rest_of_line[:first_amount.start()].strip()
                    else:
                        counterparty_name = rest_of_line
                    
                    # Look ahead for continuation lines (address, account, description)
                    j = i + 1
                    found_account = False
                    lines_to_process = []
                    
                    # Collect lines until we hit another transaction
                    while j < len(lines) and j < i + 5:
                        next_line = lines[j].strip()
                        
                        # Stop if we hit another transaction
                        if re.match(transaction_start_pattern, next_line):
                            break
                        
                        # Stop at page markers
                        if 'Wyciąg nr' in next_line or 'Dokument wygenerowany' in next_line:
                            break
                        
                        if next_line:
                            lines_to_process.append(next_line)
                        
                        j += 1
                    
                    # Now process the collected lines
                    for line_idx, next_line in enumerate(lines_to_process):
                        # Check for account number
                        acc_match = re.search(account_pattern, next_line)
                        if acc_match and not found_account:
                            account_number = acc_match.group(0).replace(' ', '')
                            found_account = True
                            
                            # Text before account number is likely address
                            before_acc = next_line[:acc_match.start()].strip()
                            if before_acc and not counterparty_address:
                                counterparty_address = before_acc
                            
                            # Text after account number is description
                            after_acc = next_line[acc_match.end():].strip()
                            # Remove balance amounts from description
                            # Balance pattern: "XXX XXX,XX PLN" at the end
                            after_acc = re.sub(r'\s*\d[\d\s]*,\d{2}\s+PLN\s*$', '', after_acc)
                            if after_acc:
                                if description:
                                    description += " " + after_acc
                                else:
                                    description = after_acc
                            continue
                        
                        # If we haven't found account yet, might be address continuation
                        if not found_account:
                            if counterparty_address:
                                counterparty_address += " " + next_line
                            else:
                                counterparty_address = next_line
                        else:
                            # After account, it's description
                            # Clean up: remove any balance amounts
                            clean_line = re.sub(r'\s*\d[\d\s]*,\d{2}\s+PLN\s*$', '', next_line)
                            if clean_line:
                                if description:
                                    description += " " + clean_line
                                else:
                                    description = clean_line
                    
                    # Build counterparty field
                    counterparty_parts = []
                    if counterparty_name:
                        counterparty_parts.append(counterparty_name)
                    if account_number:
                        counterparty_parts.append(account_number)
                    
                    counterparty = " / ".join(counterparty_parts) if counterparty_parts else ""
                    
                    if amount:  # Only add if we found an amount
                        transaction = {
                            'Data': date,
                            'Kontahent / Numer rachunku': counterparty,
                            'Opis / Typ transakcji': description,
                            'Kwota': amount
                        }
                        transactions.append(transaction)
                    
                    # Move to the line where we stopped
                    i = j - 1
                
                i += 1
    
    return transactions


def save_to_excel(transactions, output_path):
    """
    Save transactions to an Excel file.
    
    Args:
        transactions: List of transaction dictionaries
        output_path: Path where to save the Excel file
    """
    df = pd.DataFrame(transactions)
    
    # Ensure columns are in the correct order
    columns = ['Data', 'Kontahent / Numer rachunku', 'Opis / Typ transakcji', 'Kwota']
    df = df[columns]
    
    # Save to Excel
    df.to_excel(output_path, index=False, engine='openpyxl')
    print(f"  ✓ Saved {len(transactions)} transactions to {output_path}")


def merge_excel_files(excel_files, output_path):
    """
    Merge multiple Excel files into one combined file.
    
    Args:
        excel_files: List of paths to Excel files to merge
        output_path: Path where to save the combined Excel file
        
    Returns:
        Total number of transactions in the combined file
    """
    all_transactions = []
    
    for excel_file in excel_files:
        try:
            df = pd.read_excel(excel_file, engine='openpyxl')
            all_transactions.append(df)
        except Exception as e:
            print(f"  ⚠ Warning: Could not read {excel_file.name}: {e}")
    
    if not all_transactions:
        return 0
    
    # Combine all dataframes
    combined_df = pd.concat(all_transactions, ignore_index=True)
    
    # Sort by date (convert to datetime for proper sorting)
    try:
        combined_df['Data_Sort'] = pd.to_datetime(combined_df['Data'], format='%d.%m.%Y')
        combined_df = combined_df.sort_values('Data_Sort')
        combined_df = combined_df.drop('Data_Sort', axis=1)
    except Exception:
        # If date parsing fails, keep original order
        pass
    
    # Save combined file
    combined_df.to_excel(output_path, index=False, engine='openpyxl')
    
    return len(combined_df)


def process_single_pdf(pdf_path, output_path=None):
    """
    Process a single PDF file and convert to Excel.
    
    Args:
        pdf_path: Path to the PDF file
        output_path: Optional output path for Excel file
        
    Returns:
        True if successful, False otherwise
    """
    # Determine output path
    if output_path is None:
        output_path = pdf_path.with_suffix('.xlsx')
    
    print(f"Processing: {pdf_path.name}")
    
    try:
        # Extract transactions
        transactions = extract_transactions_from_pdf(pdf_path)
        
        if not transactions:
            print(f"  ⚠ Warning: No transactions found in {pdf_path.name}")
            return False
        
        # Save to Excel
        save_to_excel(transactions, output_path)
        print(f"  ✓ Created: {output_path.name}")
        return True
        
    except Exception as e:
        print(f"  ✗ Error processing {pdf_path.name}: {e}")
        return False


def main():
    """Main function to process PDF to Excel conversion."""
    if len(sys.argv) < 2:
        print("Usage: python pdf_to_xls.py <pdf_file_or_folder> [options]")
        print("\nExamples:")
        print("  Single file:")
        print("    python pdf_to_xls.py statement.pdf")
        print("    python pdf_to_xls.py statement.pdf output.xlsx")
        print("\n  Process all PDFs in a folder:")
        print("    python pdf_to_xls.py /path/to/folder")
        print("\n  Process folder and merge into one file:")
        print("    python pdf_to_xls.py /path/to/folder --merge")
        print("    python pdf_to_xls.py /path/to/folder --merge combined.xlsx")
        sys.exit(1)
    
    # Check for --merge flag
    merge_files = '--merge' in sys.argv
    args = [arg for arg in sys.argv[1:] if arg != '--merge']
    
    if not args:
        print("Error: No input path provided")
        sys.exit(1)
    
    input_path = Path(args[0])
    
    if not input_path.exists():
        print(f"Error: Path '{input_path}' not found")
        sys.exit(1)
    
    # Check if input is a directory or file
    if input_path.is_dir():
        # Process all PDF files in the directory
        pdf_files = sorted(input_path.glob('*.pdf'))
        
        if not pdf_files:
            print(f"No PDF files found in '{input_path}'")
            sys.exit(1)
        
        print(f"Found {len(pdf_files)} PDF file(s) in '{input_path}'")
        print("=" * 50)
        
        successful = 0
        failed = 0
        created_excel_files = []
        
        for pdf_file in pdf_files:
            success = process_single_pdf(pdf_file)
            if success:
                successful += 1
                excel_file = pdf_file.with_suffix('.xlsx')
                created_excel_files.append(excel_file)
            else:
                failed += 1
            print("-" * 50)
        
        print("=" * 50)
        print(f"Summary: {successful} successful, {failed} failed")
        
        # Merge files if requested
        if merge_files and created_excel_files:
            print("=" * 50)
            print("Merging all Excel files into one combined file...")
            
            # Determine combined output path
            if len(args) >= 2:
                combined_output = Path(args[1])
            else:
                combined_output = input_path / 'combined_all_statements.xlsx'
            
            total_transactions = merge_excel_files(created_excel_files, combined_output)
            
            if total_transactions > 0:
                print(f"✓ Combined {len(created_excel_files)} file(s) with {total_transactions} total transactions")
                print(f"✓ Created: {combined_output}")
            else:
                print("✗ Failed to merge files")
        
        if failed > 0:
            sys.exit(1)
            
    else:
        # Process single file
        if not input_path.suffix.lower() == '.pdf':
            print(f"Error: '{input_path}' is not a PDF file")
            sys.exit(1)
        
        # Determine output path
        if len(args) >= 2:
            output_path = Path(args[1])
        else:
            output_path = None
        
        print("=" * 50)
        success = process_single_pdf(input_path, output_path)
        print("=" * 50)
        
        if not success:
            sys.exit(1)


if __name__ == "__main__":
    main()

