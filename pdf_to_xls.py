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
    print(f"✓ Saved {len(transactions)} transactions to {output_path}")


def main():
    """Main function to process PDF to Excel conversion."""
    if len(sys.argv) < 2:
        print("Usage: python pdf_to_xls.py <pdf_file> [output_file.xlsx]")
        print("\nExample:")
        print("  python pdf_to_xls.py statement.pdf")
        print("  python pdf_to_xls.py statement.pdf output.xlsx")
        sys.exit(1)
    
    pdf_path = Path(sys.argv[1])
    
    if not pdf_path.exists():
        print(f"Error: File '{pdf_path}' not found")
        sys.exit(1)
    
    # Determine output path
    if len(sys.argv) >= 3:
        output_path = Path(sys.argv[2])
    else:
        output_path = pdf_path.with_suffix('.xlsx')
    
    print(f"Processing: {pdf_path}")
    print(f"Output: {output_path}")
    print("-" * 50)
    
    try:
        # Extract transactions
        transactions = extract_transactions_from_pdf(pdf_path)
        
        if not transactions:
            print("⚠ Warning: No transactions found in the PDF")
            print("The PDF might have a different format or no readable text.")
            sys.exit(1)
        
        # Save to Excel
        save_to_excel(transactions, output_path)
        
        print("-" * 50)
        print(f"✓ Successfully converted {pdf_path.name}")
        print(f"✓ Created: {output_path}")
        
    except Exception as e:
        print(f"✗ Error processing PDF: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()

