# Bank Statement PDF to Excel Converter

This tool converts Polish bank statement PDFs (specifically BNP Paribas format) to Excel files with structured transaction data.

## Supported Format

Currently supports BNP Paribas bank statements with transactions formatted as:
- Transaction date
- Counterparty name and address
- Account number (26 digits)
- Transaction description/reference codes
- Amount in PLN

## Output Format

The Excel file will contain the following columns:
- **Data** - Transaction date (format: DD.MM.YYYY)
- **Kontahent / Numer rachunku** - Counterparty name and account number
- **Opis / Typ transakcji** - Transaction description/reference codes and type (e.g., PRZELEW OTRZYMANY)
- **Kwota** - Transaction amount (format: X XXX,XX with comma as decimal separator)

## Installation

1. Make sure you have Python 3.7 or higher installed
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Process a single file (output will have the same name as input with .xlsx extension):
```bash
python pdf_to_xls.py statement.pdf
```

### Specify custom output file for single file:
```bash
python pdf_to_xls.py statement.pdf output.xlsx
```

### Process all PDF files in a folder:
```bash
python pdf_to_xls.py /path/to/folder
```

This will:
- Find all `.pdf` files in the specified folder
- Convert each one to a corresponding `.xlsx` file in the same folder
- Display a summary of successful and failed conversions

### Process folder and merge all outputs into one combined file:
```bash
python pdf_to_xls.py /path/to/folder --merge
```

This will:
- Process all PDF files in the folder (creating individual `.xlsx` files)
- Automatically merge all the Excel files into one combined file: `combined_all_statements.xlsx`
- Sort all transactions by date in the combined file

You can also specify a custom name for the combined file:
```bash
python pdf_to_xls.py /path/to/folder --merge my_combined_file.xlsx
```

### Example with a single file:
```bash
python pdf_to_xls.py example-statement.pdf
```

This will create: `example-statement.xlsx`

### Example with a folder:
```bash
python pdf_to_xls.py ./statements/
```

This will process all PDF files in the `statements` folder and create corresponding Excel files in the same location.

## Example Output

Given a PDF bank statement, the script will extract transactions like:

```
Data       | Kontahent / Numer rachunku          | Opis / Typ transakcji               | Kwota
-----------|-------------------------------------|-------------------------------------|----------
01.09.2025 | RATAJCZAK MACIEJ / 9310501445100... | PRZELEW OTRZYMANY MP                | 1579,00
01.09.2025 | KRAJEWSKA ANETA / 85116022020000... | PRZELEW OTRZYMANY MP                | 1579,00
01.09.2025 | AGNIESZKA NOCUŃ / 16175015146750... | PRZELEW OTRZYMANY MP                | 18000,00
```

The script successfully extracts all transaction details including dates, counterparty information, descriptions, and amounts.

## How It Works

The script:
1. Opens the PDF file and extracts text from each page
2. Identifies transaction lines using pattern matching for dates and amounts
3. Parses transaction details including dates, amounts, counterparties, and descriptions
4. Exports the data to a formatted Excel file

When using folder mode with `--merge`:
1. Processes each PDF file individually (creating separate `.xlsx` files)
2. Reads all the generated Excel files
3. Combines all transactions into a single dataframe
4. Sorts all transactions chronologically by date
5. Saves the combined result to a single Excel file

## Troubleshooting

If no transactions are found:
- The PDF might be scanned (image-based) rather than text-based. Try using OCR first.
- The format might be different from expected. You may need to adjust the parsing patterns in the script.
- Check if the PDF is encrypted or password-protected.

## Requirements

- Python 3.7+
- pdfplumber - for PDF text extraction
- pandas - for data manipulation
- openpyxl - for Excel file creation

