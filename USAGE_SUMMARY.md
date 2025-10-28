# Quick Start Guide

## What's Been Created

✅ **pdf_to_xls.py** - Main conversion script  
✅ **requirements.txt** - Python dependencies  
✅ **README.md** - Full documentation  
✅ **example-statement.xlsx** - Sample output from the example PDF  

## Quick Start

1. **Install dependencies** (one-time setup):
```bash
pip install -r requirements.txt
```

2. **Run the script**:
```bash
python pdf_to_xls.py your-statement.pdf
```

This will create `your-statement.xlsx` with columns:
- **Data** - Transaction date
- **Kontahent / Numer rachunku** - Counterparty and account number  
- **Opis / Typ transakcji** - Description and transaction type
- **Kwota** - Amount in PLN

## Example

The provided `example-statement.pdf` has been successfully converted to `example-statement.xlsx`:
- ✅ 17 transactions extracted
- ✅ All amounts correctly parsed (1579,00 PLN, 18000,00 PLN, etc.)
- ✅ Counterparty names and account numbers captured
- ✅ Transaction descriptions and reference codes preserved

## Tested Format

Currently optimized for **BNP Paribas** Polish bank statements with the format:
```
Lp. Data Kontrahent ... Amount PLN
    Address / Account number ... Description ... Balance PLN
```

## Notes

- The script extracts transaction amounts (not balance amounts)
- Account numbers are included with counterparty information
- Reference codes (CEN..., MBKOZ..., etc.) are preserved in descriptions
- Amounts use Polish formatting (space for thousands, comma for decimals)

For more details, see README.md

