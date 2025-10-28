import pandas as pd
import sys

if len(sys.argv) < 2:
    print("Usage: python check_output.py <excel_file>")
    sys.exit(1)

df = pd.read_excel(sys.argv[1])

print("All transactions:")
print("=" * 100)
for idx, row in df.iterrows():
    print(f"{idx+1:2d}. {row['Data']} | {row['Kwota']:>15s} | {row['Kontahent / Numer rachunku'][:40]}")

print("=" * 100)
print(f"\nTotal: {len(df)} transactions")
print(f"\nColumns: {df.columns.tolist()}")

