from pathlib import Path
from excel_reader import extract_account_credit_records  # or your file name

records = extract_account_credit_records(Path("company_data.xlsx"))
print(f"Found {len(records)} account credit records:")
for r in records[:5]:
    print(r.to_dict())