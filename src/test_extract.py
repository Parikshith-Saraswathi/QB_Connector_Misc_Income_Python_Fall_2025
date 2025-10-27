from pathlib import Path
from excel_reader import extract_deposits

def main():
    excel_file = Path("./company_data.xlsx")

    deposits = extract_deposits(excel_file)

    print(f"Found {len(deposits)} records:")
    for d in deposits[:10]:  # print first 10 for brevity
        print(d.to_dict())

if __name__ == "__main__":
    main()
