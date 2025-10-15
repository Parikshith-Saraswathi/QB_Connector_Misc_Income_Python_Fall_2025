# src/banking.py


import pandas as pd
import json
import argparse


# ----------------------------
# STUDENT 2: QB Integration Placeholder
# ----------------------------
def DepositAdd(customer, chart_of_account, amount, memo=None):
    """
    Placeholder for adding a deposit to QB.
    Student 2 will replace this with actual QB API or team function.
    """
    print(f"Deposit added to QB: Customer={customer}, ChartOfAccount={chart_of_account}, Amount={amount}, Memo={memo}")




# ----------------------------
# STUDENT 1: Read Excel Data
# ----------------------------
def read_excel_data(file_path: str):
    df = pd.read_excel(file_path, sheet_name='Sheet1')  # adjust sheet name if needed
    # Keep only relevant columns
    df = df[['Child ID', 'Customer', 'Chart of Account', 'Amount', 'Memo']]
    return df




# ----------------------------
# STUDENT 1: Read QB Data
# ----------------------------
def read_qb_data(file_path=None):
    """
    Since no QB file is provided, return an empty DataFrame.
    Student 1 handles reading or simulating QB data for testing.
    """
    columns = ['Child ID', 'Customer', 'Chart of Account', 'Amount']
    df = pd.DataFrame(columns=columns)
    return df




# ----------------------------
# STUDENT 1: Compare Data
# ----------------------------
def compare_data(excel_df, qb_df):
    report = {
        "same": 0,
        "excel_only": [],
        "qb_only": [],
        "conflicts": []
    }


    # Compare each Excel row with QB
    for _, row in excel_df.iterrows():
        qb_match = qb_df[qb_df['Child ID'] == row['Child ID']]
        if qb_match.empty:
            report["excel_only"].append(row.to_dict())
        else:
            qb_row = qb_match.iloc[0]
            if (row['Amount'] == qb_row['Amount']) and (row['Chart of Account'] == qb_row['Chart of Account']):
                report["same"] += 1
            else:
                report["conflicts"].append({
                    "excel": row.to_dict(),
                    "qb": qb_row.to_dict()
                })


    # Identify QB-only records
    for _, row in qb_df.iterrows():
        if row['Child ID'] not in excel_df['Child ID'].values:
            report["qb_only"].append(row.to_dict())


    return report




# ----------------------------
# STUDENT 1: Write JSON Report
# ----------------------------
def write_report(report, file_name='report.json'):
    with open(file_name, 'w') as f:
        json.dump(report, f, indent=4)




# ----------------------------
# STUDENT 2: Add Excel-Only Deposits to QB
# ----------------------------
def add_to_qb(excel_only_rows):
    for row in excel_only_rows:
        DepositAdd(
            customer=row['Customer'],
            chart_of_account=row['Chart of Account'],
            amount=row['Amount'],
            memo=row.get('Memo', None)
        )




# ----------------------------
# STUDENT 2: CLI / Main Function
# ----------------------------
def main():
    parser = argparse.ArgumentParser(description="Banking - Make Deposits")
    parser.add_argument('--excel', required=True, help='Path to company_data.xlsx')
    parser.add_argument('--qb', help='Path to QB data file (optional)')
    parser.add_argument('--report', default='report.json', help='Output JSON report file')
    args = parser.parse_args()


    # STUDENT 1: Read Data
    excel_df = read_excel_data(args.excel)
    qb_df = read_qb_data(args.qb)


    # STUDENT 1: Compare Data
    report = compare_data(excel_df, qb_df)


    # STUDENT 1: Write JSON Report
    write_report(report, args.report)


    # STUDENT 2: Add Excel-Only Deposits to QB
    add_to_qb(report["excel_only"])


    # STUDENT 2: Print Summary
    print(f"Report written to {args.report}")
    print(f"Same records: {report['same']}")
    print(f"Excel-only deposits added to QB: {len(report['excel_only'])}")
    print(f"Conflicts found: {len(report['conflicts'])}")
    print(f"QB-only records: {len(report['qb_only'])}")




if __name__ == "__main__":
    main()



