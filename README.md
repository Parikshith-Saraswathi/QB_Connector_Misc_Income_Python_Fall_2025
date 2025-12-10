# Misc Income CLI Template
## Parikshith Saraswathi, Eduardo Araujo

## Setup Project
Once you forked and cloned the repo, run:
```bash
poetry install
```
to install dependencies.
Then write code in the src/ folder.

## Quality Check
To setup pre-commit hook (you only need to do this once):
```bash
poetry run pre-commit install
```
To manually run pre-commit checks:
```bash
poetry run pre-commit run --all-file
```
To manually run ruff check and auto fix:
```bash
poetry run ruff check --fix
```

## Build

To build the CLI as a standalone Windows executable (.exe), run:

```bash
poetry run pyinstaller --onefile --name misc_income_cli --hidden-import win32timezone --hidden-import win32com.client build_exe.py
```

This will create `misc_income_cli.exe` in the `dist/` folder.

## Usage

Run the CLI with the following arguments:

```
cd dist

.\misc_income_cli.exe --workbook <path_to_excel> --bank_account <bank_account> --output <path_to_output>
```

### Arguments
- `--workbook`: Path to the Excel workbook containing the other income worksheet.
- `--bank_account`: **The bank account to use** (you must specify exactly one). This can be specified in two ways:
  - **Direct bank name**: A string name of the account in QuickBooks (e.g., `Chase`, `Wells Fargo`)
  - **JSON file path**: Path to a JSON file specifying the bank account (e.g., `src/input_settings.json`)
- `--output`: Path to write the output report (JSON).

### Examples

**Using a direct bank account name:**
```powershell
.\misc_income_cli.exe --workbook C:\path\to\company_data.xlsx --bank_account Chase --output C:\path\to\report.json
```

**Using a JSON settings file:**
```powershell
.\misc_income_cli.exe --workbook C:\path\to\company_data.xlsx --bank_account C:\path\to\input_settings.json --output C:\path\to\report.json
```

## Deduplication Behavior

The CLI implements automatic duplicate detection to prevent the same records from being added multiple times to QuickBooks. When you run the CLI:

**First Run:**
- All records from Excel that don't exist in QuickBooks are added
- A report is generated listing all added records

**Subsequent Runs:**
- The CLI compares Excel records with QuickBooks using a composite key: `(amount, chart_of_account, record_id)`
- Records that already exist in QuickBooks with the same amount, chart of account, and record ID are skipped
- Only new or modified records are added
- The console output and report will indicate how many duplicates were skipped

This ensures that running the CLI multiple times with the same Excel file and bank account will not create duplicate entries in QuickBooks.

## Bank Account Requirements

**Important:** Your QuickBooks file must contain **exactly one bank account** with the name you specify. The CLI will add all misc income records to this single account.

## JSON Settings File Format

If you choose to use a JSON settings file instead of a direct bank account name, create a file like this (e.g., `src/input_settings.json`):

```json
{
	"bank_account": "Chase"
}
```

The value for `bank_account` should match **the exact name** of the bank account in QuickBooks where deposits will be added. Then pass the path to this file using the `--bank_account` argument.
