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
misc_income_cli.exe --workbook <path_to_excel> --bank_account <path_to_json> --output <path_to_output>
```

Example:

```
misc_income_cli.exe --workbook C:\Users\SaraswathiP\repo\project\QB_Connector_Misc_Income_Python_Fall_2025\company_data.xlsx --bank_account C:\Users\SaraswathiP\repo\project\QB_Connector_Misc_Income_Python_Fall_2025\src\input_settings.json --output C:\Users\SaraswathiP\repo\project\Project_Result\report.json
```

### Arguments
- `--workbook`: Path to the Excel workbook containing the other income worksheet.
- `--bank_account`: Path to a JSON file specifying the bank account (see below for format).
- `--output`: Path to write the output report (JSON).

## Sample input_settings.json

Create a JSON file like this (e.g., `src/input_settings.json`):

```json
{
	"bank_account": "Chase"
}
```

The value for `bank_account` should match the name of the account in QuickBooks where deposits will be added.
