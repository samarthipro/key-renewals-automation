Key Renewals Automation – KeyRisk
Python-based Excel automation for insurance renewal processing, net rate comparison, and payroll lookups. The goal is to reduce repetitive spreadsheet work, improve consistency, and make the workflow easier to maintain.
Confidentiality: This workflow may use client, insurance, and payroll data. Do not commit real workbooks, policy documents, client information, credentials, or sensitive logs to GitHub. Follow your organization's data-handling rules.

Overview
The workflow currently covers:
1. Reading underwriting data from the Master worksheet.
2. Transferring selected UW fields into Master Comparision.
3. Copying required fields into Expiring Dump (KeyRisk).
4. Preparing lookup keys from Lotus Merged Payroll in Rates Payroll.xlsx.
5. Performing two separate lookups to populate Expiring Dump output columns.
6. Filling unmatched rates from the downloaded policy PDF by State and Class Code.
7. Saving updated workbooks and reporting match statistics.
Workflow details
1. UW data transfer
Source: Lotus HR Inc Template 26-27 UW V2.xlsx, worksheet Master.
The script finds the following fields by header name:
- Client Id
- Name
- Tax PayerId
- State
- Class Code
- Payroll
These records are written to columns C:H, starting at row 3, in the template workbook's Master Comparision worksheet.
2. Expiring Dump transfer
Columns D:G from Master Comparision are copied to Expiring Dump (KeyRisk), starting at D3.
3. Rates Payroll preparation
Source: Rates Payroll.xlsx, worksheet Lotus Merged Payroll.
- Keep rows where Column E begins with KEYPEO.
- Concatenate Column H and Column I without a separator.
- Write the resulting lookup key to Column P for qualifying rows.
4. Lookup logic
Lookup	Match key	Return value	Destination
First	Expiring Dump Column B matched against Rates Payroll Column P	Rates Payroll Column L	Expiring Dump Column H, starting at H3
Second	Expiring Dump Column F matched against Rates Payroll Column H	Rates Payroll Column N	Expiring Dump Column K, starting at K3


The script uses the first occurrence when duplicate keys exist. Unmatched keys are left blank and included in summary counts.

5. Policy PDF rate fallback
For blank values in Expiring Dump Column H, the script reads State from Column F and Class Code from Column G. It selects the newest PDF in the local `downloads/` folder, extracts the table column headed `Rate Per $100 of Remuneration`, and fills H when the PDF State and Class Code match exactly. Rows not found in the PDF remain blank and are logged.
Technology stack
- Python 3
- Microsoft Excel desktop
- pywin32 for Excel COM automation
- pdfplumber for extracting State, Class Code, and Rate from policy PDFs
- Git and GitHub for version control
- Kiro for development
Requirements
- Windows
- Microsoft Excel desktop installed
- Python 3 and Git
- pywin32
Install the dependency:
python -m pip install -r requirements.txt
requirements.txt includes:
- pywin32>=306
- playwright>=1.47
- python-dotenv>=1.0
- pdfplumber>=0.11
Suggested project structure
key-renewals-automation/
├── scripts/
│   └── key_renewals_combined_with_rates_lookup.py
├── docs/
│   └── workflow.md
├── .gitignore
├── requirements.txt
└── README.md
Keep input workbooks in local folders such as UW/, Expiry/, Quote/, and template/. Do not upload these workbooks unless your organization explicitly approves it.
Configuration
Scripts may contain local absolute paths such as:
C:\Users\<username>\Documents\Key Risk automation
Update the paths for your computer before running. A future improvement is to move paths into a local configuration file excluded from version control.
Running the automation

Payroll preflight check (run first)
Before the main transfer, compare UW `Master` Column R against the Quote workbook's Column J for rows whose Column F contains `Classcode`:
```
python compare_payroll_totals.py
```
If rounded totals differ by no more than $1, the script prints `MATCH` and does not modify either workbook. If they differ by more than $1, it backs up the Quote workbook and creates/replaces a `Payroll Comparison` sheet containing whole-dollar differences grouped by State and Class Code; grouped differences of $1 or less are omitted.

Main workflow
1. Save the approved script in your local scripts/ folder.
2. Confirm that all configured file paths exist.
3. Close the workbooks in Excel.
4. Run from Anaconda Prompt or a Windows terminal, using the exact script filename:
python "scripts\key_renewals_combined_with_rates_lookup.py"
Test with copies of the workbooks first and verify the results before using them in a live renewal workflow.
Validation and troubleshooting
- Confirm worksheet names and header/data start rows.
- Review counts for qualifying payroll records, matches, and unmatched keys.
- Verify duplicate-key behavior and confirm first-match behavior is appropriate.
- Check that formulas remain correct after importing larger datasets.
- If Excel automation fails to launch, confirm Excel opens normally and check the Python/pywin32 installation.
Security and Git hygiene
Before committing, review staged files:
git status
git diff --cached --stat
Make sure .gitignore excludes sensitive files such as *.xlsx, *.xlsm, *.xls, *.csv, .env files, credentials, logs, and backups containing business data.
Never commit passwords, tokens, client records, payroll details, or confidential policy information.
Roadmap
- [ ] Move machine-specific paths into local configuration.
- [ ] Add pre-run validation for files, sheets, and columns.
- [ ] Create test workbooks with synthetic, non-sensitive data.
- [ ] Produce a sanitized report of matched and unmatched keys.
- [ ] Improve formula-row handling and validation.
- [ ] Separate the workflow into step-specific scripts.
- [ ] Add a main.py orchestrator after the individual steps are stable.
- [ ] Add automated tests for lookup and data-validation logic.
Status
This README describes the current workflow and its intended behavior. Treat each step as production-ready only after testing it against approved workbooks.
