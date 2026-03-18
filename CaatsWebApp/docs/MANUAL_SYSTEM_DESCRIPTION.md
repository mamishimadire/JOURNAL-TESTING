# CAATS Web App Manual and System Description

## 1. System Description
The CAATS Web App is a guided audit analytics system for journal testing. It connects to SQL Server, ingests GL and TB data, runs configured CAAT procedures, displays exception analytics, and exports working papers.

Primary outcomes:
- Journal risk detection and exception analysis
- Reconciliation support (GL vs TB)
- Standardized Word and Excel working paper outputs

## 2. User Roles
- Auditor/Analyst: runs the workflow and reviews results.
- Manager/Reviewer: reviews outputs and sign-off comments.

## 3. Prerequisites
- Windows machine with .NET SDK installed.
- SQL Server connectivity and access rights.
- ODBC/SQL drivers available on host machine.

## 4. Start the Application
Use the prepared launcher script:
- Script path: `c:\Users\Mamishi.Madire\Desktop\Webapp\Run webapp script.ps1`

Command example:
```powershell
powershell -ExecutionPolicy Bypass -File "c:\Users\Mamishi.Madire\Desktop\Webapp\Run webapp script.ps1"
```

Optional (skip build):
```powershell
powershell -ExecutionPolicy Bypass -File "c:\Users\Mamishi.Madire\Desktop\Webapp\Run webapp script.ps1" -NoBuild
```

What the script does:
1. Stops old `dotnet` processes that lock build output.
2. Builds the project (unless `-NoBuild` is passed).
3. Starts the app on `http://127.0.0.1:5099`.
4. Waits for readiness and opens browser automatically.

## 5. End-to-End Usage Steps
1. Connect to SQL Server.
2. Select client database.
3. Choose GL table and TB table.
4. Preview data and confirm column mapping.
5. Load full data.
6. Save engagement settings.
7. Run CAAT procedures.
8. Review summary and detailed explorer tables.
9. Export working papers.

## 6. CAAT Procedures Covered
- Journal completeness/reconciliation
- Weekend manual journals
- Public holiday manual journals (SA calendar)
- Backdated entries
- Adjustment/correction descriptions
- Journals above performance materiality
- Round amount analysis
- Duplicate entries
- Unbalanced journals
- Low FSLI accounts
- User analysis
- Benford analysis
- Risk score and ranked procedure outputs
- IRI ranked and detail outputs

## 7. Output Files
Default output location:
- `%USERPROFILE%\Desktop\JE_Audit`

Export artifacts:
- Word working paper (`.docx`)
- Excel detailed workbook (`.xlsx`)
- Optional CSV/ZIP sets

Word report includes:
- Structured sections
- Table of contents
- Aligned full-width tables
- Auditor comment blocks

Excel includes:
- Procedure-specific detailed sheets
- Master population sheet
- Day name and holiday name columns in result tables

## 8. Troubleshooting
### App does not open
- Run:
```powershell
powershell -ExecutionPolicy Bypass -File "c:\Users\Mamishi.Madire\Desktop\Webapp\Run webapp script.ps1"
```
- Check launched console for startup exceptions.

### Build fails with file lock (MSB3026/MSB3021)
- Close running app and rerun script.
- The script already kills stale `dotnet` processes.

### Browser opens but page is unreachable
- Confirm URL: `http://127.0.0.1:5099`
- Confirm no firewall/policy restriction.
- Confirm command window shows `Now listening on`.

### SQL tables not visible
- Confirm login rights and selected database.
- Verify connection mode (trusted vs username/password).

## 9. Operational Notes
- `CaatsState` is in-memory and reflects the latest run in the current app process.
- Export quality depends on correct mapping and engagement setup.
- For production multi-user operation, implement persistent state and authentication.

## 10. Support Handover Checklist
- Confirm script execution on user machine.
- Confirm SQL connection and table discovery.
- Confirm one full run and one export.
- Validate Word and Excel outputs against engagement template.
- Capture sign-off on report formatting and procedure coverage.

## 11. Automated Quality Assurance (Added)
Test project path:
- `c:\Users\Mamishi.Madire\Desktop\Webapp\CaatsWebApp.Tests`

Included baseline tests:
- Core CAAT rule flags (weekend, holiday, backdated, adj/correc, above materiality, round patterns, duplicate, unbalanced, low FSLI)
- Reconciliation and Benford output validation
- Export smoke validation (Word/Excel/ZIP generation and key content checks)

Run tests:
```powershell
Set-Location "c:\Users\Mamishi.Madire\Desktop\Webapp"
dotnet test .\CaatsWebApp.Tests\CaatsWebApp.Tests.csproj
```

Latest verification result:
- Total: 4
- Passed: 4
- Failed: 0
