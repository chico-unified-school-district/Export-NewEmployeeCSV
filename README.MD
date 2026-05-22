# Export-NewEmployeeCSV

Monitors an intermediate SQL table for new employee rows and writes a CSV file into a protected file share.

## Dependencies

- PowerShell 5.1 or later
- `dbatools`
- `CommonScriptFunctions`

## What It Does

The script reads rows from the configured intermediate SQL table where status is not complete, converts each record into the CSV export format defined in `lib/attributes.txt`, and writes the file to a mapped `exports:` drive.

It can be run as a one-time dry run with `-WhatIf` or scheduled to loop until the configured stop time.

For each row, the script:

1. Looks up the employee by SSN in the employee database table.
2. Builds the export file name and destination paths.
3. Exports CSV data when status is `new` or `update`, or if a prior export file is missing.
4. Removes CSV files for rows that are old or already exist in the employee table.
5. Updates the intermediate table with the new status and import path.

## Parameters

- `-IntSQLServer`: SQL instance hosting the intermediate table
- `-IntSQLDatabase`: database containing the intermediate table
- `-IntSQLTable`: intermediate source table name
- `-IntSQLCredential`: credential for the intermediate SQL connection
- `-EmpSQLServer`: SQL instance hosting the employee table
- `-EmpSQLDatabase`: database containing the employee table
- `-EmpSQLTable`: employee table name used for SSN lookup
- `-EmpSQLCredential`: credential for the employee SQL connection
- `-FileServer`: file server hosting the share
- `-ShareName`: share name where CSV files are written
- `-FileServerCredential`: credential used to map the file share
- `-Wait`: pauses after each processed row for interactive review
- `-WhatIf` / `-wi`: dry-run mode; skips the normal long-running schedule loop

## Example Usage

Run from PowerShell without a profile, for example in a scheduled task or Jenkins job.

### Local dry run

```powershell
. .\params.ps1
.\Export-NewEmployeeCSV.ps1 @params -WhatIf -Verbose
```

### Scheduled task or CI invocation

```powershell
. .\jenkins.ps1
.\Export-NewEmployeeCSV.ps1 @scriptParams -Verbose
```

Example parameter blocks are available in [params.ps1](./params.ps1) and [jenkins.ps1](./jenkins.ps1).

## Operational Notes

- Ensure the credential objects referenced by the parameter blocks are already available in the session or injected by the calling job.
- The script creates a mapped PSDrive named `exports` that points to `\\<FileServer>\<ShareName>` and removes it at the end.
- In non-WhatIf mode, the script loops until 5:00 PM local time and sleeps 30 seconds between polls.
- In WhatIf mode, the script runs once and does not persist SQL status updates.

## Status Behavior

- `new` or `update`: CSV is exported.
- `written`: status set after CSV export.
- `complete`: set when file cleanup conditions are met.

Rows already marked `complete` are ignored in subsequent polling.

## Files Used

- [Export-NewEmployeeCSV.ps1](./Export-NewEmployeeCSV.ps1)
- [params.ps1](./params.ps1)
- [jenkins.ps1](./jenkins.ps1)
- [lib/attributes.txt](./lib/attributes.txt)
