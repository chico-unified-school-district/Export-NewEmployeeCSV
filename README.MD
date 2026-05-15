# Create-StaffAccount

Monitors an intermediate SQL table for new employee rows and writes a CSV file into a protected file share.

## Dependencies

- PowerShell 5.1 or later
- `dbatools`
- `CommonScriptFunctions`

## What It Does

The script reads rows from the configured SQL table where `status = 'new'`, converts each record into the CSV export format defined in `lib/attributes.txt`, and writes the file to the mapped `export:` drive.

It can be run as a one-time dry run with `-WhatIf` or scheduled to loop until the configured stop time.

## Parameters

- `-SQLServer`: SQL instance that hosts the intermediate table
- `-SQLDatabase`: database name containing the table
- `-SQLTable`: source table name
- `-SQLCredential`: credential used to connect to SQL Server
- `-FileServer`: file server hosting the share
- `-ShareName`: share name where CSV files are written
- `-FileServerCredential`: credential used to map the file share
- `-Wait`: pauses after each processed row for interactive review
- `-WhatIf` / `-wi`: dry-run mode; skips the normal long-running schedule loop

## Example Usage

Run from PowerShell without a profile, for example in a scheduled task or Jenkins job.

### Local dry run

```powershell
.\Export-NewEmployeeCSV.ps1 @params -WhatIf -Verbose
```

### Scheduled task or CI invocation

```powershell
.\Export-NewEmployeeCSV.ps1 @scriptParams -Verbose
```

Example parameter blocks are available in [params.ps1](./params.ps1) and [jenkins.ps1](./jenkins.ps1).

## Operational Notes

- Ensure the credential objects referenced by the parameter blocks are already available in the session or injected by the calling job.
- The script creates or reuses a mapped PSDrive named `export` that points to `\\<FileServer>\<ShareName>`.
- In non-WhatIf mode, the script loops until 5:00 PM local time and sleeps between polls.

## Files Used

- [Export-NewEmployeeCSV.ps1](./Export-NewEmployeeCSV.ps1)
- [params.ps1](./params.ps1)
- [jenkins.ps1](./jenkins.ps1)
- [lib/attributes.txt](./lib/attributes.txt)
