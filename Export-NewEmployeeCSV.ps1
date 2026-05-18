[cmdletbinding()]
param(
 [string]$IntSQlServer,
 [string]$IntSQLDatabase,
 [string]$IntSQLTable,
 [System.Management.Automation.PSCredential]$IntSQLCredential,
 [string]$EmpSQlServer,
 [string]$EmpSQLDatabase,
 [string]$EmpSQLTable,
 [System.Management.Automation.PSCredential]$EmpSQLCredential,
 [string]$FileServer,
 [string]$ShareName,
 [System.Management.Automation.PSCredential]$FileServerCredential,
 [switch]$Wait,
 [Alias('wi')][switch]$WhatIf
)

# CSV to Fileshare
function Format-CSVData {
 begin {
  $attributes = Get-Content -Path '.\lib\attributes.txt'
 }
 process {
  $_.csvData = [PSCustomObject]@{}
  foreach ($attrib in $attributes) {
   if ($_.emp.$attrib) { Write-Verbose ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.info, ($attrib + ': ' + $_.emp.$attrib)) }
   $_.csvData | Add-Member -MemberType NoteProperty -Name $attrib -Value $_.emp.$attrib
  }
  $_.csvData.DateBirth = (Get-Date $_.csvData.DateBirth).ToString('MM/dd/yyyy')
  $_
 }
}

function Format-ExportObject {
 process {
  [PSCustomObject]@{
   emp        = $_
   csvData    = $null
   fileName   = $null
   exportPath = $null
   info       = $null
   importFile = $null
   status     = $_.status
  }
 }
}

function Set-ExportInfo ($driveName, $server, $share) {
 process {
  $exportRoot = ('{0}:\' -f $driveName)
  $_.fileName = '{0}.csv' -f ($_.emp.NameFirst + '_' + $_.emp.NameLast + '_' + $_.emp.SSN.Substring($_.emp.SSN.Length - 4))
  $_.exportPath = '{0}{1}' -f $exportRoot, $_.fileName
  $_.importFile = ('\\{0}\{1}\{2}' -f $server, $share, $_.fileName)
  $_
 }
}

function Set-Info {
 process {
  $_.info = '[{0},{1},{2}]' -f ($_.emp.NameFirst + ' ' + $_.emp.NameLast), $_.emp.DateBirth, $_.emp.id
  $_
 }
}

function Set-Status {
 process {
  $content = Get-Content -Path $_.exportPath -ErrorAction SilentlyContinue
  if (!(Test-Path -Path $_.exportPath) -or $content -notmatch '\w') {
   $_.status = 'error'
   Write-Host ('{0},{1},Error: {2}' -f $MyInvocation.MyCommand.Name, $_.info, $_.exportPath) -F Red
   return $_
  }
  $_.status = 'Pending Review and New User File Import'
  $_
 }
}

function Show-Object {
 process {
  Write-Verbose ($MyInvocation.MyCommand.name, $_ | Out-String)
  if ($Wait) { Read-Host ('{0}' -f ('x' * 50)) }
 }
}

function Update-IntDB ($sqlInstance, $table) {
 begin {
  $sql = "UPDATE $table SET status = @status, importFilePath = @importFilePath WHERE id = @id;"
 }
 process {
  $sqlVars = @{id = $_.emp.id; status = $_.status; importFilePath = $_.importFile }
  Write-Verbose ('{0},{1},[{2}],[{3}]' -f $MyInvocation.MyCommand.Name, $_.info, $sql, ($sqlVars.Values -join ','))
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.info) -F Cyan
  #TODO
  # if (!$WhatIf) { Invoke-DbaQuery -SqlInstance $sqlInstance -Query $sql -SqlParameters $sqlVars }
  $_
 }
}

function Write-CSVFile {
 process {
  if ((Test-Path -Path $_.exportPath) -and ($_.status -ne 'update')) {
   Write-Host ('{0},{1},File Exists: {2}' -f $MyInvocation.MyCommand.Name, $_.info, $_.exportPath) -F Yellow
   return $_
  }
  try {
   $_.csvData | Export-Csv -Path $_.exportPath -NoTypeInformation -Force -Encoding utf8
  }
  catch {
   Write-Host ('{0},Error Exporting CSV: {1}' -f $MyInvocation.MyCommand.Name, $_.exportPath) -F Red
   $exportStatus = $false
  }
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.exportPath) -f Blue
  if ($exportStatus) { $_ }
 }
}
# ======================================================================================================================

# Remove completed files
function Format-RemoveObject {
 process {
  [PSCustomObject]@{
   file     = $_.Name
   emp      = $null
   info     = $null
   imported = $null
  }
 }
}

function Get-CsvData ($drive) {
 process {
  $filePath = ($drive + ':\' + $_.file)
  Write-Verbose ('{0},{1}' -f $MyInvocation.MyCommand.Name, $filePath)
  $_.emp = Import-Csv -Path $filePath
  $_
 }
}

function Get-EmployeeData ($instance, $db, $table) {
 begin {
  $sql = "SELECT empId,nameLast FROM $table WHERE SSNumIdFull = @ssn;"
 }
 process {
  $sqlVars = @{ssn = $_.emp.SSN }
  Write-Verbose ('{0},{1},[{2}],[{3}],[{4}]' -f $MyInvocation.MyCommand.Name, $_.info, $sql, $db, ($sqlVars.Values -join ','))
  $result = Invoke-DbaQuery -SqlInstance $instance -Query $sql -SqlParameters $sqlVars
  if ($result) { $_.imported = $true }
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.info, $_.imported) -F Blue
  $_
 }
}

function Remove-CSVfile ($drive) {
 process {
  if (!$_.imported) { return $_ }
  $filePath = ($drive + ':\' + $_.file)
  if (!(Test-Path -Path $filePath)) { return }
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.info, $filePath)
  Remove-Item -Path $filePath -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

# =====================================================

# ===================================================== main =====================================================
Import-Module -Name dbatools -Cmdlet Set-DbatoolsConfig, Invoke-DbaQuery, Connect-DbaInstance, Disconnect-DbaInstance
Import-Module -Name CommonScriptFunctions -Cmdlet Show-TestRun, New-SqlOperation, Clear-SessionData

Show-BlockInfo Main
if ($WhatIf) { Show-TestRun }


Clear-SessionData

$intSQLInstance = Connect-DbaInstance -SqlInstance $intSQLServer -Database $IntSQLDatabase -SqlCredential $IntSQLCredential
$empSQLInstance = Connect-DbaInstance -SqlInstance $EmpSQLServer -Database $EmpSQLDatabase -SqlCredential $EmpSQLCredential
# $empSQLInstance

$driveName = 'exports'
$networkPath = '\\{0}\{1}' -f $FileServer, $ShareName
$driveParams = @{
 Name        = $driveName
 PSProvider  = 'FileSystem'
 Root        = $networkPath
 Credential  = $FileServerCredential
 ErrorAction = 'Stop'
}
New-PSDrive @driveParams | Out-Null

# $newAccountSql = 'SELECT * FROM {0}' -f $SQLTable
$newAccountSql = "SELECT * FROM {0} WHERE status IN ('new','update')" -f $IntSQLTable

$stopTime = if ($WhatIf) { Get-Date } else { Get-Date '5:00pm' }
$delay = if ($WhatIf) { 0 } else { 5 } # Wait x seconds between runs

do {
 $newAccounts = Invoke-DbaQuery -SqlInstance $intSQLInstance -Query $newAccountSql | ConvertTo-Csv | ConvertFrom-Csv

 # Write-Output CSV to Fileshare, Update Status in DB, Show Output
 $newAccounts |
  Format-ExportObject |
   Set-Info |
    Set-ExportInfo -driveName $driveName -server $FileServer -share $ShareName |
     Format-CSVData |
      Write-CSVFile |
       Set-Status |
        Update-IntDB -sqlInstance $intSQLInstance -table $SQLTable |
         Show-Object

 # Remove csv files for imported employees
 Get-ChildItem -Path "${driveName}:\" -Filter *.csv |
  Format-RemoveObject |
   Get-CsvData -drive $driveName |
    Set-Info |
     Get-EmployeeData -instance $empSQLInstance -table $EmpSQLTable |
      Remove-CSVfile -drive $driveName |
       Show-Object

 Clear-SessionData

 if (!$WhatIf) { Write-Verbose ('Next Run: {0}' -f ((Get-Date).AddSeconds($delay))) }
 Start-Sleep $delay
} until ( $WhatIf -or ((Get-Date) -ge $stopTime) )

if (Test-Path -Path ('{0}:\' -f $driveName)) { Remove-PSDrive -Name $driveName -Force }

if ($WhatIf) { Show-TestRun }