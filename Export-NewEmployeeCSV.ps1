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
   if ($_.int.$attrib) { Write-Verbose ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.info, ($attrib + ': ' + $_.int.$attrib)) }
   $_.csvData | Add-Member -MemberType NoteProperty -Name $attrib -Value $_.int.$attrib
  }
  $_.csvData.DateBirth = (Get-Date $_.csvData.DateBirth).ToString('MM/dd/yyyy')
  $_
 }
}

function Format-Object {
 process {
  [PSCustomObject]@{
   emp        = $null
   csvData    = $null
   isEntryOld = $null
   fileName   = $null
   fileStatus = $null
   exportPath = $null
   info       = $null
   int        = $_
   importFile = $null
   status     = $_.status
  }
 }
}

function Set-CutoffStatus {
 begin {
  $cutoffDate = (Get-Date).AddMonths(-3)
 }
 process {
  $entryDate = Get-Date $_.int.dts
  $_.isEntryOld = if ($entryDate -lt $cutoffDate) { $true } else { $false }
  $_
 }
}

function Set-EmployeeData ($instance, $db, $table) {
 begin {
  $sql = "SELECT empId,nameLast FROM $table WHERE SSNumIdFull = @ssn;"
 }
 process {
  $sqlVars = @{ssn = $_.int.SSN }
  Write-Verbose ('{0},{1},[{2}],[{3}],[{4}]' -f $MyInvocation.MyCommand.Name, $_.info, $sql, $db, ($sqlVars.Values -join ','))
  $_.emp = Invoke-DbaQuery -SqlInstance $instance -Query $sql -SqlParameters $sqlVars | ConvertTo-Csv | ConvertFrom-Csv
  $_
 }
}

function Set-FileInfo ($driveName, $server, $share) {
 process {
  $exportRoot = ('{0}:\' -f $driveName)
  $_.fileName = '{0}.csv' -f ($_.int.NameFirst + '_' + $_.int.NameLast + '_' + $_.int.SSN.Substring($_.int.SSN.Length - 4))
  $_.exportPath = '{0}{1}' -f $exportRoot, $_.fileName
  $_.importFile = ('\\{0}\{1}\{2}' -f $server, $share, $_.fileName)
  $_
 }
}

function Set-Info {
 process {
  $_.info = '[{0},{1},{2}]' -f ($_.int.NameFirst + ' ' + $_.int.NameLast), $_.int.DateBirth, $_.int.id
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
  $sqlVars = @{id = $_.int.id; status = $_.status; importFilePath = $_.importFile }
  Write-Verbose ('{0},{1},[{2}],[{3}]' -f $MyInvocation.MyCommand.Name, $_.info, $sql, ($sqlVars.Values -join ','))
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.info) -F Cyan
  #TODO
  # if (!$WhatIf) { Invoke-DbaQuery -SqlInstance $sqlInstance -Query $sql -SqlParameters $sqlVars }
  $_
 }
}

function Write-CSVFile {
 process {
  $_.fileStatus = switch ($_) {
   { $_.emp } { @{skip = $true; msg = 'Already Imported' }; break }
   { $_.int.status -eq 'update' } { @{writeFile = $true; msg = 'Updated file written' }; break }
   { Test-Path -Path $_.exportPath } { @{writeFile = $false; msg = 'File exists' }; break }
   default { @{writeFile = $true ; msg = 'New file written' }; break }
  }
  Write-Host ('{0},{1},{2},{3}' -f $MyInvocation.MyCommand.Name, $_.info, $_.exportPath, $_.fileStatus.msg) -f Blue
  if ($_.fileStatus.skip) { return $_ }
  $_.csvData | Export-Csv -Path $_.exportPath -NoTypeInformation -Force -Encoding utf8
  $_
 }
}
# function Write-CSVFile {
#  process {
#   if ($_.imported) { return $_ }
#   if ((Test-Path -Path $_.exportPath) -and ($_.status -ne 'update')) {
#    Write-Host ('{0},{1},File Exists: {2}' -f $MyInvocation.MyCommand.Name, $_.info, $_.exportPath) -F Yellow
#    return $_
#   }
#   Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.exportPath) -f Blue
#   $_.csvData | Export-Csv -Path $_.exportPath -NoTypeInformation -Force -Encoding utf8
#   $_
#  }
# }

# function Get-CsvData ($drive) {
#  process {
#   $filePath = ($drive + ':\' + $_.file)
#   Write-Verbose ('{0},{1}' -f $MyInvocation.MyCommand.Name, $filePath)
#   $_.int = Import-Csv -Path $filePath
#   $_
#  }
# }

function Remove-CsvFile ($drive) {
 process {
  $filePath = ($drive + ':\' + $_.fileName)
  if (!(Test-Path -Path $filePath)) { return } # Skip when file missing
  if (!$_.emp) { return } # Skip when not yet in Employee mgmt system
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.info, $filePath)
  Remove-Item -Path $filePath -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

# function Remove-CsvFile ($drive) {
#  process {
#   if (!$_.imported) { return $_ }
#   $filePath = ($drive + ':\' + $_.file)
#   if (!(Test-Path -Path $filePath)) { return }
#   Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.info, $filePath)
#   Remove-Item -Path $filePath -Confirm:$false -WhatIf:$WhatIf
#   $_
#  }
# }

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
$delay = if ($WhatIf) { 0 } else { 10 } # Wait x seconds between runs

do {
 Invoke-DbaQuery -SqlInstance $intSQLInstance -Query $newAccountSql | ConvertTo-Csv | ConvertFrom-Csv |
  Format-Object |
   Set-Info |
    Set-EmployeeData -instance $empSQLInstance -table $EmpSQLTable |
     Set-FileInfo -driveName $driveName -server $FileServer -share $ShareName |
      Set-CutoffStatus |
       Format-CSVData |
        Write-CSVFile |
         #       Set-Status |
         #        Update-IntDB -sqlInstance $intSQLInstance -table $SQLTable |
         #         Remove-CsvFile -drive $driveName |
         Show-Object

 # Remove csv files for imported employees
 # Get-ChildItem -Path "${driveName}:\" -Filter *.csv |
 #  Format-RemoveObject |
 #   Get-CsvData -drive $driveName |
 #    Set-Info |
 #     Get-EmployeeData -instance $empSQLInstance -table $EmpSQLTable |
 #      Remove-CsvFile -drive $driveName |
 #       Show-Object

 Clear-SessionData

 if (!$WhatIf) { Write-Verbose ('Next Run: {0}' -f ((Get-Date).AddSeconds($delay))) }
 Start-Sleep $delay
} until ( $WhatIf -or ((Get-Date) -ge $stopTime) )

if (Test-Path -Path ('{0}:\' -f $driveName)) { Remove-PSDrive -Name $driveName -Force }

if ($WhatIf) { Show-TestRun }