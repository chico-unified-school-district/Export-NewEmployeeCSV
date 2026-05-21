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

function Export-UserDataToCSV {
 process {
  $_.exportFile = switch ($_) {
   { $_.int.status -eq 'new' } { $true; break }
   { $_.int.status -eq 'update' } { $true; break }
   { $_.int.status -ne 'complete' -and (!(Test-Path -Path $_.file.export)) } { $true; break } # File missing for some reason
   default { $false }
  }
  if ($_.exportFile -eq $false ) { return $_ }
  Write-Host ('{0},{1},{2},{3}' -f $MyInvocation.MyCommand.Name, $_.info, $_.file.export, $_.int.status) -f Blue
  $_.csvData | Export-Csv -Path $_.file.export -NoTypeInformation -Force -Encoding utf8
  $_
 }
}

function Format-CSVData {
 begin {
  $attributes = Get-Content -Path '.\lib\attributes.txt'
 }
 process {
  $_.csvData = [PSCustomObject]@{}
  foreach ($attrib in $attributes) {
   # if ($_.int.$attrib) { Write-Verbose ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.info, ($attrib + ': ' + $_.int.$attrib)) }
   # Int DB columns names ($_.int.$attrib) must match those in attributes.txt
   $_.csvData | Add-Member -MemberType NoteProperty -Name $attrib -Value $_.int.$attrib
  }
  $_.csvData.DateBirth = (Get-Date $_.csvData.DateBirth).ToString('MM/dd/yyyy')
  $_
 }
}

function New-PSObject {
 process {
  [PSCustomObject]@{
   emp        = $null
   csvData    = $null
   exportFile = $null
   isEntryOld = $null
   file       = $null
   info       = $null
   int        = $_
   removeFile = $null
   status     = $null
  }
 }
}

function Out-Object {
 process {
  Write-Verbose ($MyInvocation.MyCommand.name, $_ | Out-String)
  if ($Wait) { Read-Host ('{0}' -f ('x' * 50)) }
 }
}

function Update-OldEntry {
 begin {
  $cutoffDate = (Get-Date).AddMonths(-3)
 }
 process {
  $entryDate = Get-Date $_.int.dts
  $_.isEntryOld = if ($entryDate -lt $cutoffDate) { $true } else { $false }
  $_
 }
}

function Update-EmployeeData ($instance, $db, $table) {
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

function Update-FileInfo ($driveName, $server, $share) {
 process {
  $_.file = @{name = $null; export = $null; import = $null }
  $exportRoot = ('{0}:\' -f $driveName)
  $_.file.name = '{0}.csv' -f ($_.int.NameFirst + '_' + $_.int.NameLast + '_' + (Get-Date $_.int.dateAdded -f yyyy-MM-dd))
  $_.file.export = '{0}{1}' -f $exportRoot, $_.file.name # Path for csv export
  $_.file.import = ('\\{0}\{1}\{2}' -f $server, $share, $_.file.name) # Path for importing to employee mgmt system
  $_
 }
}

function Update-Info {
 process {
  $_.info = '[{0},dob: {1},row id: {2}]' -f ($_.int.NameFirst + ' ' + $_.int.NameLast), $_.int.DateBirth, $_.int.id
  $_
 }
}

function Update-Status {
 process {
  $_.status = switch ($_) {
   { $_.removeFile -eq $true } { 'complete'; break }
   { $_.exportFile -eq $true } { 'written' ; break }
   default { $_.int.status }
  }
  Write-Verbose ('{0},{1},{2},{3}' -f $MyInvocation.MyCommand.Name, $_.info, $_.exportPath, $_.status)
  $_
 }
}

function Update-IntDB ($sqlInstance, $table) {
 begin {
  $sql = "UPDATE $table SET status = @status, importFilePath = @importFilePath, dts = GETDATE() WHERE id = @id;"
 }
 process {
  if (($_.int.status -eq $_.status) -and ($_.int.importFilePath -eq $_.file.import)) { return $_ } # No changes
  $sqlVars = @{id = $_.int.id; status = $_.status; importFilePath = $_.file.import }
  Write-Verbose ('{0},{1},[{2}],[{3}]' -f $MyInvocation.MyCommand.Name, $_.info, $sql, ($sqlVars.Values -join ','))
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.info) -F Cyan
  if (!$WhatIf) { Invoke-DbaQuery -SqlInstance $sqlInstance -Query $sql -SqlParameters $sqlVars }
  $_
 }
}

function Remove-CsvFile ($drive) {
 process {
  $_.removeFile = switch ($_) {
   { !$_.emp -and $_.isEntryOld } { $true ; break } # Remove if old
   { $_.emp } { $true; break } # Remove if exists in employee system
   default { $false } # Do not delete
  }
  if ($_.removeFile -eq $false) { return $_ }
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.info, $_.file.export)
  if (Test-Path -Path $_.file.export) { Remove-Item -Path $_.file.export -Confirm:$false -WhatIf:$WhatIf }
  $_
 }
}

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
$newAccountSql = "SELECT * FROM {0} WHERE status NOT IN ('complete')" -f $IntSQLTable

$stopTime = if ($WhatIf) { Get-Date } else { Get-Date '5:00pm' }
$delay = if ($WhatIf) { 0 } else { 30 } # Wait x seconds between runs

do {
 Invoke-DbaQuery -SqlInstance $intSQLInstance -Query $newAccountSql | ConvertTo-Csv | ConvertFrom-Csv |
  New-PSObject |
   Update-Info |
    Update-EmployeeData -instance $empSQLInstance -table $EmpSQLTable |
     Update-FileInfo -driveName $driveName -server $FileServer -share $ShareName |
      Update-OldEntry |
       Format-CSVData |
        Export-UserDataToCSV |
         Remove-CsvFile -drive $driveName |
          Update-Status |
           Update-IntDB -sqlInstance $intSQLInstance -table $IntSQLTable |
            Out-Object

 Clear-SessionData

 if (!$WhatIf) { Write-Verbose ('Next Run: {0}' -f ((Get-Date).AddSeconds($delay))) }
 Start-Sleep $delay
} until ( $WhatIf -or ((Get-Date) -ge $stopTime) )

if (Test-Path -Path ('{0}:\' -f $driveName)) { Remove-PSDrive -Name $driveName -Force }

if ($WhatIf) { Show-TestRun }