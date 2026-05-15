[cmdletbinding()]
<#
#>
[cmdletbinding()]
param(
 [string]$SQlServer,
 [string]$SQLDatabase,
 [string]$SQLTable,
 [System.Management.Automation.PSCredential]$SQLCredential,
 [string]$FileServer,
 [string]$ShareName,
 [System.Management.Automation.PSCredential]$FileServerCredential,
 [switch]$Wait,
 [Alias('wi')][switch]$WhatIf
)

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

function Format-Object {
 process {
  [PSCustomObject]@{
   emp        = $_
   csvData    = $null
   fileName   = $null
   exportPath = $null
   info       = $null
   importFile = $null
   status     = $null
  }
 }
}

function Set-ExportInfo ($driveName, $server, $share) {
 process {
  $exportRoot = ('{0}:\' -f $driveName)
  $_.fileName = '{0}_{1:yyyy-MM-dd}.csv' -f ($_.emp.NameFirst + '_' + $_.emp.NameLast), (Get-Date)
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
 begin {
  $i = 0
 }
 process {
  $i++
  Write-Verbose ($MyInvocation.MyCommand.name, $_ | Out-String)
  if ($Wait) { Read-Host ('{0}' -f ('x' * 50)) }
 }
 end {
  Write-Host ('{0},Total Processed: {1}' -f $MyInvocation.MyCommand.Name, $i) -f Green
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
  try {
   $_.csvData | Export-Csv -Path $_.exportPath -NoTypeInformation -Force -WhatIf:$WhatIf
   (Get-Content -Path $_.exportPath) -replace '"', '' |
    Set-Content -Path $_.exportPath -Encoding UTF8 -WhatIf:$WhatIf
  }
  catch { Write-Host ('{0},Error Exporting CSV: {1}' -f $MyInvocation.MyCommand.Name, $_.exportPath) -F Red }
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.exportPath) -f Blue
  $_
 }
}

# ===================================================== main =====================================================
Import-Module -Name dbatools -Cmdlet Set-DbatoolsConfig, Invoke-DbaQuery, Connect-DbaInstance, Disconnect-DbaInstance
Import-Module -Name CommonScriptFunctions -Cmdlet Show-TestRun, New-SqlOperation, Clear-SessionData

Show-BlockInfo Main
if ($WhatIf) { Show-TestRun }


Clear-SessionData

$SQLInstance = Connect-DbaInstance -SqlInstance $SQLServer -Database $SQLDatabase -SqlCredential $SQLCredential

$exportDriveName = 'export'
$driveParams = @{
 Name        = $exportDriveName
 PSProvider  = 'FileSystem'
 Root        = ('\\{0}\{1}' -f $FileServer, $ShareName)
 Credential  = $FileServerCredential
 ErrorAction = 'Stop'
}
if (!(Test-Path -Path ('{0}:\' -f $exportDriveName))) {
 New-PSDrive @driveParams | Out-Null
}

# $newAccountSql = 'SELECT * FROM {0}' -f $SQLTable
$newAccountSql = "SELECT * FROM {0} WHERE status = 'new'" -f $SQLTable

$stopTime = if ($WhatIf) { Get-Date } else { Get-Date '5:00pm' }
$delay = if ($WhatIf) { 0 } else { 180 }


do {
 $newAccounts = Invoke-DbaQuery -SqlInstance $SQLInstance -Query $newAccountSql | ConvertTo-Csv | ConvertFrom-Csv
 if ($newAccounts -and !(Test-Path -Path ('{0}:\' -f $exportDriveName))) {
  if (Test-Path -Path ('{0}:\' -f $exportDriveName)) { Remove-PSDrive -Name $exportDriveName -Force }
  New-PSDrive @driveParams | Out-Null
 }

 $newAccounts |
  Format-Object |
   Set-Info |
    Set-ExportInfo -driveName $exportDriveName -server $FileServer -share $ShareName |
     Format-CSVData |
      Write-CSVFile |
       Set-Status |
        # Update-IntDB -sqlInstance $SQLInstance -table $SQLTable |
        Show-Object

 Clear-SessionData

 if (!$WhatIf) { Write-Verbose ('Next Run: {0}' -f ((Get-Date).AddSeconds($delay))) }
 Start-Sleep $delay
} until ( $WhatIf -or ((Get-Date) -ge $stopTime) )

if (Test-Path -Path ('{0}:\' -f $exportDriveName)) { Remove-PSDrive -Name $exportDriveName -Force }

if ($WhatIf) { Show-TestRun }