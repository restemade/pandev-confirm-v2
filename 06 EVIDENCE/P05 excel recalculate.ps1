param([Parameter(Mandatory=$true)][string]$SourceModel,[Parameter(Mandatory=$true)][string]$OutputDirectory)
# Run in PowerShell 7 on Windows with Microsoft Excel installed.
# Recalculates a COPY; disables macros and external-link updates.
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$auditCopy=Join-Path ([IO.Path]::GetFullPath($OutputDirectory)) 'recalculated_model.xlsx'
if ([IO.Path]::GetFullPath($SourceModel) -eq $auditCopy) { throw 'Source and output must differ.' }
Copy-Item -LiteralPath $SourceModel -Destination $auditCopy -Force
$auditExcel=New-Object -ComObject Excel.Application
$auditExcel.Visible=$false
$auditExcel.DisplayAlerts=$false
$auditExcel.AskToUpdateLinks=$false
$auditExcel.AutomationSecurity=3
try {
 $auditBook=$auditExcel.Workbooks.Open($auditCopy,0,$false)
 # Original E03 source: sheet 3 contains assumptions; sheet 1 contains summary.
 $auditRows=@()
 foreach($auditScenario in 1,2,3){
  $auditBook.Worksheets.Item(3).Range('F19').Value2=$auditScenario
  $auditExcel.CalculateFullRebuild()
  $auditRows += [pscustomobject]@{scenario=$auditScenario;revenue2027=$auditBook.Worksheets.Item(1).Range('D25').Value2;ebitda2027=$auditBook.Worksheets.Item(1).Range('D26').Value2;arr2027=$auditBook.Worksheets.Item(1).Range('D28').Value2;staff2027=$auditBook.Worksheets.Item(1).Range('D30').Value2}
 }
 $auditBook.Worksheets.Item(3).Range('F19').Value2=1
 $auditExcel.CalculateFullRebuild()
 $auditBook.Save()
 $auditBook.Close($false)
 $auditRows | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutputDirectory 'scenarios.json') -Encoding utf8
} finally { $auditExcel.Quit() }
