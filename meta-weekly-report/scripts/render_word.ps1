param([Parameter(Mandatory=$true)][string]$InputDocx,[Parameter(Mandatory=$true)][string]$OutputPdf)
$ErrorActionPreference = 'Stop'
$taskInput = (Resolve-Path -LiteralPath $InputDocx).Path
$taskOutput = [System.IO.Path]::GetFullPath($OutputPdf)
if ($taskInput -eq $taskOutput) { throw 'Input and output must differ.' }
$taskWord = New-Object -ComObject Word.Application
try {
    $taskWord.Visible = $false
    $taskWord.DisplayAlerts = 0
    $taskDoc = $taskWord.Documents.Open($taskInput, $false, $true)
    try { $taskDoc.ExportAsFixedFormat($taskOutput, 17) }
    finally { try { $taskDoc.Close(0) } catch {} }
} finally {
    try { $taskWord.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($taskWord) | Out-Null } catch {}
}
if (!(Test-Path -LiteralPath $taskOutput) -or (Get-Item -LiteralPath $taskOutput).Length -eq 0) { throw 'Word did not produce a PDF.' }
Get-Item -LiteralPath $taskOutput | Select-Object Name,Length
