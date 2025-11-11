# ==============================================================
# Run-Import.ps1  –  Dynamic VBA Loader + Live Log Streamer
# ==============================================================

$srcPath   =  "C:\Users\coadyj\projects\YE-AutomationV2\src"
$tempDir   =  "C:\Users\coadyj\projects\YE-AutomationV2\Temp"
$logFolder = "C:\Users\coadyj\projects\YE-AutomationV2\Logs"
$macroName = "ImportDailyYEData"

if (-not (Test-Path $tempDir)) { New-Item -ItemType Directory -Path $tempDir | Out-Null }

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$tempFile  = Join-Path $tempDir ("TempHost_" + $timestamp + ".xlsm")

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Add()
$workbook.SaveAs($tempFile, 52)

Start-Sleep -Seconds 2
$vbaProject  = $workbook.VBProject
$sourceFiles = Get-ChildItem -Path $srcPath -Recurse -Include *.bas, *.cls

foreach ($file in $sourceFiles) {
    Write-Host "Importing $($file.Name)"
    try   { $vbaProject.VBComponents.Import($file.FullName) }
    catch { Write-Host "[WARN] Failed to import $($file.Name): $($_.Exception.Message)" }
}

Write-Host "Running Import: $macroName"

try   { $excel.Run($macroName); Write-Host "`n[INFO] Macro executed successfully." }
catch { Write-Host "[ERROR] $($_.Exception.Message)" }

$workbook.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)    | Out-Null

try {
    if (Test-Path $tempFile) {
        Remove-Item $tempFile -Force
        Write-Host "[INFO] Deleted temp file: $tempFile"
    }
} catch {
    Write-Host "[WARN] Could not delete temp file (possibly locked)."
}

Write-Host "`nSupervisor Import Completed"
Write-Host "Created Backup of Master File and Deployable XSLM Version of Master."


$latestLog = Get-ChildItem -Path $logFolder -Filter *.txt | 
             Sort-Object LastWriteTime -Descending | 
             Select-Object -First 1

Write-Host "[INFO] Closing orphaned Excel instances..." -ForegroundColor Cyan

Get-Process excel -ErrorAction SilentlyContinue | ForEach-Object {
    try { $_.CloseMainWindow() | Out-Null } catch {}
    Start-Sleep -Milliseconds 300
    try { $_.Kill() } catch {}
}

Start-Sleep -Seconds 1


if ($latestLog) {
    Write-Host "`nOpening latest log: $($latestLog.Name)"
    Invoke-Item $latestLog.FullName
} else {
    Write-Host "`n[WARN] No log files found in $logFolder"
}
