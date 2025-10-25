# ==============================================================
# Run-Import.ps1  –  Dynamic VBA Loader + Live Log Streamer
# ==============================================================

$srcPath   = "C:\Users\coadyj\projects\YE-AutomationV2\src"
$tempDir   = "C:\Users\coadyj\projects\YE-AutomationV2\Temp"
$macroName = "ImportDailyYEData"

# --- Ensure temp dir exists ---
if (-not (Test-Path $tempDir)) { New-Item -ItemType Directory -Path $tempDir | Out-Null }

# --- Temp workbook path ---
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$tempFile  = Join-Path $tempDir ("TempHost_" + $timestamp + ".xlsm")

Write-Host "`n========================================================"
Write-Host "Creating temporary Excel host workbook: $tempFile"
Write-Host "========================================================"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Add()
$workbook.SaveAs($tempFile, 52)   # .xlsm

Start-Sleep -Seconds 2
$vbaProject  = $workbook.VBProject
$sourceFiles = Get-ChildItem -Path $srcPath -Recurse -Include *.bas, *.cls


foreach ($file in $sourceFiles) {
    Write-Host "Importing $($file.Name)"
    try   { $vbaProject.VBComponents.Import($file.FullName) }
    catch { Write-Host "[WARN] Failed to import $($file.Name): $($_.Exception.Message)" }
}

Write-Host "Running macro: $macroName"

try   { $excel.Run($macroName); Write-Host "[INFO] Macro executed successfully." }
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

Write-Host "`n========================================================"
Write-Host "Supervisor import completed from dynamic temp workbook."
Write-Host "========================================================"
