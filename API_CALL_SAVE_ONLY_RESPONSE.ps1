# === API Automation Script (Clean & Simple) ===

# GUI file picker
function Select-File {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $dlg.Filter = "All files (*.*)|*.*"
    if ($dlg.ShowDialog() -eq 'OK') { return $dlg.FileName }
    throw "No file selected."
}

# Check if file is locked (Excel open)
function Test-FileLock {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $false }
    try { $f=[IO.File]::Open($Path,'Open','ReadWrite','None');$f.Close();$false }
    catch { $true }
}

# Flatten nested JSON
function Flatten-Json {
    param($obj, $prefix = "")
    $flat = @{}
    foreach ($p in $obj.PSObject.Properties) {
        $name = if ($prefix) { "$prefix.$($p.Name)" } else { $p.Name }
        if ($p.Value -is [PSCustomObject]) {
            (Flatten-Json $p.Value $name).PSObject.Properties |
                ForEach-Object { $flat[$_.Name] = $_.Value }
        } else {
            $flat[$name] = $p.Value
        }
    }
    return [PSCustomObject]$flat
}

# Install ImportExcel if not present
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}
Import-Module ImportExcel

$ExcelFile = "API_Response.xlsx"

Write-Host "=== API Automation ===" -ForegroundColor Cyan
$ApiUrl  = Read-Host "Enter API URL"
$Method  = Read-Host "Enter HTTP method (GET/POST/PUT/DELETE)"

$Body = $null
if ($Method -in @("POST","PUT")) {
    $opt = Read-Host "Enter payload directly (JSON) or type 'file' to browse"
    $Body = if ($opt -eq 'file') { Get-Content (Select-File) -Raw } else { $opt }
}

# API call
Write-Host "1. Calling API ($Method $ApiUrl)..." -ForegroundColor Yellow
try {
    $resp = Invoke-RestMethod -Uri $ApiUrl -Method $Method -Body $Body -ContentType "application/json"
    Write-Host "✓ API call success!" -ForegroundColor Green
} catch {
    Write-Host "✗ API call failed: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Flatten response
Write-Host "2. Flattening response..." -ForegroundColor Yellow
$row = Flatten-Json $resp
$row | Add-Member -NotePropertyName Timestamp -NotePropertyValue (Get-Date) -Force

# Save to Excel
Write-Host "3. Saving to Excel..." -ForegroundColor Yellow
if (Test-FileLock $ExcelFile) {
    $target = "API_Response_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    Write-Host "⚠️ Excel is open, saving to $target" -ForegroundColor Yellow
} else {
    $target = $ExcelFile
}

$old = if (Test-Path $target) { Import-Excel $target -WorksheetName "Results" } else { @() }
$all = $old + $row
$all | Export-Excel $target -WorksheetName "Results" -AutoSize -BoldTopRow
Write-Host "✓ Results saved to $target" -ForegroundColor Green

Write-Host "`nAll done! 🎉" -ForegroundColor Cyan
if ((Read-Host "`nOpen Excel file now? (Y/N)") -match '^(Y|y)$') { Invoke-Item $target }
