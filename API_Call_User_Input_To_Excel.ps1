# User input endpoint directly in the console or give the file path 
# User send the payload directly in the console or give the file path

# Install ImportExcel if not available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}
Import-Module ImportExcel

$ExcelFile = "API_Results.xlsx"

function Test-FileLock { 
    param([string]$Path) 
    if (-not (Test-Path $Path)) { return $false } 
    try { $f=[IO.File]::Open($Path,'Open','ReadWrite','None');$f.Close();$false } catch { $true } 
}

Write-Host "=== API Automation ===" -ForegroundColor Cyan

# Ask for endpoint
$optApi = Read-Host "Enter POST URL directly or type 'file' to load from file"
$ApiUrl = if ($optApi -eq 'file') { Get-Content (Read-Host "Enter endpoint file path") } else { $optApi }

# Ask for payload
$optPayload = Read-Host "Enter payload directly (JSON) or type 'file' to load from file"
$MyData = if ($optPayload -eq 'file') { 
    Get-Content (Read-Host "Enter payload file path") | ConvertFrom-Json 
} else { 
    $optPayload | ConvertFrom-Json 
}

Write-Host "1. Making POST request..." -ForegroundColor Yellow
try {
    $post = Invoke-RestMethod -Uri $ApiUrl -Method POST -Body ($MyData | ConvertTo-Json -Depth 5) -ContentType "application/json"
    Write-Host "✓ POST Success! ID: $($post.id)" -ForegroundColor Green
}
catch { Write-Host "✗ POST Failed: $($_.Exception.Message)" -ForegroundColor Red; exit }

$postResult = [PSCustomObject]@{
    Operation="POST"; ID=$post.id; Name=$post.name
    Year=$post.data.year; Price=$post.data.price
    CPU=$post.data."CPU model"; HardDisk=$post.data."Hard disk size"
    CreatedAt=$post.createdAt; Timestamp=(Get-Date)
}

Write-Host "2. Making GET request..." -ForegroundColor Yellow
try {
    $get = Invoke-RestMethod -Uri "$ApiUrl/$($post.id)" -Method GET
    Write-Host "✓ GET Success! Retrieved: $($get.name)" -ForegroundColor Green
}
catch { Write-Host "✗ GET Failed: $($_.Exception.Message)" -ForegroundColor Red; exit }

$getResult = [PSCustomObject]@{
    Operation="GET"; ID=$get.id; Name=$get.name
    Year=$get.data.year; Price=$get.data.price
    CPU=$get.data."CPU model"; HardDisk=$get.data."Hard disk size"
    CreatedAt=$get.createdAt; Timestamp=(Get-Date)
}

Write-Host "3. Saving results to Excel..." -ForegroundColor Yellow
try {
    if (Test-FileLock $ExcelFile) {
        $target = "API_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
        Write-Host "⚠️ Excel is open, saving to new temp file: $target" -ForegroundColor Yellow
    } else {
        $target = $ExcelFile
    }

    $allData = if (Test-Path $target) { Import-Excel $target -WorksheetName "Results" } else { @() }
    $allData += $postResult,$getResult
    $allData | Export-Excel $target -WorksheetName "Results" -AutoSize -BoldTopRow

    Write-Host "✓ Results saved to $target" -ForegroundColor Green
}
catch {
    Write-Host "✗ Failed to save results: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nAll done! 🎉" -ForegroundColor Cyan

# Ask if user wants to open Excel
$openFile = Read-Host "`nOpen Excel file now? (Y/N)"
if ($openFile -match '^(Y|y)$') {
    Invoke-Item $target
}

