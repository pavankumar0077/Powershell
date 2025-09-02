 # File can be save even if excel is opened

# Install ImportExcel if not available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}
Import-Module ImportExcel

# Config
$ApiUrl    = "https://api.restful-api.dev/objects"
$ExcelFile = "API_Results.xlsx"
$TempFile  = "API_Results_temp.xlsx"

# Function: Check if file is locked by Excel
function Test-FileLock {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $false }
    try {
        $fs = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
        $fs.Close()
        return $false
    }
    catch { return $true }
}

Write-Host "Starting API calls..." -ForegroundColor Green

# Step 1: POST request
Write-Host "1. Making POST request..." -ForegroundColor Yellow
try {
    $MyData = @{
        name = "Apple MacBook Pro 20 Max"
        data = @{
            year = 2021
            price = 189.99
            "CPU model" = "Intel Core i8"
            "Hard disk size" = "8TB"
        }
    }
    $post = Invoke-RestMethod -Uri $ApiUrl -Method POST -Body ($MyData | ConvertTo-Json -Depth 5) -ContentType "application/json"
    Write-Host "✓ POST Success! Created ID: $($post.id)" -ForegroundColor Green
}
catch {
    Write-Host "✗ POST Failed: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Step 2: Build POST object
$postResult = [PSCustomObject]@{
    Operation = "POST"
    ID        = $post.id
    Name      = $post.name
    Year      = $post.data.year
    Price     = $post.data.price
    CPU       = $post.data."CPU model"
    HardDisk  = $post.data."Hard disk size"
    CreatedAt = $post.createdAt
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}

# Step 3: GET request
Write-Host "2. Making GET request..." -ForegroundColor Yellow
try {
    $get = Invoke-RestMethod -Uri "$ApiUrl/$($post.id)" -Method GET
    Write-Host "✓ GET Success! Retrieved: $($get.name)" -ForegroundColor Green
}
catch {
    Write-Host "✗ GET Failed: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Step 4: Build GET object
$getResult = [PSCustomObject]@{
    Operation = "GET"
    ID        = $get.id
    Name      = $get.name
    Year      = $get.data.year
    Price     = $get.data.price
    CPU       = $get.data."CPU model"
    HardDisk  = $get.data."Hard disk size"
    CreatedAt = $get.createdAt
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}

# Step 5: Save results (choose file depending on lock)
Write-Host "3. Saving results to Excel..." -ForegroundColor Yellow
try {
    $targetFile = if (Test-FileLock $ExcelFile) { 
        Write-Host "⚠️ Excel is open, saving to temporary file: $TempFile" -ForegroundColor Yellow
        $TempFile 
    } else { 
        $ExcelFile 
    }

    $allData = @()
    if (Test-Path $targetFile) {
        $allData = Import-Excel $targetFile -WorksheetName "Results" -ErrorAction SilentlyContinue
    }
    $allData += $postResult, $getResult
    $allData | Export-Excel $targetFile -WorksheetName "Results" -AutoSize -BoldTopRow

    Write-Host "✓ Results saved to $targetFile" -ForegroundColor Green
}
catch {
    Write-Host "✗ Failed to save results: $($_.Exception.Message)" -ForegroundColor Red
}

# Step 6: Done
Write-Host "`nAll done! 🎉" -ForegroundColor Cyan
Write-Host "Excel file updated: $ExcelFile (or $TempFile if open)" -ForegroundColor Cyan
Write-Host "Total operations: 2 (1 POST + 1 GET)" -ForegroundColor Cyan

# Ask if user wants to open Excel
$openFile = Read-Host "`nOpen Excel file now? (Y/N)"
if ($openFile -match '^(Y|y)$') {
    Invoke-Item $ExcelFile
}
