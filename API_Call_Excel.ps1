# Install ImportExcel if not available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}
Import-Module ImportExcel

# Config
$ApiUrl = "https://api.restful-api.dev/objects"
$ExcelFile = "API_Results.xlsx"

Write-Host "Starting API calls..." -ForegroundColor Green

# Step 1: POST request
Write-Host "1. Making POST request..." -ForegroundColor Yellow
try {
    $MyData = @{
        name = "ChromeBook Plus"
        data = @{
            year = 2021
            price = 189.99
            "CPU model" = "Intel Core i8"
            "Hard disk size" = "8 TB"
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

# Step 5: Save both results safely
Write-Host "3. Saving results to Excel..." -ForegroundColor Yellow
try {
    $allData = @()
    if (Test-Path $ExcelFile) {
        # Load existing data if file is locked (read-only)
        $allData = Import-Excel $ExcelFile -WorksheetName "Results" -ErrorAction SilentlyContinue
    }
    $allData += $postResult, $getResult
    $allData | Export-Excel $ExcelFile -WorksheetName "Results" -AutoSize -BoldTopRow
    Write-Host "✓ Results saved to Excel (even if file was open)" -ForegroundColor Green
}
catch {
    Write-Host "✗ Failed to save results: $($_.Exception.Message)" -ForegroundColor Red
}

# Step 6: Done
Write-Host "`nAll done! 🎉" -ForegroundColor Cyan
Write-Host "Excel file updated: $ExcelFile" -ForegroundColor Cyan
Write-Host "Total operations: 2 (1 POST + 1 GET)" -ForegroundColor Cyan

# Ask if user wants to open Excel
$openFile = Read-Host "`nOpen Excel file now? (Y/N)"
if ($openFile -match '^(Y|y)$') {
    Invoke-Item $ExcelFile
}
