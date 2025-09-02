# Simple PowerShell API to Excel Script
# Install Excel module if needed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}
Import-Module ImportExcel

# Configuration - Change these if needed
$ApiUrl = "https://api.restful-api.dev/objects"
$ExcelFile = "API_Results.xlsx"

# Your data to send (you can modify this)
$MyData = @{
    name = "Apple MacBook Pro 18 Max"
    data = @{
        year = 2021
        price = 189.99
        "CPU model" = "Intel Core i8"
        "Hard disk size" = "4 TB"
    }
}

Write-Host "Starting API calls..." -ForegroundColor Green

# Step 1: Make POST request
Write-Host "1. Making POST request..." -ForegroundColor Yellow
try {
    $jsonBody = $MyData | ConvertTo-Json -Depth 10
    $postResponse = Invoke-RestMethod -Uri $ApiUrl -Method POST -Body $jsonBody -ContentType "application/json"
    Write-Host "✓ POST Success! Created ID: $($postResponse.id)" -ForegroundColor Green
}
catch {
    Write-Host "✗ POST Failed: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Step 2: Save POST result to Excel
Write-Host "2. Saving POST result to Excel..." -ForegroundColor Yellow
$postResult = [PSCustomObject]@{
    Operation = "POST"
    ID = $postResponse.id
    Name = $postResponse.name
    Year = $postResponse.data.year
    Price = $postResponse.data.price
    CPU = $postResponse.data."CPU model"
    HardDisk = $postResponse.data."Hard disk size"
    CreatedAt = $postResponse.createdAt
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}
$postResult | Export-Excel -Path $ExcelFile -WorksheetName "Results" -AutoSize -BoldTopRow
Write-Host "✓ POST result saved to Excel" -ForegroundColor Green

# Step 3: Make GET request using the ID from POST
Write-Host "3. Making GET request..." -ForegroundColor Yellow
try {
    $getUrl = "$ApiUrl/$($postResponse.id)"
    $getResponse = Invoke-RestMethod -Uri $getUrl -Method GET
    Write-Host "✓ GET Success! Retrieved: $($getResponse.name)" -ForegroundColor Green
}
catch {
    Write-Host "✗ GET Failed: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Step 4: Save GET result to Excel
Write-Host "4. Adding GET result to Excel..." -ForegroundColor Yellow
$getResult = [PSCustomObject]@{
    Operation = "GET"
    ID = $getResponse.id
    Name = $getResponse.name
    Year = $getResponse.data.year
    Price = $getResponse.data.price
    CPU = $getResponse.data."CPU model"
    HardDisk = $getResponse.data."Hard disk size"
    CreatedAt = $getResponse.createdAt
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}
$getResult | Export-Excel -Path $ExcelFile -WorksheetName "Results" -Append -AutoSize
Write-Host "✓ GET result added to Excel" -ForegroundColor Green

# Step 5: Done!
Write-Host "`nAll done! 🎉" -ForegroundColor Cyan
Write-Host "Excel file created: $ExcelFile" -ForegroundColor Cyan
Write-Host "Total operations: 2 (1 POST + 1 GET)" -ForegroundColor Cyan

# Ask if user wants to open Excel file
$openFile = Read-Host "`nOpen Excel file now? (Y/N)"
if ($openFile -eq 'Y' -or $openFile -eq 'y') {
    Invoke-Item $ExcelFile
}