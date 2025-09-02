# PowerShell Script for API Calls with Excel Export
# Requires ImportExcel module: Install-Module -Name ImportExcel -Force

# Check and install required modules
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

Import-Module ImportExcel

# Configuration
$BaseUrl = "https://api.restful-api.dev/objects"
$ExcelPath = "API_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
$WorksheetName = "API_Results"

# Sample data for testing multiple iterations
$TestData = @(
    @{
        name = "Apple MacBook Pro 16"
        data = @{
            year = 2019
            price = 1849.99
            "CPU model" = "Intel Core i9"
            "Hard disk size" = "2 TB"
        }
    },
    @{
        name = "Dell XPS 13"
        data = @{
            year = 2023
            price = 1299.99
            "CPU model" = "Intel Core i7"
            "Hard disk size" = "512 GB"
        }
    },
    @{
        name = "HP EliteBook 840"
        data = @{
            year = 2022
            price = 1099.99
            "CPU model" = "AMD Ryzen 7"
            "Hard disk size" = "1 TB"
        }
    }
)

# Function to make POST API call
function Invoke-PostAPI {
    param(
        [string]$Url,
        [hashtable]$RequestBody
    )
    
    try {
        $jsonBody = $RequestBody | ConvertTo-Json -Depth 10
        Write-Host "Making POST request to: $Url" -ForegroundColor Green
        Write-Host "Request Body: $jsonBody" -ForegroundColor Cyan
        
        $headers = @{
            "Content-Type" = "application/json"
        }
        
        $response = Invoke-RestMethod -Uri $Url -Method POST -Body $jsonBody -Headers $headers -ErrorAction Stop
        Write-Host "POST Success - ID: $($response.id)" -ForegroundColor Green
        return $response
    }
    catch {
        Write-Error "POST API call failed: $($_.Exception.Message)"
        return $null
    }
}

# Function to make GET API call
function Invoke-GetAPI {
    param(
        [string]$Url,
        [string]$ObjectId
    )
    
    try {
        $getUrl = "$Url/$ObjectId"
        Write-Host "Making GET request to: $getUrl" -ForegroundColor Green
        
        $response = Invoke-RestMethod -Uri $getUrl -Method GET -ErrorAction Stop
        Write-Host "GET Success - Retrieved: $($response.name)" -ForegroundColor Green
        return $response
    }
    catch {
        Write-Error "GET API call failed: $($_.Exception.Message)"
        return $null
    }
}

# Function to convert nested object to flat structure for Excel
function ConvertTo-FlatObject {
    param(
        [object]$InputObject,
        [string]$Prefix = ""
    )
    
    $flatObject = @{}
    
    foreach ($property in $InputObject.PSObject.Properties) {
        $key = if ($Prefix) { "$Prefix.$($property.Name)" } else { $property.Name }
        
        if ($property.Value -is [PSCustomObject] -or $property.Value -is [hashtable]) {
            $nestedFlat = ConvertTo-FlatObject -InputObject $property.Value -Prefix $key
            foreach ($nestedProp in $nestedFlat.GetEnumerator()) {
                $flatObject[$nestedProp.Key] = $nestedProp.Value
            }
        } else {
            $flatObject[$key] = $property.Value
        }
    }
    
    return $flatObject
}

# Function to add row to Excel iteratively
function Add-RowToExcel {
    param(
        [string]$ExcelPath,
        [string]$WorksheetName,
        [object]$Data,
        [int]$RowNumber,
        [string]$Operation
    )
    
    try {
        # Convert nested object to flat structure
        $flatData = ConvertTo-FlatObject -InputObject $Data
        
        # Add operation type and timestamp
        $flatData["Operation"] = $Operation
        $flatData["Timestamp"] = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $flatData["Row_Number"] = $RowNumber
        
        # Convert to PSCustomObject for better Excel handling
        $excelRow = [PSCustomObject]$flatData
        
        # Check if file exists to determine if we need headers
        $fileExists = Test-Path $ExcelPath
        
        if (-not $fileExists) {
            # First row - create file with headers
            $excelRow | Export-Excel -Path $ExcelPath -WorksheetName $WorksheetName -AutoSize -BoldTopRow -FreezeTopRow
            Write-Host "Created Excel file: $ExcelPath (Row $RowNumber)" -ForegroundColor Magenta
        } else {
            # Append to existing file
            $excelRow | Export-Excel -Path $ExcelPath -WorksheetName $WorksheetName -Append -AutoSize
            Write-Host "Added row $RowNumber to Excel file" -ForegroundColor Magenta
        }
        
        return $true
    }
    catch {
        Write-Error "Failed to add row to Excel: $($_.Exception.Message)"
        return $false
    }
}

# Main execution function
function Start-APIToExcelProcess {
    Write-Host "Starting API to Excel Process..." -ForegroundColor Yellow
    Write-Host "Output file will be: $ExcelPath" -ForegroundColor Yellow
    Write-Host "=" * 60 -ForegroundColor Yellow
    
    $rowCounter = 1
    $createdObjects = @()
    
    # Phase 1: POST Operations (Creating objects)
    Write-Host "`nPhase 1: Creating objects via POST API..." -ForegroundColor Blue
    
    foreach ($item in $TestData) {
        Write-Host "`n--- Processing Item $rowCounter ---" -ForegroundColor White
        
        # Make POST API call
        $postResponse = Invoke-PostAPI -Url $BaseUrl -RequestBody $item
        
        if ($postResponse) {
            # Store created object ID for GET operations
            $createdObjects += $postResponse.id
            
            # Add POST result to Excel iteratively
            $success = Add-RowToExcel -ExcelPath $ExcelPath -WorksheetName $WorksheetName -Data $postResponse -RowNumber $rowCounter -Operation "POST"
            
            if ($success) {
                Write-Host "Row $rowCounter added successfully" -ForegroundColor Green
            }
            
            # Small delay between operations
            Start-Sleep -Milliseconds 500
        }
        
        $rowCounter++
    }
    
    # Phase 2: GET Operations (Retrieving created objects)
    Write-Host "`nPhase 2: Retrieving objects via GET API..." -ForegroundColor Blue
    
    foreach ($objectId in $createdObjects) {
        Write-Host "`n--- Retrieving Object $rowCounter ---" -ForegroundColor White
        
        # Make GET API call
        $getResponse = Invoke-GetAPI -Url $BaseUrl -ObjectId $objectId
        
        if ($getResponse) {
            # Add GET result to Excel iteratively
            $success = Add-RowToExcel -ExcelPath $ExcelPath -WorksheetName $WorksheetName -Data $getResponse -RowNumber $rowCounter -Operation "GET"
            
            if ($success) {
                Write-Host "Row $rowCounter added successfully" -ForegroundColor Green
            }
            
            # Small delay between operations
            Start-Sleep -Milliseconds 500
        }
        
        $rowCounter++
    }
    
    # Final summary
    Write-Host "`n" + "=" * 60 -ForegroundColor Yellow
    Write-Host "Process completed!" -ForegroundColor Green
    Write-Host "Total rows processed: $($rowCounter - 1)" -ForegroundColor Green
    Write-Host "Excel file location: $ExcelPath" -ForegroundColor Green
    Write-Host "Created objects: $($createdObjects.Count)" -ForegroundColor Green
    
    # Open Excel file if desired
    $openFile = Read-Host "`nDo you want to open the Excel file? (Y/N)"
    if ($openFile -eq 'Y' -or $openFile -eq 'y') {
        try {
            Invoke-Item $ExcelPath
        }
        catch {
            Write-Host "Could not open Excel file automatically. Please open: $ExcelPath" -ForegroundColor Yellow
        }
    }
}

# Advanced function for custom data input
function Start-CustomAPIProcess {
    param(
        [array]$CustomData,
        [string]$CustomExcelPath = $null
    )
    
    if ($CustomExcelPath) {
        $script:ExcelPath = $CustomExcelPath
    }
    
    if ($CustomData) {
        $script:TestData = $CustomData
    }
    
    Start-APIToExcelProcess
}

# Function to process single item (for individual testing)
function Process-SingleItem {
    param(
        [hashtable]$Item,
        [string]$OutputPath = $null
    )
    
    if ($OutputPath) {
        $script:ExcelPath = $OutputPath
    }
    
    $postResponse = Invoke-PostAPI -Url $BaseUrl -RequestBody $Item
    
    if ($postResponse) {
        Add-RowToExcel -ExcelPath $ExcelPath -WorksheetName $WorksheetName -Data $postResponse -RowNumber 1 -Operation "POST"
        
        Start-Sleep -Seconds 1
        
        $getResponse = Invoke-GetAPI -Url $BaseUrl -ObjectId $postResponse.id
        if ($getResponse) {
            Add-RowToExcel -ExcelPath $ExcelPath -WorksheetName $WorksheetName -Data $getResponse -RowNumber 2 -Operation "GET"
        }
        
        Write-Host "Single item processed. Excel file: $ExcelPath" -ForegroundColor Green
    }
}

# Execute the main process
Write-Host "API to Excel PowerShell Script" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan

# Uncomment the line below to run the full process
Start-APIToExcelProcess

# Example of how to use custom data:
<#
$customItems = @(
    @{
        name = "Custom Item 1"
        data = @{
            year = 2024
            price = 999.99
            "CPU model" = "Custom CPU"
            "Hard disk size" = "500 GB"
        }
    }
)

Start-CustomAPIProcess -CustomData $customItems -CustomExcelPath "Custom_API_Results.xlsx"
#>

# Example of processing a single item:
<#
$singleItem = @{
    name = "Test Item"
    data = @{
        year = 2024
        price = 599.99
        "CPU model" = "Test CPU"
        "Hard disk size" = "256 GB"
    }
}

Process-SingleItem -Item $singleItem -OutputPath "Single_Item_Test.xlsx"
#>