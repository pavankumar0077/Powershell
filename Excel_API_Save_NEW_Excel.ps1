# Excel COM API Integration Script with PowerShell
# Author: DevOps Engineer
# Purpose: Use Excel COM APIs to read input, process data, and save to another Excel file

#Requires -Version 5.1

# Global variables for Excel COM objects
$global:ExcelApp = $null
$global:InputWorkbook = $null
$global:OutputWorkbook = $null

# Function to initialize Excel COM Application
function Initialize-ExcelApplication {
    try {
        Write-Host "Initializing Excel COM Application..." -ForegroundColor Green
        
        # Create Excel COM object
        $global:ExcelApp = New-Object -ComObject Excel.Application
        $global:ExcelApp.Visible = $false  # Keep Excel hidden
        $global:ExcelApp.DisplayAlerts = $false  # Disable Excel alerts
        $global:ExcelApp.ScreenUpdating = $false  # Improve performance
        
        Write-Host "Excel COM Application initialized successfully" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to initialize Excel COM Application: $($_.Exception.Message)"
        Write-Host "Make sure Microsoft Excel is installed on this system" -ForegroundColor Red
        return $false
    }
}

# Function to cleanup Excel COM objects
function Close-ExcelApplication {
    try {
        Write-Host "Cleaning up Excel COM objects..." -ForegroundColor Yellow
        
        # Close workbooks if open
        if ($global:InputWorkbook) {
            $global:InputWorkbook.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($global:InputWorkbook) | Out-Null
        }
        
        if ($global:OutputWorkbook) {
            $global:OutputWorkbook.Close($true)  # Save the output workbook
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($global:OutputWorkbook) | Out-Null
        }
        
        # Quit Excel application
        if ($global:ExcelApp) {
            $global:ExcelApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($global:ExcelApp) | Out-Null
        }
        
        # Force garbage collection
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Host "Excel COM cleanup completed" -ForegroundColor Green
    }
    catch {
        Write-Warning "Error during Excel cleanup: $($_.Exception.Message)"
    }
}

# Function to read input values from Excel using COM API
function Read-ExcelInputUsingCOM {
    param(
        [string]$InputFilePath,
        [string]$WorksheetName = "Sheet1",
        [string]$InputColumn = "A",
        [int]$StartRow = 2  # Start from row 2 (assuming row 1 has headers)
    )
    
    try {
        Write-Host "Reading input from Excel file using COM API: $InputFilePath" -ForegroundColor Green
        
        # Check if file exists
        if (-not (Test-Path $InputFilePath)) {
            throw "Input Excel file not found: $InputFilePath"
        }
        
        # Open the input workbook
        $global:InputWorkbook = $global:ExcelApp.Workbooks.Open($InputFilePath)
        $worksheet = $global:InputWorkbook.Worksheets.Item($WorksheetName)
        
        # Find the last used row in the input column
        $lastRow = $worksheet.Cells($worksheet.Rows.Count, $InputColumn).End(-4162).Row  # xlUp = -4162
        
        Write-Host "Found data from row $StartRow to row $lastRow" -ForegroundColor Cyan
        
        # Read values from the specified column
        $inputValues = @()
        for ($row = $StartRow; $row -le $lastRow; $row++) {
            $cellValue = $worksheet.Cells.Item($row, $InputColumn).Value2
            if ($cellValue -and $cellValue.ToString().Trim() -ne "") {
                $inputValues += @{
                    Value = $cellValue.ToString().Trim()
                    RowNumber = $row
                    OriginalPosition = "$InputColumn$row"
                }
            }
        }
        
        Write-Host "Successfully read $($inputValues.Count) input values using Excel COM API" -ForegroundColor Green
        return $inputValues
    }
    catch {
        Write-Error "Error reading Excel file using COM API: $($_.Exception.Message)"
        return $null
    }
}

# Function to perform Excel API operations (Reading operations)
function Invoke-ExcelReadOperation {
    param(
        [object]$InputData,
        [string]$OperationType = "CellRead"
    )
    
    try {
        Write-Host "Performing Excel Read Operation: $OperationType for value: $($InputData.Value)" -ForegroundColor Blue
        
        $result = @{
            Success = $true
            OperationType = $OperationType
            InputValue = $InputData.Value
            InputPosition = $InputData.OriginalPosition
            Timestamp = Get-Date
        }
        
        switch ($OperationType) {
            "CellRead" {
                # Read additional cell properties
                $worksheet = $global:InputWorkbook.Worksheets.Item(1)
                $cell = $worksheet.Cells.Item($InputData.RowNumber, 1)  # Column A
                
                $result.CellProperties = @{
                    Value = $cell.Value2
                    Formula = $cell.Formula
                    Address = $cell.Address
                    Row = $cell.Row
                    Column = $cell.Column
                    HasFormula = $cell.HasFormula
                    NumberFormat = $cell.NumberFormat
                }
            }
            
            "RangeRead" {
                # Read a range of cells around the input
                $worksheet = $global:InputWorkbook.Worksheets.Item(1)
                $startRow = [Math]::Max(1, $InputData.RowNumber - 1)
                $endRow = [Math]::Min($worksheet.UsedRange.Rows.Count, $InputData.RowNumber + 1)
                
                $range = $worksheet.Range("A$startRow", "C$endRow")
                $rangeData = @()
                
                for ($r = 1; $r -le $range.Rows.Count; $r++) {
                    $rowData = @()
                    for ($c = 1; $c -le $range.Columns.Count; $c++) {
                        $rowData += $range.Cells.Item($r, $c).Value2
                    }
                    $rangeData += ,@($rowData)
                }
                
                $result.RangeData = $rangeData
            }
            
            "WorksheetInfo" {
                # Get worksheet information
                $worksheet = $global:InputWorkbook.Worksheets.Item(1)
                
                $result.WorksheetInfo = @{
                    Name = $worksheet.Name
                    UsedRangeAddress = $worksheet.UsedRange.Address
                    RowCount = $worksheet.UsedRange.Rows.Count
                    ColumnCount = $worksheet.UsedRange.Columns.Count
                    LastCell = $worksheet.UsedRange.Cells($worksheet.UsedRange.Cells.Count).Address
                }
            }
        }
        
        Start-Sleep -Milliseconds 100  # Small delay to simulate processing
        return $result
    }
    catch {
        Write-Warning "Excel Read Operation '$OperationType' failed for $($InputData.Value): $($_.Exception.Message)"
        return @{
            Success = $false
            Error = $_.Exception.Message
            OperationType = $OperationType
            InputValue = $InputData.Value
            Timestamp = Get-Date
        }
    }
}

# Function to perform Excel API write operations
function Invoke-ExcelWriteOperation {
    param(
        [object]$ProcessedData,
        [string]$OutputFilePath,
        [string]$WorksheetName = "Results",
        [switch]$IterativeMode
    )
    
    try {
        Write-Host "Performing Excel Write Operation to: $OutputFilePath" -ForegroundColor Blue
        
        # Check if we need to create a new workbook or open existing one
        if ($IterativeMode -and (Test-Path $OutputFilePath)) {
            Write-Host "Opening existing output file for iterative mode" -ForegroundColor Cyan
            $global:OutputWorkbook = $global:ExcelApp.Workbooks.Open($OutputFilePath)
        } else {
            Write-Host "Creating new output workbook" -ForegroundColor Cyan
            $global:OutputWorkbook = $global:ExcelApp.Workbooks.Add()
        }
        
        # Get or create the target worksheet
        try {
            $outputWorksheet = $global:OutputWorkbook.Worksheets.Item($WorksheetName)
        }
        catch {
            $outputWorksheet = $global:OutputWorkbook.Worksheets.Add()
            $outputWorksheet.Name = $WorksheetName
        }
        
        # Find the next available row for writing
        $nextRow = 1
        if ($IterativeMode -and $outputWorksheet.UsedRange) {
            $nextRow = $outputWorksheet.UsedRange.Rows.Count + 1
        }
        
        # Write headers if this is the first write
        if ($nextRow -eq 1) {
            $headers = @("Timestamp", "InputValue", "InputPosition", "ReadOperation1_Status", "ReadOperation1_Data", 
                        "ReadOperation2_Status", "ReadOperation2_Data", "ReadOperation3_Status", "ReadOperation3_Data", 
                        "ProcessingStatus", "RowNumber", "BatchID")
            
            for ($col = 1; $col -le $headers.Count; $col++) {
                $outputWorksheet.Cells.Item(1, $col).Value2 = $headers[$col - 1]
                $outputWorksheet.Cells.Item(1, $col).Font.Bold = $true
                $outputWorksheet.Cells.Item(1, $col).Interior.Color = 15773696  # Light blue background
            }
            $nextRow = 2
        }
        
        # Write the processed data
        $currentRow = $nextRow
        foreach ($item in $ProcessedData) {
            $outputWorksheet.Cells.Item($currentRow, 1).Value2 = $item.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")
            $outputWorksheet.Cells.Item($currentRow, 2).Value2 = $item.InputValue
            $outputWorksheet.Cells.Item($currentRow, 3).Value2 = $item.InputPosition
            
            # Read Operation 1 results
            $outputWorksheet.Cells.Item($currentRow, 4).Value2 = if ($item.ReadOp1) { $item.ReadOp1.Success.ToString() } else { "N/A" }
            $outputWorksheet.Cells.Item($currentRow, 5).Value2 = if ($item.ReadOp1 -and $item.ReadOp1.Success) { "Operation completed successfully" } else { if ($item.ReadOp1) { $item.ReadOp1.Error } else { "Not executed" } }
            
            # Read Operation 2 results
            $outputWorksheet.Cells.Item($currentRow, 6).Value2 = if ($item.ReadOp2) { $item.ReadOp2.Success.ToString() } else { "N/A" }
            $outputWorksheet.Cells.Item($currentRow, 7).Value2 = if ($item.ReadOp2 -and $item.ReadOp2.Success) { "Range read completed" } else { if ($item.ReadOp2) { $item.ReadOp2.Error } else { "Not executed" } }
            
            # Read Operation 3 results
            $outputWorksheet.Cells.Item($currentRow, 8).Value2 = if ($item.ReadOp3) { $item.ReadOp3.Success.ToString() } else { "N/A" }
            $outputWorksheet.Cells.Item($currentRow, 9).Value2 = if ($item.ReadOp3 -and $item.ReadOp3.Success) { "Worksheet info retrieved" } else { if ($item.ReadOp3) { $item.ReadOp3.Error } else { "Not executed" } }
            
            $outputWorksheet.Cells.Item($currentRow, 10).Value2 = $item.OverallStatus
            $outputWorksheet.Cells.Item($currentRow, 11).Value2 = $item.OriginalRowNumber
            $outputWorksheet.Cells.Item($currentRow, 12).Value2 = $item.BatchID
            
            $currentRow++
        }
        
        # Auto-fit columns
        $outputWorksheet.Columns.AutoFit() | Out-Null
        
        # Add borders to the data
        $dataRange = $outputWorksheet.Range($outputWorksheet.Cells.Item(1, 1), $outputWorksheet.Cells.Item($currentRow - 1, 12))
        $dataRange.Borders.LineStyle = 1  # xlContinuous
        $dataRange.Borders.Weight = 2     # xlThin
        
        # Save the workbook
        if (-not $IterativeMode -or -not (Test-Path $OutputFilePath)) {
            $global:OutputWorkbook.SaveAs($OutputFilePath)
        } else {
            $global:OutputWorkbook.Save()
        }
        
        Write-Host "Successfully wrote $($ProcessedData.Count) records to Excel using COM API" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Excel Write Operation failed: $($_.Exception.Message)"
        return $false
    }
}

# Function to process Excel operations in sequence
function Process-ExcelOperations {
    param(
        [object]$InputData,
        [string]$BatchID
    )
    
    Write-Host "`n--- Processing Excel Operations for: $($InputData.Value) ---" -ForegroundColor Magenta
    
    # Initialize result object
    $result = @{
        InputValue = $InputData.Value
        InputPosition = $InputData.OriginalPosition
        OriginalRowNumber = $InputData.RowNumber
        ReadOp1 = $null
        ReadOp2 = $null
        ReadOp3 = $null
        Timestamp = Get-Date
        BatchID = $BatchID
        OverallStatus = "Processing"
    }
    
    # Excel Read Operation 1: Cell Read
    Write-Host "Executing Excel Read Operation 1: Cell Properties" -ForegroundColor Blue
    $result.ReadOp1 = Invoke-ExcelReadOperation -InputData $InputData -OperationType "CellRead"
    
    if ($result.ReadOp1.Success) {
        # Excel Read Operation 2: Range Read
        Write-Host "Executing Excel Read Operation 2: Range Read" -ForegroundColor Blue
        $result.ReadOp2 = Invoke-ExcelReadOperation -InputData $InputData -OperationType "RangeRead"
        
        if ($result.ReadOp2.Success) {
            # Excel Read Operation 3: Worksheet Info
            Write-Host "Executing Excel Read Operation 3: Worksheet Info" -ForegroundColor Blue
            $result.ReadOp3 = Invoke-ExcelReadOperation -InputData $InputData -OperationType "WorksheetInfo"
            
            if ($result.ReadOp3.Success) {
                $result.OverallStatus = "All Operations Completed Successfully"
                Write-Host "All Excel operations completed successfully for: $($InputData.Value)" -ForegroundColor Green
            } else {
                $result.OverallStatus = "Failed at Excel Operation 3"
            }
        } else {
            $result.OverallStatus = "Failed at Excel Operation 2"
        }
    } else {
        $result.OverallStatus = "Failed at Excel Operation 1"
    }
    
    return $result
}

# Main execution function
function Start-ExcelCOMProcessing {
    param(
        [string]$InputExcelFile = ".\input.xlsx",
        [string]$OutputExcelFile = ".\output_results.xlsx",
        [string]$InputWorksheet = "Sheet1",
        [string]$InputColumn = "A",
        [int]$StartRow = 2,
        [switch]$IterativeMode,
        [int]$BatchSize = 10
    )
    
    Write-Host "=== Excel COM API Processing Script Started ===" -ForegroundColor Cyan
    Write-Host "Input File: $InputExcelFile" -ForegroundColor White
    Write-Host "Output File: $OutputExcelFile" -ForegroundColor White
    Write-Host "Iterative Mode: $IterativeMode" -ForegroundColor White
    Write-Host "Batch Size: $BatchSize" -ForegroundColor White
    
    $allResults = @()
    
    try {
        # Initialize Excel COM Application
        if (-not (Initialize-ExcelApplication)) {
            throw "Failed to initialize Excel COM Application"
        }
        
        # Read input values from Excel using COM API
        $inputData = Read-ExcelInputUsingCOM -InputFilePath $InputExcelFile -WorksheetName $InputWorksheet -InputColumn $InputColumn -StartRow $StartRow
        
        if (-not $inputData -or $inputData.Count -eq 0) {
            throw "No input values found in the Excel file"
        }
        
        # Process inputs in batches
        $totalInputs = $inputData.Count
        $processedCount = 0
        $batchID = (Get-Date).ToString("yyyyMMdd_HHmmss")
        
        for ($i = 0; $i -lt $totalInputs; $i += $BatchSize) {
            $batchEnd = [Math]::Min($i + $BatchSize - 1, $totalInputs - 1)
            $currentBatch = $inputData[$i..$batchEnd]
            
            Write-Host "`n--- Processing Batch: $($i + 1) to $($batchEnd + 1) of $totalInputs ---" -ForegroundColor Yellow
            
            $batchResults = @()
            foreach ($inputItem in $currentBatch) {
                # Process each input through Excel COM operations
                $result = Process-ExcelOperations -InputData $inputItem -BatchID $batchID
                $batchResults += $result
                $processedCount++
                
                # Progress indicator
                $percentComplete = [Math]::Round(($processedCount / $totalInputs) * 100, 2)
                Write-Progress -Activity "Processing Excel COM Operations" -Status "Processed $processedCount of $totalInputs inputs" -PercentComplete $percentComplete
            }
            
            # Add batch results to all results
            $allResults += $batchResults
            
            # Write results to Excel using COM API
            $writeSuccess = Invoke-ExcelWriteOperation -ProcessedData $batchResults -OutputFilePath $OutputExcelFile -WorksheetName "Results" -IterativeMode:$IterativeMode
            
            if ($writeSuccess) {
                Write-Host "Batch results written to Excel successfully" -ForegroundColor Green
            } else {
                Write-Warning "Failed to write batch results to Excel"
            }
        }
        
        # Summary report
        Write-Host "`n=== Processing Summary ===" -ForegroundColor Cyan
        Write-Host "Total Inputs Processed: $($allResults.Count)" -ForegroundColor White
        Write-Host "Successful: $($allResults | Where-Object { $_.OverallStatus -eq 'All Operations Completed Successfully' } | Measure-Object | Select-Object -ExpandProperty Count)" -ForegroundColor Green
        Write-Host "Failed: $($allResults | Where-Object { $_.OverallStatus -ne 'All Operations Completed Successfully' } | Measure-Object | Select-Object -ExpandProperty Count)" -ForegroundColor Red
        Write-Host "Output saved to: $OutputExcelFile" -ForegroundColor White
        
        return $allResults
    }
    catch {
        Write-Error "Script execution failed: $($_.Exception.Message)"
        return $null
    }
    finally {
        Write-Progress -Activity "Processing Excel COM Operations" -Completed
        Close-ExcelApplication
        Write-Host "`n=== Excel COM API Processing Script Completed ===" -ForegroundColor Cyan
    }
}

# Function to create sample input Excel file for testing
function New-SampleExcelFile {
    param([string]$FilePath = ".\sample_input.xlsx")
    
    try {
        # Initialize Excel if not already done
        if (-not $global:ExcelApp) {
            Initialize-ExcelApplication
        }
        
        # Create new workbook
        $sampleWorkbook = $global:ExcelApp.Workbooks.Add()
        $worksheet = $sampleWorkbook.Worksheets.Item(1)
        
        # Add headers
        $worksheet.Cells.Item(1, 1).Value2 = "InputValue"
        $worksheet.Cells.Item(1, 2).Value2 = "Description"
        $worksheet.Cells.Item(1, 1).Font.Bold = $true
        $worksheet.Cells.Item(1, 2).Font.Bold = $true
        
        # Add sample data
        $sampleData = @(
            @("Sample Data 1", "First test input for Excel COM processing")
            @("Test Value 2", "Second sample input with different content")
            @("Excel API 3", "Third test case for API operations")
            @("Demo Item 4", "Fourth demonstration input")
            @("Processing Test 5", "Fifth input for batch testing")
            @("COM Test 6", "Sixth input to test COM operations")
            @("Final Sample 7", "Last sample input for comprehensive testing")
        )
        
        for ($row = 0; $row -lt $sampleData.Count; $row++) {
            $worksheet.Cells.Item($row + 2, 1).Value2 = $sampleData[$row][0]
            $worksheet.Cells.Item($row + 2, 2).Value2 = $sampleData[$row][1]
        }
        
        # Auto-fit columns
        $worksheet.Columns.AutoFit() | Out-Null
        
        # Save the file
        $sampleWorkbook.SaveAs($FilePath)
        $sampleWorkbook.Close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sampleWorkbook) | Out-Null
        
        Write-Host "Sample Excel input file created: $FilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to create sample Excel file: $($_.Exception.Message)"
    }
}

# Usage examples
function Show-UsageExamples {
    Write-Host "`n=== Excel COM API Usage Examples ===" -ForegroundColor Green
    
    Write-Host "`n1. Basic usage:" -ForegroundColor Yellow
    Write-Host 'Start-ExcelCOMProcessing -InputExcelFile ".\input.xlsx" -OutputExcelFile ".\results.xlsx"' -ForegroundColor White
    
    Write-Host "`n2. Iterative mode (appends to existing output file):" -ForegroundColor Yellow
    Write-Host 'Start-ExcelCOMProcessing -InputExcelFile ".\input.xlsx" -OutputExcelFile ".\results.xlsx" -IterativeMode' -ForegroundColor White
    
    Write-Host "`n3. Custom configuration:" -ForegroundColor Yellow
    Write-Host 'Start-ExcelCOMProcessing -InputExcelFile ".\data.xlsx" -OutputExcelFile ".\output.xlsx" -InputWorksheet "Data" -InputColumn "B" -StartRow 3 -BatchSize 5 -IterativeMode' -ForegroundColor White
    
    Write-Host "`n4. Create sample input file:" -ForegroundColor Yellow
    Write-Host 'New-SampleExcelFile -FilePath ".\test_input.xlsx"' -ForegroundColor White
}

# Display usage examples
Show-UsageExamples

# Uncomment the lines below for testing:
# New-SampleExcelFile -FilePath ".\sample_input.xlsx"
# Start-ExcelCOMProcessing -InputExcelFile ".\sample_input.xlsx" -OutputExcelFile ".\com_results.xlsx" -IterativeMode