# Login to Azure account
Connect-AzAccount

# Path to input and output Excel files
$inputFile = "C:\Scripts\DataDiskDetails\VMList.xlsx"
$outputFile = "C:\Scripts\DataDiskDetails\VMDataDisksOutput.xlsx"

# Create Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($inputFile)
$worksheet = $workbook.Sheets.Item(1)

# Get used range to determine number of rows
$usedRange = $worksheet.UsedRange
$rowCount = $usedRange.Rows.Count

# Create array to store output objects
$outputObjects = @()

# Loop starting from row 2 assuming row 1 has headers
for ($row = 2; $row -le $rowCount; $row++) {
    $vmName = $worksheet.Cells.Item($row, 1).Text
    $subscription = $worksheet.Cells.Item($row, 2).Text
    $resourceGroup = $worksheet.Cells.Item($row, 3).Text

    # Set Azure subscription context
    Set-AzContext -Subscription $subscription -ErrorAction Stop

    # Get VM object
    $vm = Get-AzVM -Name $vmName -ResourceGroupName $resourceGroup

    $dataDisks = $vm.StorageProfile.DataDisks
    $numDisks = $dataDisks.Count
    $totalSizeGB = ($dataDisks | Measure-Object -Property DiskSizeGB -Sum).Sum

    # Store results in a PS object
    $outputObjects += [PSCustomObject]@{
        VMName = $vmName
        Subscription = $subscription
        ResourceGroup = $resourceGroup
        NumberOfDataDisks = $numDisks
        TotalSizeGB = $totalSizeGB
    }
}

# Close the input workbook
$workbook.Close($false)
$excel.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

# Create new Excel application for output
$excelOut = New-Object -ComObject Excel.Application
$excelOut.Visible = $false
$workbookOut = $excelOut.Workbooks.Add()
$worksheetOut = $workbookOut.Sheets.Item(1)

# Write headers
$headers = @("VMName", "Subscription", "ResourceGroup", "NumberOfDataDisks", "TotalSizeGB")
for ($i = 0; $i -lt $headers.Count; $i++) {
    $worksheetOut.Cells.Item(1, $i + 1) = $headers[$i]
}

# Write data rows
$rowIndex = 2
foreach ($obj in $outputObjects) {
    $worksheetOut.Cells.Item($rowIndex, 1) = $obj.VMName
    $worksheetOut.Cells.Item($rowIndex, 2) = $obj.Subscription
    $worksheetOut.Cells.Item($rowIndex, 3) = $obj.ResourceGroup
    $worksheetOut.Cells.Item($rowIndex, 4) = $obj.NumberOfDataDisks
    $worksheetOut.Cells.Item($rowIndex, 5) = $obj.TotalSizeGB
    $rowIndex++
}

# Save output workbook
$workbookOut.SaveAs($outputFile)
$workbookOut.Close($true)
$excelOut.Quit()

# Release COM objects for output
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheetOut) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookOut) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelOut) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Output "Data processed and saved to $outputFile"
