# Ensure you are logged in to Azure
Connect-AzAccount

# Path to your Excel file
$excelFile = "C:\Scripts\VMInfo.xlsx"

# Launch Excel COM object
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($excelFile)
$sheet = $workbook.Sheets.Item(1)

# Start reading from the second row, assuming first row is headers
$row = 2
while ($sheet.Cells.Item($row, 1).Value() -ne $null) {
    $vmName = $sheet.Cells.Item($row, 1).Value()
    $resourceGroup = $sheet.Cells.Item($row, 2).Value()
    $subscription = $sheet.Cells.Item($row, 3).Value()

    # Set subscription context
    Set-AzContext -SubscriptionId $subscription

    # Start the VM
    try {
        Start-AzVM -Name $vmName -ResourceGroupName $resourceGroup
        Write-Host "Start command sent for $vmName in $resourceGroup (Subscription: $subscription)"
    } catch {
        Write-Warning "Failed to start $vmName in $resourceGroup (Subscription: $subscription): $_"
    }
    $row++
}

# Clean up Excel COM objects
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel
