################################################################################
#
# Script Checks the current connections and outputs them to an Excel Document
# Must be run on a computer with Excel 
# Used from a work station with a remote PowerShell connection to Exchange 2010
#
#################################################################################

#Finds current connections and creates an excel sheet
Write-Progress -Activity "Generating list" -Status "Initializing Excel Document" -PercentComplete (0)
#Open Excel and create a worksheet
$objExcel = new-Object -comobject Excel.Application   
$objWorkbook = $objExcel.Workbooks.Add()
$objWorksheet = $objWorkbook.Worksheets.Item(1)
$objWorksheet.Name = $sheetname

Write-Progress -Activity "Generating list" -Status "Initializing Excel Document" -PercentComplete (2)
#Format Sheet with name and headers
$objWorksheet.Cells.Item(1,1) = "Name"
$objWorksheet.Cells.Item(1,2) = "Username"
$objWorksheet.Cells.Item(1,3) = "Email"
$objWorksheet.cells.item(1,4) = "Last Accessed"
$objWorksheet.cells.item(1,5) = "Logon Time"
$objWorksheet.cells.item(1,6) = "Total Item Size"
$objWorksheet.cells.item(1,7) = "Total Deleted Item Size"
$objWorksheet.cells.item(1,8) = "Storage Limit`nStatus"
$objWorksheet.cells.item(1,9) = "Latency"
$objWorksheet.cells.item(1,10) = "Current Open`nFolders"
$objWorksheet.cells.item(1,11) = "Client Version"
$objWorksheet.cells.item(1,12) = "Client Mode"


Write-Progress -Activity "Generating list" -Status "Initializing Excel Document" -PercentComplete (4)
#Stylize the titles
$selection = $objWorksheet.Range("A1:T1")
$selection.Interior.ColorIndex = 15
$selection.Borders.LineStyle = 1
$selection.Style = "Title"


Write-Progress -Activity "Generating list" -Status "Gathering Connections" -PercentComplete (5)
$row = 2
$count = 0
$serverNames = @("gold","silver")
$connected = $serverNames | %{Get-LogonStatistics -server $_ | where{$_.LastAccessTime -gt (get-date).AddMinutes(-5)} | sort Username -Uniq}
$total = ($connected.Count) * 3

Write-Progress -Activity "Generating list" -Status "Writing Mailboxes" -PercentComplete (10)
$connected | ForEach{
    $count++
    Write-Progress -Activity "Generating list" -Status "Processing $_.UserName" -CurrentOperation "Mailbox Statistics" -PercentComplete (10+((($count*85)/($total*85))))
    $stats = Get-MailboxStatistics $_.Windows2000Account
    $count++

    Write-Progress -Activity "Generating list" -Status "Writting $_.UserName" -PercentComplete (10+((($count*85)/($total*85)))) 
    $objWorksheet.Cells.Item($row,1) = $_.UserName
    $objWorksheet.Cells.Item($row,2) = $_.Windows2000Account
    $objWorksheet.Cells.Item($row,3) = (get-mailbox $_.Windows2000Account).PrimarySMTPAddress
    $objWorksheet.cells.item($row,4) = $stats.LastAccessTime
    $objWorksheet.cells.item($row,5) = $_.LogonTime
    $objWorksheet.cells.item($row,6) = [string]$Stats.TotalItemSize
    $objWorksheet.cells.item($row,7) = [string]$Stats.TotalDeletedItemSize
    $objWorksheet.cells.item($row,8) = [string]$Stats.StorageLimitStatus
    $objWorksheet.cells.item($row,9) = $_.Latency
    $objWorksheet.cells.item($row,10) = $_.CurrentOpenFolders
    $objWorksheet.cells.item($row,11) = $_.ClientVersion
    $objWorksheet.cells.item($row,12) = $_.ClientMode
    $count ++
    $row++
}

Write-Progress -Activity "Generating list" -Status 'Processing $_.name' -PercentComplete (97)
$objRange = $objWorksheet.UsedRange 
$objRange.EntireColumn.Autofit()
Write-Progress -Activity "Generating list" -Status 'Opening Program' -PercentComplete (100)
$objExcel.visible = $True