# Gets information about different accounts and outputs then into an Excel file

param(
	[string]$search = "sum11*",
	[string]$sheetname = [string]("Report")
)

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
$objWorksheet.cells.item(1,5) = "(Sortable)"
$objWorksheet.cells.item(1,6) = "Item Count"
$objWorksheet.cells.item(1,7) = "Total Item Size"
$objWorksheet.cells.item(1,8) = "Total Deleted Item Size"
$objWorksheet.cells.item(1,9) = "Time Zone"
$objWorksheet.cells.item(1,10) = "Language"
$objWorksheet.cells.item(1,11) = "Active Sync Device"
$objWorksheet.cells.item(1,11) = "First Sync"
$objWorksheet.cells.item(1,11) = "Last Sync"

Write-Progress -Activity "Generating list" -Status "Initializing Excel Document" -PercentComplete (4)
#Stylize the titles
$selection = $objWorksheet.Range("A1:T1")
$selection.Interior.ColorIndex = 15
$selection.Borders.LineStyle = 1
$selection.Style = "Title"


Write-Progress -Activity "Generating list" -Status "Gathering Mailboxes" -PercentComplete (5)
$row = 2
$count = 0
$boxes = get-mailbox $search
$total = $boxes.Count

#Process mailboxes
Write-Progress -Activity "Generating list" -Status "Writing Mailboxes" -PercentComplete (10)
$boxes | ForEach{
    Write-Progress -Activity "Generating list" -Status "Processing $_" -CurrentOperation "Mailbox Statistics" -PercentComplete (10+(($count/$total)*85))
    $stats = Get-MailboxStatistics $_.name
    Write-Progress -Activity "Generating list" -Status "Processing $_" -CurrentOperation "Regional Settings" -PercentComplete (10+(($count/$total)*85))
    $regional = Get-MailboxRegionalConfiguration $_.name
    Write-Progress -Activity "Generating list" -Status "Processing $_" -CurrentOperation "Device Configuration" -PercentComplete (10+(($count/$total)*85))
    $sync = @()
    $sync = Get-ActiveSyncDeviceStatistics -Mailbox $_.primarysmtpAddress
    
    Write-Progress -Activity "Generating list" -Status "Writting $_" -PercentComplete (10+(($count/$total)*85))
    $objWorksheet.Cells.Item($row,1) = $_.displayname
    $objWorksheet.Cells.Item($row,2) = $_.name
    $objWorksheet.Cells.Item($row,3) = $_.PrimarySMTPAddress
    $objWorksheet.cells.item($row,4) = $stats.LastLogonTime
    if($stats.LastLogonTime.Month -ne $null) {$objWorksheet.cells.item($row,5) = [string]$stats.LastLogonTime.Month + '.' + [string]$stats.LastLogonTime.Day}
    $objWorksheet.cells.item($row,6) = $stats.ItemCount
    $objWorksheet.cells.item($row,7) = [string]$stats.TotalItemSize
    $objWorksheet.cells.item($row,8) = [string]$stats.TotalDeletedItemSize
    $objWorksheet.cells.item($row,9) = $regional.TimeZone
    $objWorksheet.cells.item($row,10) = $regional.Language
    $col = 11
    foreach($dev in $sync){
        $objWorksheet.cells.item($row,$col) = $dev.DeviceFriendlyName
        $col++
        $objWorksheet.cells.item($row,$col) = $dev.FirstSyncTime
        $col++
        $objWorksheet.cells.item($row,$col) = $dev.LastSyncTime
        $col++
    }
    $row++
}

Write-Progress -Activity "Generating list" -Status 'Processing $_.name' -PercentComplete (97)
$objRange = $objWorksheet.UsedRange 
$objRange.EntireColumn.Autofit()
Write-Progress -Activity "Generating list" -Status 'Opening Program' -PercentComplete (100)
$objExcel.visible = $True
