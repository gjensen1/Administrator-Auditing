$Date = Get-Date -Format "yyyy_MM_dd"
$ChangeDir = Join-Path "\\<networkShare>\ServerSpecialInfo\AdministratorAuditing" "$Date\Changes"
$InputFile = Join-Path $ChangeDir "ChangeSummary.csv"

# **** Create and Prepare Excel Worksheet ****

#Create Excel COM Object
$excel = New-Object -ComObject excel.application
#Make Visible
$excel.Visible = $true
#Add A Workbook
$workbook = $excel.Workbooks.Add()
#Remove other Worksheets
#1..2 | ForEach {
#	$workbook.Worksheets.item(2).Delete()
#	}

#Connect to first worksheet to rename and make active
$ServerInfoSheet = $workbook.Worksheets.Item(1)
$ServerInfoSheet.Name = 'AuditAdminChg'
$ServerInfoSheet.Activate() | Out-Null

#Create a Title for the First Worksheet
$row = 1
$Column = 1
$ServerInfoSheet.Cells.Item($row,$column) = "Wintel Admin Change Report $Date"
$serverInfoSheet.Cells.Item($row,$column).HorizontalAlignment = -4108
$range = $ServerInfoSheet.Range("a1","e2")
$range.Style = 'Title'
$range = $ServerInfoSheet.Range("a1","e2")
$range.Merge() | Out-Null
$range.VerticalAlignment = -4108



#Increment row for next set of data
$row++;$row++

#Save the initial row so it can be used later to create a border
$initialRow =$row

#Get First Row of CSV file to use as Headers for the Excel File
$csvColumnName = (Get-Content $InputFile | Select-Object -First 1).Split(",")

#Build Header Row

$serverInfoSheet.Cells.Item($row,$column)= $csvColumnName[0].Trim($(34))
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= $csvColumnName[1].Trim($(34))
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= $csvColumnName[2].Trim($(34))
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= $csvColumnName[3].Trim($(34))
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= $csvColumnName[4].Trim($(34))
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++

#Increment Row and reset Column back to first column
$row++
$Column = 1

#Load the CSV File

$ServerList = Import-Csv $InputFile

ForEach ($server in $ServerList) {
	$ServerInfoSheet.Cells.Item($row,$Column) = $server.Server
	$Column++
	$ServerInfoSheet.Cells.Item($row,$Column) = $server.User
	$Column++
	$ServerInfoSheet.Cells.Item($row,$Column) = $server.Type
	$Column++
	$ServerInfoSheet.Cells.Item($row,$Column) = $server.Domain
	$Column++
	$ServerInfoSheet.Cells.Item($row,$Column) = $server.Change
	$Column++
	
	#Increment to next row and reset Column to 1
	$Column = 1
	$row++
	}
	
#Auto fit everything
$usedRange = $ServerInfoSheet.UsedRange
$usedRange.EntireColumn.AutoFit() | Out-Null

#Save the file
$workbook.SaveAs($ChangeDir+"\ChangeSummary"+$Date+".xlsx")

#Quit the application
$excel.Quit()


#Close leftover Excel Process
Get-Process -Name Excel | Stop-Process