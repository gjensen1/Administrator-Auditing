# ----------------
# Define Variables
# ----------------
$Date = Get-Date -Format "yyyy_MM_dd"
$InputDir = Join-Path "\\<networkShare>\ServerSpecialInfo\AdministratorAuditing" "$Date"
$ChangeDir = Join-Path "\\<networkShare>\ServerSpecialInfo\AdministratorAuditing" "$Date\Changes"
$ChangeFile = Join-Path $ChangeDir "ChangeSummary.csv"
$InputFile = $null
$Global:column = $null
$Global:row = $null
$Global:initialRow = $null
$Global:usedRange = $null
$Global:ServerInfoSheet = $null
$Global:range = $null
$Global:workbook = $null
$Global:Computer = $null
$Global:excel = $null
$Global:Count = $null
$Global:colSheets = $null

#***********************************
# Create and Prepare Excel Worksheet
#***********************************
function Initialize-Spreadsheet {

    #Create Excel COM Object
    $Global:excel = New-Object -ComObject excel.application
    #Make Visible
    $Global:excel.Visible = $true
    #Add A Workbook
    $Global:workbook = $Global:excel.Workbooks.Add()
   
    #Remove other Worksheets
    #1..2 | ForEach {
    #	$Global:workbook.Worksheets.item(2).Delete()
    #	}
    $Global:Count = 1
}
#**************************************
# End Function - Initialize-Spreadsheet
#**************************************

#**********************
# Build Title Worksheet
#**********************
function Build-Title {
    #Connect to worksheet to rename and make active
    $Global:ServerInfoSheet = $Global:workbook.Worksheets.Item($Global:workbook.Worksheets.count)
    $Global:ServerInfoSheet.Name = "TitleSheet"
    $Global:ServerInfoSheet.Activate() | Out-Null

    #Create Main Title for the First Worksheet
    $Global:row = 11
    $Global:column = 2
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column) = "SysTrust Admin Report"
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column).HorizontalAlignment = -4108
    $Global:range = $Global:ServerInfoSheet.Range("a1","o19")
    $Global:range.Style = 'Heading 1'
    $Global:range.Merge() | Out-Null
    $Global:range.VerticalAlignment = -4108

    #Create Date Title for the First Worksheet
    $Global:row = 20
    $Global:column = 1
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column) = "Date:  $Date"
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column).HorizontalAlignment = -4108
    $Global:range = $Global:ServerInfoSheet.Range("a20","o20")
    $Global:range.Style = 'Heading 2'
    $Global:range.Merge() | Out-Null
    $Global:range.VerticalAlignment = -4160
    $Global:range.HorizontalAlignment = -4131

    #Create Produced By Title for the First Worksheet
    $Global:row = 21
    $Global:column = 1
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column) = "Produced By: Automated Powershell Script"
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column).HorizontalAlignment = -4108
    $Global:range = $Global:ServerInfoSheet.Range("a21","o21")
    $Global:range.Style = 'Heading 2'
    $Global:range.Merge() | Out-Null
    $Global:range.VerticalAlignment = -4160
    $Global:range.HorizontalAlignment = -4131

    #Create Report Store Location Title for the First Worksheet
    $Global:row = 22
    $Global:column = 1
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column) = "Report Location: $InputDir\AdminReport.xlsx"
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column).HorizontalAlignment = -4108
    $Global:range = $Global:ServerInfoSheet.Range("a22","o22")
    $Global:range.Style = 'Heading 2'
    $Global:range.Merge() | Out-Null
    $Global:range.VerticalAlignment = -4160
    $Global:range.HorizontalAlignment = -4131
    
   
    #Auto fit everything
    $Global:usedRange = $Global:ServerInfoSheet.UsedRange
    $Global:usedRange.EntireColumn.AutoFit() | Out-Null
   
}
#***************************
# End Function - Build-Title
#***************************

#****************************
# Build Change Summary Report
#****************************
function Build-Access-Change-Report {
    #Connect to worksheet to rename and make active
    $Global:ServerInfoSheet = $Global:workbook.Worksheets.Item($Global:workbook.Worksheets.count)
    $Global:ServerInfoSheet.Name = "AuditAdminChg"
    $Global:ServerInfoSheet.Activate() | Out-Null

    #Create a Title for the Worksheet
    $Global:row = 1
    $Global:Column = 1
    $Global:ServerInfoSheet.Cells.Item($row,$column) = "SysTrust Admin Change Report"
    $Global:serverInfoSheet.Cells.Item($row,$column).HorizontalAlignment = -4108
    $Global:range = $ServerInfoSheet.Range("a1","e2")
    $Global:range.Style = 'Title'
    $Global:range = $ServerInfoSheet.Range("a1","e2")
    $Global:range.Merge() | Out-Null
    $Global:range.VerticalAlignment = -4108

    #Increment row for next set of data
    $Global:row++;$Global:row++

    #Save the initial row so it can be used later to create a border
    $Global:initialRow =$Global:row

    #Get First Row of CSV file to use as Headers for the Excel File
    $Global:csvColumnName = (Get-Content $ChangeFile | Select-Object -First 1).Split(",")

    #Build Header Row

    $Global:ServerInfoSheet.Cells.Item($row,$column)= $csvColumnName[0].Trim($(34))
    $Global:ServerInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $Global:ServerInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Global:Column++
    $Global:ServerInfoSheet.Cells.Item($row,$column)= $csvColumnName[1].Trim($(34))
    $Global:ServerInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $Global:ServerInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Global:Column++
    $Global:ServerInfoSheet.Cells.Item($row,$column)= $csvColumnName[2].Trim($(34))
    $Global:ServerInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $Global:ServerInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Global:Column++
    $Global:ServerInfoSheet.Cells.Item($row,$column)= $csvColumnName[3].Trim($(34))
    $Global:ServerInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $Global:ServerInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Global:Column++
    $Global:ServerInfoSheet.Cells.Item($row,$column)= $csvColumnName[4].Trim($(34))
    $Global:ServerInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
    $Global:ServerInfoSheet.Cells.Item($row,$column).Font.Bold=$True
    $Global:Column++

    #Increment Row and reset Column back to first column
    $Global:row++
    $Global:Column = 1

    #Load the CSV File
    $ServerList = Import-Csv $ChangeFile

    foreach ($server in $ServerList) {
	    $Global:ServerInfoSheet.Cells.Item($row,$Column) = $server.Server
	    $Global:Column++
	    $Global:ServerInfoSheet.Cells.Item($row,$Column) = $server.User
	    $Global:Column++
        if ($server.Type -eq "user") {
            $server.Type = "DomainUser"
            }
	    $Global:ServerInfoSheet.Cells.Item($row,$Column) = $server.Type
	    $Global:Column++
	    $Global:ServerInfoSheet.Cells.Item($row,$Column) = $server.Domain
	    $Global:Column++
	    $Global:ServerInfoSheet.Cells.Item($row,$Column) = $server.Change
	    $Global:Column++
	
	    #Increment to next row and reset Column to 1
	    $Global:Column = 1
	    $Global:row++
	    }
	
    #Auto fit everything
    $Global:usedRange = $Global:ServerInfoSheet.UsedRange
    $Global:usedRange.EntireColumn.AutoFit() | Out-Null

}

#********************************
# Add a Worksheet to the Workbook
#********************************
function Add-Worksheet {
    $Global:workbook.Worksheets.Add([System.Reflection.Missing]::Value,$Global:workbook.Worksheets.Item($Global:workbook.Worksheets.count)) | Out-Null
    }
#*****************************
# End Function - Add-Worksheet
#*****************************

#***********************************************
# Initialize Worksheet - Rename and create title
#***********************************************
function Initialize-Worksheet {
      
    #Connect to worksheet to rename and make active
    $Global:ServerInfoSheet = $Global:workbook.Worksheets.Item($Global:workbook.Worksheets.count)
    $Global:ServerInfoSheet.Name = $Global:Computer
    $Global:ServerInfoSheet.Activate() | Out-Null

    #Create a Title for the Worksheet
    $Global:row = 1
    $Global:column = 1
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column) = "$Global:Computer"
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column).HorizontalAlignment = -4108
    $Global:range = $Global:ServerInfoSheet.Range("a1","d2")
    $Global:range.Style = 'Title'
    $Global:range = $Global:ServerInfoSheet.Range("a1","d2")
    $Global:range.Merge() | Out-Null
    $Global:range.VerticalAlignment = -4108

    #Increment row for next set of data
    $Global:row++;$Global:row++


    $Global:Count++
    
    #Save the initial row so it can be used later to create a border
    $Global:initialRow =$Global:row
}
#**************************************
# End Function - Initialize-Spreadsheet
#************************************** 

#***************************
# Build Worksheet Header Row
#***************************
function Build-Header {

    #Get First Row of CSV file to use as Headers for the Excel File
    $csvColumnName = (Get-Content $InputFile | Select-Object -First 1).Split(",")
   
    #Build Header Row

    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column)= "ServerName"
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column).Interior.ColorIndex =48
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column).Font.Bold=$True
    $Global:column++
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column)= $csvColumnName[0].Trim($(34))
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column).Interior.ColorIndex =48
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column).Font.Bold=$True
    $Global:column++
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column)= $csvColumnName[1].Trim($(34))
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column).Interior.ColorIndex =48
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column).Font.Bold=$True
    $Global:column++
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column)= $csvColumnName[2].Trim($(34))
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column).Interior.ColorIndex =48
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column).Font.Bold=$True
    $Global:column++

    
    #Increment Row and reset Column back to first column
    $Global:row++
    $Global:column = 1
}
#****************************
# End Function - Build-Header
#**************************** 

#**************
# Load the Data
#**************
function Load-Data {

    $ServerList = Import-Csv $InputFile

    $Global:column = 1
    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column) = $Global:Computer
	$Global:column++

    foreach ($server in $ServerList) {
	    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column) = $server.Account
	    $Global:column++
        
        if ($server.Type -eq "user") {
            $server.Type = "DomainUser"
            }

	    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column) = $server.Type
	    $Global:column++
	    $Global:ServerInfoSheet.Cells.Item($Global:row,$Global:column) = $server.Domain
	    $Global:column++

	
	    #Increment to next row and reset Column to 2
	    $Global:column = 2
	    $Global:row++
	}
    #Auto fit everything
    $Global:usedRange = $Global:ServerInfoSheet.UsedRange
    $Global:usedRange.EntireColumn.AutoFit() | Out-Null
}
#*************************
# End Function - Load-Data
#*************************

function activate-first-sheet {
      
    #Connect to worksheet to rename and make active
    $Global:ServerInfoSheet = $Global:workbook.Worksheets.Item(1)
    $Global:ServerInfoSheet.Activate() | Out-Null
}

#*********************
# Clean up Spreadsheet
#*********************
function Clean-Up {	


    #Save the file
    $Global:workbook.SaveAs($InputDir+"AdminReport.xlsx")

    #Quit the application
    $Global:excel.Quit()


    #Close leftover Excel Process
    #Get-Process -Name Excel | Stop-Process
}
#************************
# End Function - Clean-Up
#************************

#***************
# Execute Script
#***************
Initialize-Spreadsheet
Build-Title
Add-Worksheet
Build-Access-Change-Report
$FileList = Get-ChildItem $InputDir -Name

$InputFile = Join-Path $InputDir $FileList[0]



foreach ($File in $FileList[1..($FileList.Length -1)]){
    $SplitData = $File.Split("-")
    $Global:Computer = $SplitData[0]
    $InputFile = Join-Path $InputDir $File
    
    Write-Host "Processing $File"
    Add-Worksheet
    Initialize-Worksheet
    Build-Header
    Load-Data
    }
activate-first-sheet
Clean-Up