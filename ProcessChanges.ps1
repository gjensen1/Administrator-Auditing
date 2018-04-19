
# ----------------
# Define Variables
# ----------------
#$Computer = "localhost"
#$Computer = $env:COMPUTERNAME
$Date = Get-Date -Format "yyyy_MM_dd"
$Yesterday = "2015_10_30"
#$Yesterday = (Get-Date).AddDays(-1).ToString('yyyy_MM_dd')
$OutputDir = Join-Path "\\<networkShare>\ServerSpecialInfo\AdministratorAuditing" "$Date\Changes"
$TodaysMembers = Join-Path "\\<networkShare>\ServerSpecialInfo\AdministratorAuditing" $Date
$YesterdaysMembers = Join-Path "\\<networkShare>\ServerSpecialInfo\AdministratorAuditing" $Yesterday
$ChangeDir = Join-Path "\\<networkShare>\ServerSpecialInfo\AdministratorAuditing" "$Date\Changes"
$ChangeSummary = Join-Path $ChangeDir "ChangeSummary.csv"
$FileList = $null
$CurrentFile = $null
$YesterdayFile = $null

#*************************************************
# Check for Folder Structure if not present create
#*************************************************
Function Verify-Folders {

    If (!(Test-Path $OutputDir)) {
        New-Item $OutputDir -type Directory
        }
}
#*************************************************
# EndFunction Verify-Folders
#*************************************************

#**************************************************************
# Process Members-Lists and Compare to Yesterdays Lists
#**************************************************************
Function Check-For-Change {
$Changes = Compare-Object -ReferenceObject (Get-Content (Join-Path $YesterdaysMembers $YesterdayFile)) -DifferenceObject (Get-Content (Join-Path $TodaysMembers $CurrentFile))

If ($Changes -ne $null) {
    $Changes | Export-Csv (Join-Path $OutputDir $CurrentFile) -NoTypeInformation
    }
}
#********************************************************************
# EndFunction - Check-For-Change
#********************************************************************  

#*******************************
# Initialize Change Summary File
#*******************************
Function Init-ChgSummaryFile {
    
    Add-Content -Path $ChangeSummary -Value "Server, User, Type, Domain, Change"
}
#**********************************
# EndFunction - Init-ChgSummaryFile
#**********************************

#****************
# Process Changes
#****************
Function Process-Changes {
    $ChangeList = Import-Csv (Join-Path $ChangeDir $CurrentFile)
    ForEach ($Line in $ChangeList){
        $Details = $Line.InputObject
        $UserDetails = $Details.split(",")
        $Username = $UserDetails[0].Trim("`"")
        $UserType = $UserDetails[1].Trim("`"")
        $UserDomain = $UserDetails[2].Trim("`"")
        $StateChange = $Line.SideIndicator
        If ($StateChange -eq "=>"){
            $State = "Added"
            } Else {
                $State = "Removed"
                }
        
        $SplitData = $CurrentFile.Split("-")
        $Computer = $SplitData[0]
        Add-Content -Path $ChangeSummary -Value "$Computer, $Username, $UserType, $UserDomain, $State"       
        }
}
#******************************
# EndFunction - Process-Changes
#******************************





#***************
# Execute Script
#***************

Verify-Folders

$FileList = Get-ChildItem $TodaysMembers -Name
ForEach ($CurrentFile in $FileList[1..($FileList.Length -1)]){
    $SplitData = $CurrentFile.Split("-")
    $Computer = $SplitData[0]
    $YesterdayFile = "$Computer-$Yesterday-LocalAdministrators.csv"
    
    Write-Host "Processing $CurrentFile"
    Check-For-Change
    }


Init-ChgSummaryFile

$FileList = Get-ChildItem $ChangeDir -Name
ForEach ($CurrentFile in $FileList){
    Process-Changes
    }
