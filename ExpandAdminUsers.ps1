<# *******************************************************************************************************************
*******************************************************************************************************************
Purpose of Script: Part of the Administrators Auditing Set of Scripts

        The purpose of this script is to collect all members of the Local Administrators Group
        from the server and write the Raw Data to a share for further processing.
         
*******************************************************************************************************************  
Authored Date:    Sep-16-2015
Original Author:  Graham Jensen
*************************
Development Environment:  Win2K12 R2

Additional Modules Required: None

Usage:   To be determined

OutPut:  To be determined

===================================================================================================================
Update Log:   Please use this section to document changes made to this script
===================================================================================================================
-----------------------------------------------------------------------------
Update <Date>
   Author:    <Name> 
   Description of Change:
      <Description>
-----------------------------------------------------------------------------
Update <Date>
   Author:    <Name>
   Description of Change:
      <Description>
-----------------------------------------------------------------------------
*******************************************************************************************************************


^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Operational Notes
-----------------
****Location of other scripts called:  ****
N/A

^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
****Scripts Called and passed parameters ****
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
+     Script Name                 +    Parameters Passed                     +
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
+                                 +                                          +
+---------------------------------+------------------------------------------+
+                                 +                                          +
+---------------------------------+------------------------------------------+
#>

#
# ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Import-Module ActiveDirectory

# ----------------
# Define Variables
# ----------------
#$Computer = "localhost"
#$Computer = $env:COMPUTERNAME
#$Date = "2015_09_30"
$Date = Get-Date -Format "yyyy_MM_dd"
$LocalGroupName = "Administrators"
$OutputDir = Join-Path "\\<networkShare>\ServerSpecialInfo\AdministratorAuditing" $Date
$Computer = $null
$TodaysMembers = Join-Path "\\<networkShare>" $Date

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
# Process Initial Members and Expand any Local or Domain Groups
#**************************************************************
Function Expand-LocalAdminMembers {

$InputFile = Join-Path $TodaysMembers "$Computer-$Date-Initial.csv"
$GroupList = Import-Csv $InputFile

# Initialize Output File
$ExpandedOutputFile = Join-Path $OutputDir "$Computer-$Date-LocalAdministrators.csv"
Write-Verbose "Script will write the output to $OutputFile folder"
Add-Content -Path $ExpandedOutPutFile -Value "Account, Type, Domain"

ForEach ($line in $GroupList) {
    $ComputerName = $line.ComputerName
    $LocalGroup = $line.LocalGroupName
    $Status = $line.Status
    $MemberType = $line.MemberType
    $MemberDomain = $line.MemberDomain
    $MemberName= $line.MemberName

    If ($MemberType -eq "DomainGroup") {
        Try {
            #Write-Host "Getting members of " $MemberDomain "\" $MemberName
            $RawGroupMembers = Get-ADGroupMember $MemberName -Server $MemberDomain -Recursive
            $GroupMembers = $RawGroupMembers | Select-Object -Property SamAccountName,objectClass,@{Name='distinguishedName'; Expression = {$_.DistinguishedName.Split(',')[-5].Split('=')[1].ToUpper()}}
            $GroupMembers | ConvertTo-Csv | Select -Skip 2| Add-Content -Path $ExpandedOutputFile
            } Catch {
                Write-Host "Failed to query details of $MemberName"
                Add-Content -Path $ExpandedOutputFile -Value "$MemberName, $MemberType, $MemberDomain, Failed to query group details on domain"
                }
        } Else {
            Add-Content -Path $ExpandedOutputFile -Value "$MemberName, $MemberType, $MemberDomain"
            }
    }
}
#********************************************************************
# EndFunction - Expand-LocalAdminMembers
#********************************************************************  

#*********
# Clean Up
#********* 
Function CleanUp {
$InputFile = Join-Path $TodaysMembers "$Computer-$Date-Initial.csv"
Remove-Item $InputFile
}
#********************
# EndFunction CleanUp
#******************** 

#***************
# Execute Script
#***************

Verify-Folders
$FileList = Get-ChildItem $TodaysMembers -Name

ForEach ($CurrentFile in $FileList){
    $SplitData = $CurrentFile.Split("-")
    $Computer = $SplitData[0]
    Write-Host "Processing $Computer"
    Expand-LocalAdminMembers
    #CleanUp
    }