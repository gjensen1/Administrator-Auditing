<# *******************************************************************************************************************
*******************************************************************************************************************
Purpose of Script: Part of the Administrators Auditing Set of Scripts

        The purpose of this script is to collect all members of the Local Administrators Group
        from the server and write the Raw Data to c:\temp.  CSV file generated in this folder will need
        to be email to the requestor.
   
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

# ----------------
# Define Variables
# ----------------
$Computer = "localhost"
$ComputerName = $env:COMPUTERNAME
$Date = Get-Date -Format "yyyy_MM_dd"
$LocalGroupName = "Administrators"
$OutputDir = "c:\temp"



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

#*********************************************************
# Retrieve Initial Members from Local Administrators Group
#*********************************************************
Function Get-LocalAdmin-Members {

# Initialize Output File

$OutputFile = Join-Path $OutputDir "$ComputerName-$Date-Initial.csv"
Write-Verbose "Script will write the output to $OutputFile folder"
Add-Content -Path $OutPutFile -Value "ComputerName, LocalGroupName, Status, MemberType, MemberDomain, MemberName"

$group = [ADSI]"WinNT://$Computer/$LocalGroupName"
$members = @($group.Invoke("Members"))
Write-Verbose "Successfully queries the members of $computer"
if(!$members) {
    Add-Content -Path $OutputFile -Value "$ComputerName,$LocalGroupName,NoMembersFound"
    Write-Verbose "No members found in the group"
    continue
    } else {
        foreach($member in $members) {
                try {
                    $MemberName = $member.GetType().Invokemember("Name","GetProperty",$null,$member,$null)
                    $MemberType = $member.GetType().Invokemember("Class","GetProperty",$null,$member,$null)
                    $MemberPath = $member.GetType().Invokemember("ADSPath","GetProperty",$null,$member,$null)
                    $MemberDomain = $null
                    if($MemberPath -match "^Winnt\:\/\/(?<domainName>\S+)\/(?<CompName>\S+)\/") {
                        if($MemberType -eq "User") {
                            $MemberType = "LocalUser"
                        } elseif($MemberType -eq "Group"){
                            $MemberType = "LocalGroup"
                        }
                        $MemberDomain = $matches["CompName"]
 
                    } elseif($MemberPath -match "^WinNT\:\/\/(?<domainname>\S+)/") {
                        if($MemberType -eq "User") {
                            $MemberType = "DomainUser"
                        } elseif($MemberType -eq "Group"){
                            $MemberType = "DomainGroup"
                        }
                        $MemberDomain = $matches["domainname"]
 
                    } else {
                        $MemberType = "Unknown"
                        $MemberDomain = "Unknown"
                        }
                Add-Content -Path $OutPutFile -Value "$ComputerName, $LocalGroupName, SUCCESS, $MemberType, $MemberDomain, $MemberName"
                } catch {
                    Write-Verbose "failed to query details of a member. Details $_"
                    Add-Content -Path $OutputFile -Value "$ComputerName,,FailedQueryMember"
                }
            }
        }
}
#***************************************************************
# EndFunction Get-LocalAdmin-Members
#***************************************************************

#***************
# Execute Script
#***************
Verify-Folders
Get-LocalAdmin-Members
