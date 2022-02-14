#requires -version 4 -modules ExchangeOnlineManagement,ImportExcel 

<#
.SYNOPSIS
  SOURCE - https://www.sharepointdiary.com/2018/04/get-office-365-groups-using-powershell.html ; https://www.sharepointdiary.com/2019/04/get-office-365-group-members-using-powershell.html
  Trimmed down, modified, wrapped up to export to multiple worksheets in one spreadsheet

.DESCRIPTION

Get all groups from a tenant and list each members of the group on separate worksheets. 

.PARAMETER <Parameter_Name>
  <Brief description of parameter input required. Repeat this attribute if required>

.INPUTS
  <Inputs if any, otherwise state None>

.OUTPUTS
  <Outputs if any, otherwise state None>

.NOTES
  Version:        1.0
  Author:         <Name>
  Creation Date:  <Date>
  Purpose/Change: Initial script development

.EXAMPLE
    See if this works
  
    .\Get-O365GroupMembersExport.ps1 -tenanturl https://learnshrpt.sharepoint.com/ -path c:\temp
    
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------
[cmdletbinding()]
Param (  
    [Parameter(
        Mandatory = $true,
        Position = 0
    )]
    $TenantURL,
    [Parameter(Mandatory = $true)]
    $Path = "C:\Temp\",
    [Parameter(Mandatory = $true)]
    [pscredential]$Credential = (Get-Credential)
)

#Set Error Action
$ErrorActionPreference = 'Continue'

#import modules
import-module ExchangeOnlineManagement
import-module ImportExcel

#Connect to Exchange Online
Connect-ExchangeOnline -Credential $Credential -ShowBanner:$False
$tenantFile = "$($TenantURL.Replace('https://','').Replace('/',''))-groups.xlsx"
$ReportFile = join-path $path $tenantFile

write-verbose "Save Report to - $reportfile" -verbose

#Get all Office 365 Group
$curTennant = ($credential.username -replace "\w+@")
write-verbose "Get all groups for $curTennant " -verbose
$uGroups = Get-UnifiedGroup
write-verbose "- Found $($ugroups.count) groups"


foreach ($group in $ugroups){
  write-verbose "Check users in - $($group.displayname)" -verbose
    $curUsers = $group | Get-UnifiedGroupLinks -LinkType Member
    write-verbose "- Found $($curUsers.count) users" -verbose

    $curData = @(
        foreach ($user in $curUsers){
            $curobject = [PSCustomObject]@{
                GroupName = $group.Displayname
                #GroupAlias = $group.Alias
                #GroupType = $group.GroupType
                UserName = $user.Name
                UserDisplayName = $user.Displayname
                #UserTitle = $user.title 
            }
            $curobject
        }
    )
    $curData |export-excel -path $ReportFile -WorksheetName $group.Alias -FreezeTopRow -BoldTopRow -AutoSize
}

write-verbose "REPORTING COMPLETE - DISCONNECTING"

#Disconnect Exchange Online
Disconnect-ExchangeOnline -Confirm:$False