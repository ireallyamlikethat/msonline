#requires -version 4 -modules ExchangeOnlineManagement,ImportExcel 

<#
.SYNOPSIS
  SOURCE - https://www.sharepointdiary.com/2018/04/get-office-365-groups-using-powershell.html ; https://www.sharepointdiary.com/2019/04/get-office-365-group-members-using-powershell.html
  Trimmed down, modified, wrapped up to export to multiple worksheets in one spreadsheet

.DESCRIPTION
  Get all groups from a tenant and list each members of the group on separate worksheets. 

.PARAMETER TenantURL
  URL of an M365 tenant, eg:   learnshrpt.sharepoint.com

.PARAMETER Path
  Folder where files will be exported. 

.PARAMETER combine
  Using this switch provides a MasterList worksheet for all user and group data

.INPUTS
  <Inputs if any, otherwise state None>

.OUTPUTS
  Excel spreadsheets with group and user information

.NOTES
  Version:        1.2
  Author:         Dave Nicholls
  Creation Date:  <Date>
  Purpose/Change: Add Combine switch, clean up notes, update for MFA use. 

.EXAMPLE
    Export data to separate worksheets
  
    .\Get-O365GroupMembersExport.ps1 -tenanturl https://learnshrpt.sharepoint.com/ -path c:\temp
    
.EXAMPLE
  Export data to only a MasterList worksheet

  .\Get-O365GroupMembersExport.ps1 -tenanturl learnshrpt.sharepoint.com -path c:\temp -combine
    
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
    $Path,
    [switch]$master
)

#Set Error Action
$ErrorActionPreference = 'Continue'

#import modules
import-module ExchangeOnlineManagement
import-module ImportExcel

#Connect to Exchange Online
Connect-ExchangeOnline -ShowBanner:$False
$tenantBasic = $($TenantURL.Replace('https://','').Replace('/',''))
$tenantFile = "$tenantBasic-groups-$(get-date -format MMddyyyy).xlsx"
$ReportFile = join-path $path $tenantFile

write-verbose "Save Report to - $reportfile" -verbose

#Get all Office 365 Group
$curTennant = $TenantURL
write-verbose "Get all groups for $curTennant " -verbose
$uGroups = Get-UnifiedGroup |sort-object Alias
write-verbose "- Found $($ugroups.count) groups"


foreach ($group in $ugroups){
  write-verbose "Check users in - $($group.displayname)" -verbose
    $curUsers = $group | Get-UnifiedGroupLinks -LinkType Member
    write-verbose "- Found $($curUsers.count) users" -verbose
 
    $curData = @(
       #if an owner is not a member include it anyway
       foreach ($owner in $group.ManagedBy|sort-object ){
          
            $curobject = [PSCustomObject]@{
              GroupName = $group.Displayname
              State = "OWNER"
              UserName = $owner
              UserDisplayName = (get-user $owner ).displayname
            }
               
            $curobject
        }
        
        #process group members
        foreach ($user in $curUsers){
            $curobject = [PSCustomObject]@{
                GroupName = $group.Displayname
                State = "MEMBER"
                UserName = $user.Name
                UserDisplayName = $user.Displayname
            }
            if ($user.name -in $group.managedby){
              $curobject.State  = "OWNER"
            }
            $curobject
        }
    )
    if ($master.ispresent){
        #export to masterlist only        
        write-verbose -verbose "Write to MasterList only"
        
      $curData | sort-object UserName | 
        export-excel -path $ReportFile -WorksheetName MasterList -FreezeTopRow -BoldTopRow -AutoSize -Append
    }else {
      #export to multiple pages
      write-verbose -verbose "Write to multiple pages"
      $curData | sort-object UserName| 
      export-excel -path $ReportFile -WorksheetName $group.Alias -FreezeTopRow -BoldTopRow -AutoSize
    }
    
}

write-verbose "REPORTING COMPLETE - DISCONNECTING"

#Disconnect Exchange Online
Disconnect-ExchangeOnline -Confirm:$False