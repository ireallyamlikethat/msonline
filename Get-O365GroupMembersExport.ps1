#requires -version 4 -modules microsoft.graph,ImportExcel 

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

.PARAMETER Master
  Using this switch provides only a MasterList worksheet for all user and group data

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

  .\Get-O365GroupMembersExport.ps1 -tenanturl learnshrpt.sharepoint.com -path c:\temp -Master
    
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------
[cmdletbinding()]
Param (  
    [Parameter(Mandatory = $true,Position = 0)]
    $TenantURL,
    [Parameter(Mandatory = $true)]
    $Path,
    [switch]$Master
)

#Set Error Action
$ErrorActionPreference = 'Continue'

#import modules
import-module ImportExcel
import-module microsoft.graph

#Connect to Exchange Online
##onnect-ExchangeOnline -ShowBanner:$False

#connect to microsoft graph use web auth
$RequiredScopes = @("Directory.AccessAsUser.All", "Directory.ReadWrite.All","User.Read","Application.Read.All")
connect-mggraph -Scopes  $RequiredScopes

$tenantBasic = $($TenantURL.Replace('https://','').Replace('/',''))
$tenantFile = "$tenantBasic-groups-$(get-date -format MMddyyyy).xlsx"
$ReportFile = join-path $path $tenantFile

write-verbose "Save Report to - $reportfile" -verbose

#Get all Office 365 Group
$curTennant = $TenantURL
write-verbose "Get all groups for $curTennant " -verbose
$allGroups = get-mggroup |sort-object displayname
write-verbose "- Found $($allGroups.count) groups"


foreach ($group in $allGroups){
  write-verbose "Check users in - $($group.displayname)" -verbose
    ##$curUsers = $group | Get-UnifiedGroupLinks -LinkType Member
    $curUsers = Get-MgGroupMember -GroupId $group.id | foreach-object {Get-MgUser -UserId $_.id}    
    $curOwners = Get-MgGroupOwner -GroupId $group.id | foreach-object {Get-MgUser -UserId $_.id}
    write-verbose "- Found $($curUsers.count) users" -verbose
    write-verbose "- Found $($curOwners.count) owners" -verbose
 
    $curData = @(
        #groups w/out owners or users
        if ( ($null -eq $curusers) -and ($null -eq $curowners) ){
          $curobject = [PSCustomObject]@{
            GroupName = $group.Displayname
            State = "EMPTY_OWNER"
            UserName = "EMPTY"
            UserDisplayName = "NONE"
            MailEnabled = $group.MailEnabled
            SecurityEnabled = $group.SecurityEnabled
          }
          $curobject
        }
    
        #include owners, in case they are not also members
        foreach ($owner in $curOwners|sort-object ){            
              $curobject = [PSCustomObject]@{
                GroupName = $group.Displayname
                State = "OWNER"
                UserName = $owner.UserPrincipalName
                UserDisplayName = $owner.displayname
                MailEnabled = $group.MailEnabled
                SecurityEnabled = $group.SecurityEnabled
              }
              $curobject
          }       
        
        #process group members
        foreach ($user in $curUsers){
            $curobject = [PSCustomObject]@{
                GroupName = $group.Displayname
                State = "MEMBER"
                UserName = $user.UserPrincipalName
                UserDisplayName = $user.Displayname
                MailEnabled = $group.MailEnabled
                SecurityEnabled = $group.SecurityEnabled
            }
            if ($user.UserPrincipalName -in $group.UserPrincipalName){
              $curobject.State  = "OWNER"
            }
            $curobject
        }
    )
    if ($master.ispresent){
        #export to masterlist only        
        write-verbose -verbose "Write to worksheet MasterList only $ReportFile"
        
      $curData | sort-object UserName | 
        export-excel -path $ReportFile -WorksheetName MasterList -FreezeTopRow -BoldTopRow -AutoSize -Append
    } else {
      #export to multiple pages
      write-verbose -verbose "Write to worksheet $($group.Displayname) in $ReportFile"
      $curData | sort-object UserName| 
        export-excel -path $ReportFile -WorksheetName $group.Displayname -FreezeTopRow -BoldTopRow -AutoSize
    }
}

write-verbose "REPORTING COMPLETE - DISCONNECTING"

#disconnect msgraph
# 'Demo Mail Security Group' and a 'Demo Security Group' in our tenant.
# 'Demo Mail Security Empty Group' so i can see an owner of an empty group
# 'Demo Security Empty Group' so i can see an owner of an empty group
Disconnect-MgGraph