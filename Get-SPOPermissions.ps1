#requires -version 4 -modules pnp.powershell,ImportExcel 

<#
.SYNOPSIS
  SOURCE - https://www.sharepointdiary.com/2019/09/sharepoint-online-user-permissions-audit-report-using-pnp-powershell.html
  Trimmed down, modified, wrapped up to export to multiple worksheets in one spreadsheet
  Note - this calls get-pnppermissions.ps1 which is required for detail information on sites. 

.DESCRIPTION

USes to see each site, 
    the groups for that site, 
    each list or library in the site 
        and each group and permission level for those lists and libraries 
    exclude file/folder level detail

.PARAMETER TenantURL
  URL of an M365 tenant, eg:   learnshrpt.sharepoint.com

.PARAMETER Path
  Folder where files will be exported. 

.PARAMETER Master
  Using this switch provides only a MasterList worksheet for all user and group data

.INPUTS
  <Inputs if any, otherwise state None>

.OUTPUTS
  <Outputs if any, otherwise state None>

.NOTES
  Version:        1.2
  Author:         Dave Nicholls
  Creation Date:  <Date>
  Purpose/Change: Add master switch, clean up notes, update for MFA use. 

.EXAMPLE
    Export data to separate worksheets

    .\Get-SPOPermissions.ps1 -TenantURL https://learnshrpt.sharepoint.com/ -Path c:\temp
     
.EXAMPLE
    Export data to only a MasterList worksheet

    .\Get-SPOPermissions.ps1 -TenantURL https://learnshrpt.sharepoint.com/ -Path c:\temp -Master
    
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------
[cmdletbinding()]
Param (  
    [Parameter(
        Mandatory = $true,
        Position = 0
    )]
    $TenantURL = "https://learnshrpt.sharepoint.com",
    [Parameter(Mandatory = $true)]
    $Path,
    [switch]$Master
)

#Set Error Action
$ErrorActionPreference = 'Continue'

#Function to Get Permissions Applied on a particular Object, such as: Web, List, Folder or List Item
Function Get-PnPPermissions([Microsoft.SharePoint.Client.SecurableObject]$Object)
{
    write-verbose "-- Run Get-PnPPermissions against $($Object |out-string)"
    #Determine the type of the object
    Switch($Object.TypedObject.ToString())
    {
        "Microsoft.SharePoint.Client.Web"  { $ObjectType = "Site" ; $ObjectURL = $Object.URL; $ObjectTitle = $Object.Title }
        "Microsoft.SharePoint.Client.ListItem"
        { 
            If($Object.FileSystemObjectType -eq "Folder")
            {
                $ObjectType = "Folder"
                #Get the URL of the Folder 
                $Folder = Get-PnPProperty -ClientObject $Object -Property Folder
                $ObjectTitle = $Object.Folder.Name
                $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl,''),$Object.Folder.ServerRelativeUrl)
            }
            Else #File or List Item
            {
                #Get the URL of the Object
                Get-PnPProperty -ClientObject $Object -Property File, ParentList
                If($Null -ne $Object.File.Name)
                {
                    $ObjectType = "File"
                    $ObjectTitle = $Object.File.Name
                    $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl,''),$Object.File.ServerRelativeUrl)
                }
                else
                {
                    $ObjectType = "List Item"
                    $ObjectTitle = $Object["Title"]
                    #Get the URL of the List Item
                    $DefaultDisplayFormUrl = Get-PnPProperty -ClientObject $Object.ParentList -Property DefaultDisplayFormUrl                     
                    $ObjectURL = $("{0}{1}?ID={2}" -f $Web.Url.Replace($Web.ServerRelativeUrl,''), $DefaultDisplayFormUrl,$Object.ID)
                }
            }
        }
        Default
        { 
            $ObjectType = "List or Library"
            $ObjectTitle = $Object.Title
            #Get the URL of the List or Library
            $RootFolder = Get-PnPProperty -ClientObject $Object -Property RootFolder     
            $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl,''), $RootFolder.ServerRelativeUrl)
        }
    }
   
    #Get permissions assigned to the object
    Get-PnPProperty -ClientObject $Object -Property HasUniqueRoleAssignments, RoleAssignments
 
    #Check if Object has unique permissions
    $HasUniquePermissions = $Object.HasUniqueRoleAssignments
     
    #Loop through each permission assigned and extract details
    $PermissionCollection = @()
    Foreach($RoleAssignment in $Object.RoleAssignments)
    { 
        #Get the Permission Levels assigned and Member
        Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
 
        #Get the Principal Type: User, SP Group, AD Group
        $PermissionType = $RoleAssignment.Member.PrincipalType
    
        #Get the Permission Levels assigned
        $PermissionLevels = $RoleAssignment.RoleDefinitionBindings | select-object -ExpandProperty Name
 
        #Remove Limited Access
        $PermissionLevels = ($PermissionLevels | Where-object { $_ -ne "Limited Access"}) -join ","
 
        #Leave Principals with no Permissions
        If($PermissionLevels.Length -eq 0) {Continue}
 
        #Get SharePoint group members
        If($PermissionType -eq "SharePointGroup")
        {
            #Get Group Members
            $GroupMembers = Get-PnPGroupMembers -Identity $RoleAssignment.Member.LoginName
                 
            #Leave Empty Groups
            If($GroupMembers.count -eq 0){Continue}
            $GroupUsers = ($GroupMembers | select-object -expandProperty Title) -join ","
 
            #Add the Data to Object
            $Permissions = New-Object PSObject
            $Permissions | Add-Member NoteProperty Object($ObjectType)
            $Permissions | Add-Member NoteProperty Title($ObjectTitle)
            $Permissions | Add-Member NoteProperty URL($ObjectURL)
            $Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
            $Permissions | Add-Member NoteProperty Users($GroupUsers)
            $Permissions | Add-Member NoteProperty Type($PermissionType)
            $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
            $Permissions | Add-Member NoteProperty GrantedThrough("SharePoint Group: $($RoleAssignment.Member.LoginName)")
            $PermissionCollection += $Permissions
        }
        Else
        {
            #Add the Data to Object
            $Permissions = New-Object PSObject
            $Permissions | Add-Member NoteProperty Object($ObjectType)
            $Permissions | Add-Member NoteProperty Title($ObjectTitle)
            $Permissions | Add-Member NoteProperty URL($ObjectURL)
            $Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
            $Permissions | Add-Member NoteProperty Users($RoleAssignment.Member.Title)
            $Permissions | Add-Member NoteProperty Type($PermissionType)
            $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
            $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")
            $PermissionCollection += $Permissions
        }
    }
    #Output PermissionCollection
    $PermissionCollection

}#END Get-PnPPermissions
   
#MAIN Function to get sharepoint online site permissions report
Function New-PnPSitePermissionRpt()
{
[cmdletbinding()]
Param 
(    
    [Parameter(Mandatory=$false)] [String] $SiteURL, 
    [Parameter(Mandatory=$false)] [switch] $Recursive,
    [Parameter(Mandatory=$false)] [switch] $ScanItemLevel,
    [Parameter(Mandatory=$false)] [switch] $IncludeInheritedPermissions       
)  
    Try {
        #Get the Web
        $Web = Get-PnPWeb
 
        write-verbose "Getting Site Collection Administrators..."
        #Get Site Collection Administrators
        $SiteAdmins = Get-PnPSiteCollectionAdmin
         
        $SiteCollectionAdmins = ($SiteAdmins | select-object -expandProperty Title) -join ","
        #Add the Data to Object
        $Permissions = New-Object PSObject
        $Permissions | Add-Member NoteProperty Object("Site Collection")
        $Permissions | Add-Member NoteProperty Title($Web.Title)
        $Permissions | Add-Member NoteProperty URL($Web.URL)
        $Permissions | Add-Member NoteProperty HasUniquePermissions("TRUE")
        $Permissions | Add-Member NoteProperty Users($SiteCollectionAdmins)
        $Permissions | Add-Member NoteProperty Type("Site Collection Administrators")
        $Permissions | Add-Member NoteProperty Permissions("Site Owner")
        $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")
           
        #output Permissions
        $Permissions
   
        #Function to Get Permissions of All List Items of a given List
        Function Get-PnPListItemsPermission([Microsoft.SharePoint.Client.List]$List)
        {
            write-verbose "-- Run Get-PnPListItemsPermission against $($List.Title)"

            write-verbose "`t `t Getting Permissions of List Items in the List: $($List.Title)"
  
            #Get All Items from List in batches
            #$ListItems = Get-PnPListItem -List $List -PageSize 500
            $ListItems = $List | Get-PnPListItem  -PageSize 500
            write-verbose "-- found $($list.count) list items"
  
            $ItemCounter = 0
            #Loop through each List item
            ForEach($ListItem in $ListItems)
            {
                #Get Objects with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                If($IncludeInheritedPermissions)
                {
                    Get-PnPPermissions -Object $ListItem
                }
                Else
                {
                    #Check if List Item has unique permissions
                    $HasUniquePermissions = Get-PnPProperty -ClientObject $ListItem -Property HasUniqueRoleAssignments
                    If($HasUniquePermissions -eq $True)
                    {
                        #Call the function to generate Permission report
                        Get-PnPPermissions -Object $ListItem
                    }
                }
                $ItemCounter++
                Write-Progress -PercentComplete ($ItemCounter / ($List.ItemCount) * 100) -Activity "Processing Items $ItemCounter of $($List.ItemCount)" -Status "Searching Unique Permissions in List Items of '$($List.Title)'"
            }
        }#END Get-PnPListItemsPermission
 
        #Function to Get Permissions of all lists from the given web
        Function Get-PnPListPermission([Microsoft.SharePoint.Client.Web]$Web)
        {
            write-verbose "-- Run Get-PnPListPermission against $($Web.URL)"
            #Get All Lists from the web
            $Lists = Get-PnPProperty -ClientObject $Web -Property Lists
            write-verbose "-- found $($lists.count) lists"
   
            #Exclude system lists
            $ExcludedLists = @("Access Requests","App Packages","appdata","appfiles","Apps in Testing","Cache Profiles","Composed Looks","Content and Structure Reports","Content type publishing error log","Converted Forms",
            "Device Channels","Form Templates","fpdatasources","Get started with Apps for Office and SharePoint","List Template Gallery", "Long Running Operation Status","Maintenance Log Library", "Images", "site collection images"
            ,"Master Docs","Master Page Gallery","MicroFeed","NintexFormXml","Quick Deploy Items","Relationships List","Reusable Content","Reporting Metadata", "Reporting Templates", "Search Config List","Site Assets","Preservation Hold Library",
            "Site Pages", "Solution Gallery","Style Library","Suggested Content Browser Locations","Theme Gallery", "TaxonomyHiddenList","User Information List","Web Part Gallery","wfpub","wfsvc","Workflow History","Workflow Tasks", "Pages")
             
            $Counter = 0
            #Get all lists from the web   
            ForEach($List in $Lists)
            {
                #Exclude System Lists
                If($List.Hidden -eq $False -and $ExcludedLists -notcontains $List.Title)
                {
                    $Counter++
                    Write-Progress -PercentComplete ($Counter / ($Lists.Count) * 100) -Activity "Exporting Permissions from List '$($List.Title)' in $($Web.URL)" -Status "Processing Lists $Counter of $($Lists.Count)"
 
                    #Get Item Level Permissions if 'ScanItemLevel' switch present
                    If($ScanItemLevel)
                    {
                        #Get List Items Permissions
                        Get-PnPListItemsPermission -List $List
                    }
 
                    #Get Lists with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                    If($IncludeInheritedPermissions)
                    {
                        Get-PnPPermissions -Object $List
                    }
                    Else
                    {
                        #Check if List has unique permissions
                        $HasUniquePermissions = Get-PnPProperty -ClientObject $List -Property HasUniqueRoleAssignments
                        If($HasUniquePermissions -eq $True)
                        {
                            #Call the function to check permissions
                            Get-PnPPermissions -Object $List
                        }
                    }
                }
            }
        }#END Get-PnPListPermission
   
        #Function to Get Webs's Permissions from given URL
        Function Get-PnPWebPermission([Microsoft.SharePoint.Client.Web]$Web) 
        {
            write-verbose "-- Run Get-PnPWebPermission against $($Web.URL)"

            #Call the function to Get permissions of the web
            write-verbose "Getting Permissions of the Web: $($Web.URL)..." 
            Get-PnPPermissions -Object $Web
   
            #Get List Permissions
            write-verbose "`t Getting Permissions of Lists and Libraries..."
            Get-PnPListPermission($Web)
 
            #Recursively get permissions from all sub-webs based on the "Recursive" Switch
            If($Recursive)
            {
                #Get Subwebs of the Web
                $Subwebs = Get-PnPProperty -ClientObject $Web -Property Webs
 
                #Iterate through each subsite in the current web
                Foreach ($Subweb in $web.Webs)
                {
                    #Get Webs with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                    If($IncludeInheritedPermissions)
                    {
                        Get-PnPWebPermission($Subweb)
                    }
                    Else
                    {
                        #Check if the Web has unique permissions
                        $HasUniquePermissions = Get-PnPProperty -ClientObject $SubWeb -Property HasUniqueRoleAssignments
   
                        #Get the Web's Permissions
                        If($HasUniquePermissions -eq $true) 
                        { 
                            #Call the function recursively                            
                            Get-PnPWebPermission($Subweb)
                        }
                    }
                }
            }
        }
        #END Get-PnPWebPermission
 
        #Call the function with RootWeb to get site collection permissions
        Get-PnPWebPermission $Web
   
        write-verbose "`n*** Site Permission Report Generated Successfully!***"
     }
    Catch {
        write-error "Error Generating Site Permission Report!"
        Write-Error $_.Exception    
   }
}#END New-PnPSitePermissionRpt

#-----------------------------------------------------------[Execution]------------------------------------------------------------
Try {
    #import module
    import-module pnp.powershell
    import-module ImportExcel

    #Connect to Admin Center
    $Cred = Get-Credential -message "Enter credentials for $TenantURL"
    Connect-PnPOnline -Url $TenantURL -Credentials $Cred
    $tenantFile = "$($TenantURL.Replace('https://','').Replace('/','')).xlsx"

    #Get All Site collections - Exclude: Seach Center, Mysite Host, App Catalog, Content Type Hub, eDiscovery and Bot Sites
    $SitesCollections = Get-PnPTenantSite | 
        Where-object -Property Template -NotIn ("SRCHCEN#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1") |
        sort-object url
    
    #Loop through each site collection    
    write-verbose "connected to tenant - $(get-pnpconnection | select-object -expandproperty url)"
    
    ForEach($Site in $SitesCollections )
    {
       
        #$ReportName = $site.url.replace("$tenanturl","").replace('/','_')
        #$ReportFile = join-path $path $tenantFile

        #$tenantBasic = $site.url.replace("$tenanturl","").replace('/','_')
        $tenantBasic = $($TenantURL.Replace('https://','').Replace('/',''))
        $tenantFile = "$tenantBasic-permissions-$(get-date -format MMddyyyy).xlsx"        
        $ReportFile = join-path $path $tenantFile
        write-verbose "Generating Report for Site: $($Site.Url)" 
        write-verbose "-- file for site: $tenantFile" 
         
        #Connect to site collection
        try { $SiteConn = Connect-PnPOnline -Url $Site.Url -Credentials $cred -ErrorAction SilentlyContinue }        
        Catch { 
            
            if ($master.ispresent){
                #export to masterlist only        
                write-verbose -verbose "Write to MasterList in file $ReportFile"
                "no access to $($site.url)" |export-excel -path $ReportFile -WorksheetName MasterList -FreezeTopRow -BoldTopRow -AutoSize -Append
            }else {
                #export to multiple pages
                write-verbose -verbose "Write to file $ReportFile ; worksheet $($site.title)"
                "no access to $($site.url)" | export-excel -path $ReportFile -WorksheetName $site.title -FreezeTopRow -BoldTopRow -AutoSize
            }
            continue 
        }
        write-verbose "connected to site - $(get-pnpconnection | select-object -expandproperty url)"
    
        #Call the Function for site collection 
        $npSpParams = @{
            SiteUrl = $site.url
            Recursive = $true
            ScanItemLevel = $false
            IncludeInheritedPermissions = $true
        }

        $curSiteData = New-PnPSitePermissionRpt @npSpParams -verbose
    
        if ($master.ispresent){
            #export to masterlist only        
            write-verbose -verbose "Write to MasterList in file $ReportFile"
            $curSiteData |export-excel -path $ReportFile -WorksheetName MasterList -FreezeTopRow -BoldTopRow -AutoSize -Append
        }else {
            #export to multiple pages
            write-verbose -verbose "Write to file $ReportFile ; worksheet $($site.title)"
            $curSiteData |export-excel -path $ReportFile -WorksheetName $site.title -FreezeTopRow -BoldTopRow -AutoSize
        }
        
        Disconnect-PnPOnline -Connection $SiteConn
    }
}
Catch {
    Write-Error $_.Exception
    continue
}
