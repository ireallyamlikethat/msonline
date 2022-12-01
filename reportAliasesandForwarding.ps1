<#

.Example 
$users = 'kayla'
OR
$users = get-content .\usernames.txt    #just a text file with a list of usernames

.\createAliasesandForwarding.ps1 -aliasdomain 365.newsweek.com -forwarddomain gw.newsweek.com -username $users -Path c:\temp
#>

Param
(
    [Parameter(Mandatory = $true)]
    $Path
)

#import modules
import-module ExchangeOnlineManagement
import-module ImportExcel



Connect-ExchangeOnline
$connectedUser = get-connectioninformation |select-object -expand userprincipalname
$reportdomain = ($connecteduser -replace "\w+@") -replace "\.\w+"
write-verbose -verbose "Connected as $connectedUser"


$tenantFile = "$reportDomain-mailboxes-$(get-date -format MMddyyyy-hhmm).xlsx"
$ReportFile = join-path $path $tenantFile

$allMailboxes = get-mailbox
write-verbose -verbose "Found $($allmailboxes.count) mailboxes to report on"

$curData = @(
    $allMailboxes | select-object Identity,displayname,PrimarySmtpAddress,
    @{"L"="ForwardingSmtpAddress";"E"={ $_.ForwardingSmtpAddress  -replace "smtp:"} },
    @{"L"="EmailAddressSummary";"E"={ (($_.EmailAddresses -match "smtp") -join ",") -replace "smtp:"} }
)

write-verbose -verbose "Export validation file - $reportfile"
$curData | sort-object Identity | 
        export-excel -path $ReportFile -WorksheetName mailboxes -FreezeTopRow -BoldTopRow -AutoSize -Append