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
    [string]$aliasdomain,
    [Parameter(Mandatory = $true)]
    [string]$forwarddomain,
    [Parameter(Mandatory = $true)]
    [string[]]$username,
    [Parameter(Mandatory = $true)]
    $Path
)

#import modules
import-module ExchangeOnlineManagement
import-module ImportExcel

$tenantFile = "$aliasDomain-mailboxes-$(get-date -format MMddyyyy-hhmm).xlsx"
$ReportFile = join-path $path $tenantFile

Connect-ExchangeOnline

$curData = @(
    foreach ($user in $username){
        write-verbose -verbose "Process user: $user"
        
        $curMailbox = get-mailbox $user 
        set-mailbox -identity $curMailbox -EmailAddresses @{add = "$user@$aliasdomain"} 
        set-mailbox -identity $curMailbox -DeliverToMailboxAndForward $true -ForwardingSMTPAddress "$user@$forwarddomain"
        $curValMailbox = get-mailbox $user |  select-object Identity,displayname,PrimarySmtpAddress,ForwardingSmtpAddress,
            @{"L"="EmailAddressSummary";"E"={ ($_.EmailAddresses -match "smtp") -join ","} }
        $curValMailbox
    }
)

write-verbose -verbose "Export validation file - $reportfile"
$curData | sort-object Identity | 
        export-excel -path $ReportFile -WorksheetName mailboxes -FreezeTopRow -BoldTopRow -AutoSize -Append