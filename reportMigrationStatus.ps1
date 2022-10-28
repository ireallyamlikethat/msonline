<#
.Example 
$users = Get-MigrationUser -ResultSize Unlimited -BatchId "MigrationBatch01"
.\reportmigrationstatus -username $users -path C:\temp
#>

Param
(
    [Parameter(Mandatory = $true)]
    [string[]]$BatchId,
    [Parameter(Mandatory = $true)]
    $Path
)

Connect-ExchangeOnline

#import modules
import-module ExchangeOnlineManagement
import-module ImportExcel

$tenantFile = "MigrationStatistics-$(get-date -format MMddyyyy-hhmm).xlsx"
$ReportFile = join-path $path $tenantFile

Connect-ExchangeOnline

$curData = @(
    foreach ($batch in $BatchId){
        write-verbose -verbose "Process Batch: $batch"
        $users = Get-MigrationUser -ResultSize Unlimited -BatchId $batch
        foreach ($user in $users){            
            write-verbose -verbose "Process $($user.identity)"
            $userStats = Get-MigrationUserStatistics -Identity $user.identity -IncludeSkippedItems -IncludeReport 
                foreach ($skItem in $userstats.skippeditems){
                    $skItem |select-object Identity , Kind, ScoringClassifications, FolderName, Sender, Recipient, Subject, MessageSize, DateSent, DateReceived, Isvalid, Failure
                }
        }        
    }
)

write-verbose -verbose "Export validation file - $reportfile"
$curData | sort-object Identity | 
        export-excel -path $ReportFile -WorksheetName migration -FreezeTopRow -BoldTopRow -AutoSize -Append