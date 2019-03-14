<#

WHAT IT DOES:
Automatically searches your exchange servers for tracking logs in the last 15 days (default setting) for whatever sender addresses you place on a txt file.


REQUIREMENTS: 
1 - this script requires RSAT (Remote Server Administrative Tools)  to be installed on the workstation where it'll be executed
2 - requires credentials with the proper permissions to search tracking logs

COOKBOOK:
1 - place a file called senders.txt in the same folder as the script. 1 adress per line, allows for wildcards.

Example:

"
someone@gmail.com
*art@hotmail.com
billgates@micr*

"

2 - run the script. It will ask you for credentials
3 - wait for it to finish.

general progress will be recorded in .\searchlog_$datettoday.txt
results will be recorded in .\results_$datetoday.csv

#>



$usercredential = get-credential

#put here whatever is the name of the exchange server where you'll be connecting to:
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exhcangeserver.yourdomain.com/PowerShell/ -Authentication Kerberos -Credential $UserCredential

Import-PSSession $Session -DisableNameChecking -AllowClobber

$senders = get-content .\senders.txt
$senderscount = (get-content .\senders.txt).count

#edit the value here if you want a different search period:
$date = (get-date).Adddays(-15)

$datetoday = get-date -Format yyyyMMdd
$outputcsv = Write-Output .\results_$datetoday.csv
$log = Write-Output .\searchlog_$datetoday.txt

#in my case this formula will catch the exchange servers on my network. You'll want to change this to reflect your network:
$servers = Get-ADComputer -Filter 'name -like "*exc1*"' -Properties name|Select-Object -expandproperty name


write-host "$([DateTime]::Now)    search will be made for $senders since $date"
Write-Output "$([DateTime]::Now)    search will be made for $senders since $date"|out-file $log -Append

write-host "$([DateTime]::Now)    sender count: $senderscount"
Write-Output "$([DateTime]::Now)    ender count: $senderscount"|out-file $log -Append



foreach ($server in $servers)
{

write-host "$([DateTime]::Now)    beginning search on $server"
Write-Output "$([DateTime]::Now)    beginning search on  $server"|out-file $log -Append

Get-MessageTrackingLog -ResultSize Unlimited -Start $date -server $server| 

where-object{
        $ComparisonResults = foreach ($sender in $senders) {
            $_.sender –like $sender
        }
    
        $ComparisonResults -contains $true
    } | 
select-object Timestamp,Sender,Source,EventId,MessageSubject,{$_.Recipients} | 
export-csv -encoding unicode -notypeinformation $outputcsv -Append

write-host "$([DateTime]::Now)    finished search on $server"
Write-Output "$([DateTime]::Now)    finished search on  $server"|out-file $log -Append

}

remove-PSSession $Session

write-host "$([DateTime]::Now)    Finished. Results available at: $outputcsv"
Write-Output "$([DateTime]::Now)    Finished. Results available at: $outputcsv"|out-file $log -Append

Read-Host -Prompt "Press Enter to exit"