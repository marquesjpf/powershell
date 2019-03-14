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
