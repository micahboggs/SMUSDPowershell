

# Gets time stamps for all computers in the domain that have NOT logged in since after specified date
 
$time = Read-host "Enter a date in format mm/dd/yyyy"
#$time = '08/28/2016'
$time = get-date ($time)
$date = get-date ($time) -UFormat %d.%m.%y
$filenamedate = get-date ($time) -UFormat %m.%d.%y

 #scriptpath
    $ScriptRootPath = Split-Path -parent $MyInvocation.MyCommand.Definition#results output
    $ResultsFile = Join-Path $ScriptRootPath "OldComputers.$filenamedate.csv"
 
# Get all AD computers with lastLogonTimestamp less than our time
$computers = Get-ADComputer -Filter {LastLogonTimeStamp -lt $time} -Properties LastLogonTimeStamp, OperatingSystem

foreach ($computer in $computers) {
$matched = $computer.DistinguishedName -match ".*,OU=(?<content>.*),OU=[a-zA-Z]*,DC=smusd,DC=local" 
if ($matches) { $computer.smusdsite = $matches['content'] }
Remove-Variable matches  -ErrorAction SilentlyContinue

}
$computers | sort smusdsite, name |

# Output hostname and lastLogonTimestamp into CSV
select-object @{Name="Site"; Expression={$_.smusdsite}},Name,@{Name="Last Logon"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}},OperatingSystem,DistinguishedName | export-csv $ResultsFile -notypeinformation
write-host "file located at $resultsfile"