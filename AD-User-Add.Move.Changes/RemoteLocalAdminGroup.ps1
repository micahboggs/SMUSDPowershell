$computers = Get-ADComputer -SearchBase "ou=do,ou=smusd,dc=smusd,dc=local" -filter *



foreach ($computer in $computers.name) {
    write-host $computer
    if(Test-Connection -ComputerName $computer -count 1 -Quiet) {
        Invoke-Command { 
            net localgroup administrators | where {$_ -AND $_ -notmatch "command completed successfully"} | select -skip 4 
        } -computer $computer 
    } else {
        write-host "$computer offline"
    }
}