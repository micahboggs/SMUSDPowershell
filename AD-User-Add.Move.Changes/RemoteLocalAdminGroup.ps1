

function GoGoGo {

    $computers = Get-ADComputer -SearchBase "ou=do,ou=smusd,dc=smusd,dc=local" -filter * -Properties LastLogonDate
    #$computers = Get-ADComputer "DO-PUR-813440" -Properties LastLogonDate
    $i=1
    $computercount=$computers.count

    foreach ($computer in $computers) {
            $computername = $computer.name
            $percent = [math]::round($i / $computercount*100)

            Write-Progress -Activity "($percent%) Finding out who's an Admin" -Status  "Looking up $computername" -PercentComplete  $percent
            $i++
            #Reset the failures or set if first one
            $failures = @()
            $administrators = @()



            $Out = '' | Select-Object Computer, LastLogon, Administrators, failures
            $Out.Computer = $computer.name
            $Out.LastLogon = $computer.LastLogonDate


    
        if(Test-Connection -ComputerName $computer.name -count 1 -Quiet) {
            try
            {
                #if ($true) {
                if (test-wsman $computer.name -ErrorAction Stop) {
            
                    try{
                        $administrators = Invoke-Command { 
                            net localgroup administrators | where {$_ -AND $_ -notmatch "command completed successfully"} | select -skip 4 
                        } -computer $computer.name 
                    } 
                    catch {
                    
                        $failures += "Cannot lookup info"
                    }
                }
            
            }
            catch 
            {
            
                $failures += "WinRM not enabled"
                continue
            }
        } else {
        
            $failures += "offline"
        }

        $out.Administrators = $administrators -join ';'
        $out.failures = $failures -join ';'
        $out 

        Remove-Variable out, failures, administrators
    }
}

GoGoGo | export-csv "DOAdmins.csv" -NoTypeInformation