
function main {
    Write-Host "Reading in Synergy Export file: SMUSDUserChanges.CSV..."
    $synergyusers = import-csv "\\smusd.local\netlogon\Student Export\SMUSDUserChanges.CSV"

    Write-Host "Checking AD for students with mismatched UserPrincipalName and EmailAddress. This may take a while."
    $messedupusers = get-aduser -filter * -searchscope onelevel -searchbase "ou=students,ou=smusd,dc=smusd,dc=local" -properties emailaddress, streetaddress, displayname | ? { $_.userprincipalname -ne $_.emailaddress}
    
    #create object from the $synergyusers object, but only include those that are messed up in active directory.
    $crossreference = $synergyusers | where-object -property samaccountname -in -value $messedupusers.samaccountname

    write-host ("Number of student accounts in AD where UPN does not equal email address:", $messedupusers.count)
    write-host ("Number of these that appear in synergy export: ", $crossreference.count)
    $areyousure = read-host "Type 'YES' if you would like to reset the UserPrincipalName and DisplayName for these accounts."
    if ($areyousure -match "YES") {
        $count = 0
        $total = $crossreference.count
        foreach ($user in $crossreference) {
            
            $count ++
            Write-Progress -Activity "Setting UserPrincipalName and DisplayName" -CurrentOperation $user.Samaccountname -PercentComplete ($count/$total * 100)

            $Failures = @()
            $status = 'Success'
            $verified = 'Failed'

            $sn = $user.sn
            $givenname = $user.givenname
            $displayname = "$sn, $givenname"
            $UPN = $user.mail
            $samaccountname = $user.samaccountname
            try {-whatif
            }
            catch {
                $writewarning = "Unable to set attributes for '$samaccountname' - "
                Write-Error "$writewarning $_"
                $Failures += $writewarning + $_.ToString()
                Remove-Variable writewarning
                $status = "Failure"
                
            }
            #verify it was actually set correctly
            try {
                $aduser = get-aduser -identity $samaccountname -properties displayname
            }
            catch {
                $writewarning = "Unable to lookup '$samaccountname' to verify changes - "
                Write-Error "$writewarning $_"
                $Failures += $writewarning + $_.ToString()
                Remove-Variable writewarning
                $status = "Failure"
            }
            if (($displayname -match $aduser.displayname) -and ($UPN -match $aduser.userprincipalname)) {
                $verified = "Passed"
            }


            logoutput -status $status -SamAccountName $samaccountname -Failures $Failures -Verified $verified
        }
    } else {
        write-warning "Aborting..."
    }


}

function logoutput {
param($Status,$SamAccountName,$Failures, $verified)
            #OUTPUT for logging
            $Out = '' | Select-Object Status, SamAccountName, Verification, Warnings
            $Out.Status = $status
            $Out.SamAccountName = $SamAccountName
            $Out.verified = $verified
            $Out.Warnings = $Failures -join ';'
            $Out
}



$desktop = [Environment]::GetFolderPath("desktop")
$csvfile = "StudentUPNandDNUpdate.csv"
$fullpath = join-path $desktop $csvfile

Main | export-csv $fullpath -NoTypeInformation

write-Host "Script complete. Log file '$csvfile' saved to desktop ($desktop)."
read-host "Press Enter to exit..."
