

function Main {
    Write-host "Looking up Student Accounts."
    $students = get-aduser -filter * -SearchBase "ou=students,ou=smusd,dc=smusd,dc=local" -properties streetaddress
    $count = $students.count
    Write-host "Found $count student accounts."
    $areyousure = read-host "Type 'YES' if you would like to reset the password to the 8 digit birthdate stored in the street address field of each account"
    $script:failurecount = 0
    if ($areyousure -eq 'YES') {
        $currentrecord = 0
        foreach ($student in $students) {
            $currentrecord++
            Write-Progress -Activity "Reseting Password" -CurrentOperation $student.Samaccountname -PercentComplete ($currentrecord/$count * 100)
            #$student | select samaccountname, streetaddress
            $status = "Success"
            #Reset the failures or set if first one
            $Failures = @()
            if ($student.streetaddress -match '\d{8}') {
                $newpwd = ConvertTo-SecureString -string $student.streetaddress -AsPlainText -Force
            } else {
                $writewarning = $student.samaccountname + ": Street Address field does not contain a valid 8 digit string. Street address is '" + $student.streetaddress + "'"
                Write-Warning $writewarning
                $Failures += $writewarning
                $status = "Failure"
                $script:failurecount++
                logoutput -status $status -SamAccountName $student.samaccountname -StreetAddress $student.streetaddress -Failures $Failures
                Remove-Variable writewarning

                continue
            }
            try {
                set-adaccountpassword $student.samaccountname -newpassword $newpwd -Reset
            
            }
            catch {
                $writewarning = "Unable to set password for '" + $student.samaccountname + "' - "
                Write-Error "$writewarning $_"
                $Failures += $writewarning + $_.ToString()
                Remove-Variable writewarning
                $status = "Failure"
                $script:failurecount++
            }
                logoutput -status $status -SamAccountName $student.samaccountname -StreetAddress $student.streetaddress -Failures $Failures

        }
    } else {
        write-host "Passwords will not be reset"
    }
}


function logoutput {
param($Status,$SamAccountName,$StreetAddress,$Failures)
            #OUTPUT for logging
            $Out = '' | Select-Object Status, SamAccountName, StreetAddress, Warnings
            $Out.Status = $status
            $Out.SamAccountName = $SamAccountName
            $OUT.streetaddress = $StreetAddress
            $Out.Warnings = $Failures -join ';'
            $Out
}
$desktop = [Environment]::GetFolderPath("desktop")
$csvfile = "StudentPasswordResetResults.csv"
$fullpath = join-path $desktop $csvfile
$script:failurecount = 0
Main | export-csv $fullpath -NoTypeInformation
if ($script:failurecount -gt '0') {
    write-warning "$script:failurecount Failures detected. Please check log file"
}
write-host "log file saved to $fullpath"
read-host "Press enter to finish."
