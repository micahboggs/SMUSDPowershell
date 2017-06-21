#################################
# SMUSD User Termination Script
# Written by Micah Boggs (micah.boggs@gmail.com)
#
# Will Disable users specified in a CSV, Remove them from all groups, and move the account to the disabled OU
# Should log what groups the user was a member of.
#
#################################

##### Region Module Import ########

Import-module ActiveDirectory

##### End Region ###########

######### Region Configuration ##############
    
    $version = "1.1.5"

    # Uncomment this if testing and you don't want it to send out emails
    # $testing = "y"

    #Confirm Terminations:
    $Confirm = "Always" #Always ask for confirmation
    #$Confirm = "NotExact" #Only ask for confirmation for users where Initials doesn't match
    #$Confirm = "Never" #Never ask for confirmation. Be Very Careful with this


    #scriptpath
    $ScriptRootPath = Split-Path -parent $MyInvocation.MyCommand.Definition

    #The smtp relay address
    $PSEmailServer = 'smusd-relay.smusd.local'

    #results output
    $ResultsFile = Join-Path $ScriptRootPath 'TerminatedUserResults.csv'

    #CSV file location
    $CSVFile = Join-Path $ScriptRootPath 'TerminatedUsers.csv'
    If (-not (Test-Path $CSVFile)){
        #File Doesn't Exist, abort
        Write-Error "$CSVFile doesn't exist. This file is required"
        Read-Host -Prompt "Press enter to finish..."
        Exit
    }


    $ScriptRunAs = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.split('\')
    $ScriptRunAsADObject = Get-ADUser $ScriptRunAs[1] -Properties EmailAddress
    $ScriptRunFirstName = $ScriptRunAsADObject.GivenName
    $ScriptRunLastName = $ScriptRunAsADObject.surname


    #Build config hash for splatting
    $TermUserConfig = @{

        #Email from
        EmailFrom = '"IT Department" <noreply@smusd.org>'
        #EmailFrom = '"' + "$ScriptRunFirstName $ScriptRunLastName" + '"' + '<' + $ScriptRunAsADObject.EmailAddress + '>'

        #Email address for the service desk, used when email is sent as a point of contact if they have question
        ServiceDeskEmail = 'helpdesk@smusd.org'

    }


    ######## Pull in the Email variables from another file. This is just done so I don't sync email addresses into github ################
    ## needs to contain the arrays:   $EmailCC, $ADEmail $CESEmail $DISEmail $DPSEmail $FHSEmail $JAESEmail $KHEmail $LCMEmail $MHHSEmail $MOEmail $PALEmail $RLEmail $SEESEmail 
    ##      $SEMSEmail $SMESEmail $SMMSEmail $SMHSEmail $TOESEmail $TOHSEmail $WPMSEmail $DOEmail $TestEmailAddress 
    $EmailFile = join-path $ScriptRootPath "..\EmailVariables.ps1"
    If (Test-Path $EmailFile){
        #File exists
        . $EmailFile
    } Else {
        #File Doesn't Exist, abort
        Write-Error "$EmailFile doesn't exist. This file is required"
        Read-Host -Prompt "Press enter to finish..."
        Exit
    }



###### End Region Configuration ############



###### Region Functions #############
Function EmailTemplate {
param($GivenName,$Surname,$ServiceDeskEmail)

@"
Removed/Disabled user account for $GivenName $Surname as requested.

Technician notes: Please remove user files from server and/or move them to a new location

Thank you,
Technology Department
$ServiceDeskEmail

"@

}



function Term-User {
param(
    [string]$EmailFrom,
    [string]$ServiceDeskEmail
    )

    foreach($User IN $input){
        
        #Expecting CSV with Following Fields: GivenName, Surname, Initials, company
        
        #Reset the failures or set if first one
        $Failures = @()

        #Try and get the SamAccountName from Name Provided
        try {
            $GivenName = $User.GivenName.trim().trim('�')
            $SurName = $User.Surname.trim().trim('�')
            $Initials = $User.Initials.trim().trim('�')

            if($Initials){
                $terminateduser = Get-ADUser -Filter {(GivenName -eq $GivenName) -and (Surname -eq $Surname) -and (Initials -eq $Initials)} -Properties Name, SamAccountName, MemberOf, Initials, company, displayname -ErrorAction Stop
            }else{
                $terminateduser = Get-ADUser -Filter {(GivenName -eq $GivenName) -and (Surname -eq $Surname) } -Properties Name, SamAccountName, MemberOf, Initials, company, displayname -ErrorAction Stop
                $noInitials = $true
            }


            
            if (($terminateduser | measure).count -eq "1" -and -not $noInitials) { #Only one user found that matches, Ok to proceed.
                $SamAccountName = $terminateduser.SamAccountName
                if ($terminateduser.displayname.contains(',')) {
                    $OriginalOU = $terminateduser.DistinguishedName.Substring($terminateduser.DistinguishedName.IndexOf(",")+2)
                } else {
                    $OriginalOU = $terminateduser.DistinguishedName.Substring($terminateduser.DistinguishedName.IndexOf(",")+1)
                }
                $AccountDN = $terminateduser.distinguishedname
                $UserFound = $true
            } elseif (($terminateduser | measure).count -gt 1) {
                #More than one account found that matches. Warn, do nothing with accounts and continue
                $writewarning = "More than one account that matches '" + $GivenName + " " + $Initials + " " + $Surname + "'"
                Write-Warning $writewarning
                $Failures += $writewarning
                Remove-Variable writewarning
                $UserFound = $false
            } elseif (($terminateduser | measure).count -eq 0 -or $noInitials) { #No Users match information given. Try to find a user without using the initials

                #$terminateduser = Get-ADUser -Filter {(GivenName -eq $GivenName) -and (Surname -eq $Surname)} -Properties Name, SamAccountName, MemberOf, Initials, company, displayname -ErrorAction Stop

                if (($terminateduser | measure).count -eq 1) { #Only one user found that matches, Ok to proceed, but warn it wasn't an exact match.

                    $SamAccountName = $terminateduser.SamAccountName
                    $AccountDN = $terminateduser.distinguishedname
                    if ($terminateduser.displayname.contains(',')) {
                        $OriginalOU = $terminateduser.DistinguishedName.Substring($terminateduser.DistinguishedName.IndexOf(",")+2)
                    } else {
                        $OriginalOU = $terminateduser.DistinguishedName.Substring($terminateduser.DistinguishedName.IndexOf(",")+1)
                    }
                    $writewarning = "Couldn't find match with Initials, but found: '" + $GivenName + " " + $terminateduser.Initials + " " + $Surname + "'"
                    Write-Warning $writewarning
                    $Failures += $writewarning
                    Remove-Variable writewarning
                    $UserFound = $true
                    $NotExact = $true
                } elseif (($terminateduser | measure).count -gt 1) { #More than one account found that matches. Warn, do nothing with accounts and continue

                    $writewarning = "No Exact Matches, More than one account that matches '" + $GivenName + " " + $Surname + "'"
                    Write-Warning $writewarning
                    $Failures += $writewarning
                    Remove-Variable writewarning
                    $UserFound = $false
                } elseif (($terminateduser | measure).count -eq 0) { #No Matches, unknown user. Warn and move on.
                    $writewarning = "No user matches '" + $GivenName + " " + $Surname + "'"
                    Write-Warning $writewarning
                    $Failures += $writewarning
                    Remove-Variable writewarning
                    $UserFound = $false
                }
            }

        }
        catch {
            $writewarning = "Unable to Lookup Account '" + $GivenName + " " + $Surname + "' - "
            Write-Error "$writewarning $_"
            $Failures += $writewarning + $_.ToString()
            Remove-Variable writewarning
            $UserFound = $false
        }

        if ($UserFound) { #A single account was identified to be disabled
            
            #Confirm User termination if required
            if ($NotExact -and $Confirm -ne "Never") { 
                $TerminateAllowed = read-host -prompt "Initials do not match for Username: $SamAccountName, are you sure you want to terminate this user?  (y/n)"
            } elseif ($Confirm -eq "Always") {
                $TerminateAllowed = read-host -prompt "Are you sure you want to terminate $GivenName $Initials $Surname , (Username: $SamAccountName)?  (y/n)"
            } elseif ($Confirm -eq "Never") {
                $TerminateAllowed = "y"
            }

            if ($TerminateAllowed -eq "y") { #Got Confirmation, or no Confirmation required, proceed with terminating account.
                
                #First, disable the account
                try {
                    Disable-ADAccount $SamAccountName -ErrorAction Stop
                    
                }
                catch {
                    $writewarning = "Unable to Disable Account '" + $SamAccountName + "' - "
                    Write-Error "$writewarning $_"
                    $Failures += $writewarning + $_.ToString()
                    Remove-Variable writewarning
                    $DisableFailure = $true

                }

                if (-not $DisableFailure) { #Account Was disabled, Keep Going


                #Pick OU to move account to depending on Month script is run
                    if ((get-date).Month -ge 4 -and (get-date).Month -le 9) {
                        $TargetOUDN = "OU=Termination date between April 1 - Sep 30,OU=Disabled Users,OU=SMUSD,DC=smusd,DC=local"
                    } else {
                        $TargetOUDN = "OU=Termination date between Oct 1 - March 31,OU=Disabled Users,OU=SMUSD,DC=smusd,DC=local"
                    }
                    try {
                        Move-ADObject -Identity $AccountDN -TargetPath $TargetOUDN -ErrorAction Stop
                    }
                    catch {
                        $writewarning = "Unable to Move Account '" + $SamAccountName + "' - "
                        Write-Error "$writewarning $_"
                        $Failures += $writewarning + $_.ToString()
                        Remove-Variable writewarning
                        $MoveFailure = $true                        
                    }

                    if (-not $MoveFailure) { #Account was moved to a disabled OU, Keep Going

                        #list group membership, and then remove from groups.
                        $Groups = $terminateduser.memberof
                        foreach($Group IN $Groups){
                            try {
                                Remove-ADGroupMember -Identity $Group -Members $SamAccountName -confirm:$false -ErrorAction Stop 
                            }
                            catch {
                                $writewarning = "Failed to remove from group - "
                                Write-Error "$writewarning $_"
                                $Failures += $writewarning + $_.ToString()
                                Remove-Variable writewarning
                                $GroupRemoveFailure = $true
                            }
                        }


                        # Figure out who we should send the email to
                        if ($terminateduser.company) {
                            $Company = $terminateduser.company
                        } else {
                            $Company = "default"
                        }


                        switch($Company)
                        {
                            ("Alvin Dunn Elementery School")
                                {
                                    $EmailTo = $ADEmail
                                    break
                                }
                            ("Carrillo Elementary School")
                                {
                                    $EmailTo = $CESEmail
                                    break
                                }
                            ("Double Peak School")
                                {
                                    $EmailTo = $DPSEmail
                                    break
                                }
                            ("Discovery Elementary School")
                                {
                                    $EmailTo = $DISEmail
                                    break
                                }
                            ("Foothills High School")
                                {
                                    $EmailTo = $FHSEmail
                                    break
                                }
                            ("Joli Ann Leichtag Elementary School")
                                {
                                    $EmailTo = $JAESEmail
                                    break
                                }
                            ("Knob Hill Elementary School")
                                {
                                    $EmailTo = $KHEmail
                                    break
                                }
                            ("La Costa Meadows Elementary School")
                                {
                                    $EmailTo = $LCMEmail
                                    break
                                }
                            ("Mission Hills High School")
                                {
                                    $EmailTo = $MHHSEmail
                                    break
                                }
                            ("Paloma Elementary School")
                                {
                                    $EmailTo = $PALEmail
                                    break
                                }
                            ("Richland Elementary School")
                                {
                                    $EmailTo = $RLEmail
                                    break
                                }
                            ("San Elijo Elementary School")
                                {
                                    $EmailTo = $SEESEmail
                                    break
                                }
                            ("San Elijo Middle School")
                                {
                                    $EmailTo = $SEMSEmail
                                    break
                                }
                            ("San Marcos Elementary School")
                                {
                                    $EmailTo = $SMESEmail
                                    break
                                }
                            ("San Marcos Middle School")
                                {
                                    $EmailTo = $SMMSEmail
                                    break
                                }
                            ("San Marcos High School")
                                {
                                    $EmailTo = $SMHSEmail
                                    break
                                }
                            ("Twin Oaks Elementary School")
                                {
                                    $EmailTo = $TOESEmail
                                    break
                                }
                            ("Twin Oaks High School")
                                {
                                    $EmailTo = $TOHSEmail
                                    break
                                }
                            ("Woodland Park Middle School")
                                {
                                    $EmailTo = $WPMSEmail
                                    break
                                }
                            ("DO Accounting")
                                {
                                    $EmailTo = $DOEmail
                                    break
                                }
                            ("DO Business Svs.")
                                {
                                    $EmailTo = $DOEmail
                                    break
                                }
                            ("DO Child Nutrition Svs.")
                                {
                                    $EmailTo = $DOEmail
                                    break
                                }
                            ("DO Curriculum")
                                {
                                    $EmailTo = $DOEmail
                                    break

                                }
                            ("DO Human Resources")
                                {
                                    $EmailTo = $DOEmail
                                    break
                                }
                            ("DO Instructional Svs.")
                                {
                                    $EmailTo = $DOEmail
                                    break
                                }
                            ("DO Kids on Campus")
                                {
                                    $EmailTo = $DOEmail
                                    break
                                }
                            ("Kids on Campus")
                                {
                                    $EmailTo = $DOEmail
                                    break
                                }
                            ("DO Pupil Personnel Svs.")
                                {
                                    $EmailTo = $DOEmail
                                    break
                                }
                            ("DO Purchasing")
                                {
                                    $EmailTo = $DOEmail
                                    break
                                }
                            ("DO Special Education")
                                {
                                    $EmailTo = $DOEmail
                                    break
                                }
                            ("DO Technology")
                                {
                                    $EmailTo = $DOEmail
                                    break
                                }
                            ("Facilities Dept.")
                                {
                                    $EmailTo = $DOEmail
                                    break
                                }
                            ("Language Assessment Center")
                                {
                                    $EmailTo = $DOEmail
                                    break
                                }
                            ("Maintenance and Operations")
                                {
                                    $EmailTo = $MOEmail
                                    break
                                }
                            ("Transportation")
                                {
                                    $EmailTo = $MOEmail
                                    break
                                }
                            ("Adult Transition Progam")
                                {
                                    $EmailTo = $DOEmail
                                    break
                                }


                            default 
                                {
                                    $EmailTo = $ScriptRunAsADObject.EmailAddress
                                }
                        }



                        try
                        {
                            #All seems great so far so lets send the email

                            ### For testing so it doesn't email everybody. 
                            if ($testing -eq 'y' ) {
                                Write-host "Testing, only send emails to test email address"
                                $EmailTo = $TestEmailAddress
                                $EmailCC = $TestEmailAddress
                            }

                            $EmailSubject = "Account / email for $GivenName $Surname removed"
                            $Body = EmailTemplate -GivenName $GivenName -surname $Surname -ServiceDeskEmail $ServiceDeskEmail
                            Send-MailMessage -To $EmailTo -CC $EmailCC -Body $Body  -From $EmailFrom -Subject $EmailSubject -ErrorAction Stop
                                    
                        }
                        catch
                        {
                            $writewarning = "Failed to send removal email - "
                            Write-Warning "$writewarning $_"
                            $Failures += $writewarning + $_.ToString()
                            Remove-Variable writewarning
                        }
                    }

                }

            } else { #confirmation failed

                $writewarning = "Termination not confirmed."
                Write-Warning "$writewarning"
                $Failures += $writewarning
                Remove-Variable writewarning
                
            }

        }

        if ($MoveFailure -or $DisableFailure -or $GroupRemoveFailure -or $TerminateAllowed -ne "y") {
            $Status = "Failed/Warning"
        } else {
            $Status = "Success"
        }


        #OUTPUT for logging
        $Out = '' | Select-Object Status, GivenName, Surname, Initials, SamAccountName, OriginalOU, OU, Groups, Warnings
        $OUT.Status = $Status
        $Out.GivenName = $GivenName
        $Out.Surname = $Surname
        $Out.Initials = $Initials
        $Out.SamAccountName = $SamAccountName
        $OUT.OriginalOU = $OriginalOU
        $Out.OU = $TargetOUDN
        $Out.Groups = $Groups -join ';'
        $Out.Warnings = $Failures -join ';'
        $Out

        #Cleanup Variables so they don't bork us later
        Remove-Variable MoveFailure, DisableFailure, GroupRemoveFailure, UserFound, Failures, terminateduser, AccountDN, SamAccountName, TargetOUDN, NotExact, Status, OriginalOU, Company -ErrorAction SilentlyContinue

    }
}

# End Region


# Region Execution
    Write-host "Terminate-Users v$version"


    if ($testing -eq 'y' ) {
        Write-warning "System is in test mode!"
        Write-warning "Accounts will be terminated, but emails only sent to $testemailaddress"
        $ContinueTest = read-host -prompt "Do you want to continue?  (y/n)"
        if ($ContinueTest -ne 'y') {
            Read-Host -Prompt "Aborting... Press enter to finish..."
            exit
        }
    }


    Import-Csv $CSVFile | Term-User @TermUserConfig | Export-Csv $ResultsFile -NoTypeInformation

# End Region

