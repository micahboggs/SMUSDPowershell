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
    
  
    #scriptpath
    $ScriptRootPath = Split-Path -parent $MyInvocation.MyCommand.Definition


    #results output
    $outfile = Join-Path $ScriptRootPath 'outfile.csv'
    $ActiveUsersNoAccount = Join-Path $ScriptRootPath 'ActiveUsersNoAccount.csv'
    $TerminatedUsersWithAccount  = Join-Path $ScriptRootPath 'TerminatedUsersWithAccount.csv'
    $TerminatedUsersMaybeAccount = Join-Path $ScriptRootPath 'TerminatedUsersMaybeAccount.csv'

    #CSV file location
    $CSVFile = Join-Path $ScriptRootPath 'Employees.csv'
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






###### End Region Configuration ############



###### Region Functions #############


function CheckEmployee {
param(
    )

    foreach($Employee IN $input){
        
        #Expecting CSV with Following Fields: LastName, FirstName, MiddleInitial, Status, Site, PositionTitle, TerminationDate

        
        #Reset the failures or set if first one
        $Failures = @()

        #Try and get the SamAccountName from Name Provided
        try {

            #Sanitize the strings
            $AlphaOnlypattern ='[^a-zA-Z]'
            $namePattern = "[^a-zA-Z0-9.' '`'-/]"

            $GivenName = $Employee.FirstName -replace $namePattern
            $SurName = $Employee.LastName -replace $namePattern
            $Initials = $Employee.MiddleInitial -replace $AlphaOnlypattern
            $Status = $Employee.Status -replace $namePattern
            $Site = $Employee.Site -replace $namePattern
            $PositionTitle = $Employee.PositionTitle -replace $namepatern


            ### to avoid error, don't filter on initials unless initials are provided.

            if($Initials){
                $UserAccount = Get-ADUser -Filter {(GivenName -eq $GivenName) -and (Surname -eq $Surname) -and (Initials -eq $Initials)} -Properties Name, SamAccountName, Initials, company, displayname, title, emailaddress, enabled -ErrorAction Stop
            }else{
                $UserAccount = Get-ADUser -Filter {(GivenName -eq $GivenName) -and (Surname -eq $Surname) } -Properties Name, SamAccountName, Initials, company, displayname, title, emailaddress, enabled -ErrorAction Stop
                $noInitials = $true
            }


            #try and get the SAM Account name. Basically, search based on provided name and see if we can get an exact match.
            #Since Initials might not be in active directory, we need to treat them as optional and output a spreadsheet of maybe matches.
            #Any Account with a status of Active, Leave of Absense, or Administrative Leave, should have an account. Output a csv of matching employees that do not have an account.
            #Any other status should either be disabled, or no account. Output a csv of accounts that have an active account, but not one of the above statuses.


            if (($UserAccount | measure).count -eq "1" -and -not $noInitials) { #Only one user found that matches, Ok to proceed. Don't want to match a user in AD with initials if the spreadsheet doesn't have initials
                $SamAccountName = $UserAccount.SamAccountName
                $AccountDN = $UserAccount.distinguishedname
                $UserFound = $true
            } elseif (($UserAccount | measure).count -gt 1) {
                #More than one account found that matches. Warn, do nothing with accounts and continue
                $writewarning = "More than one account that matches '" + $GivenName + " " + $Initials + " " + $Surname + "'"
                Write-Warning $writewarning
                $Failures += $writewarning
                Remove-Variable writewarning
                $UserFound = $false
                $multipleusers = $true
            } elseif (($UserAccount | measure).count -eq 0 -or $noInitials) { #No Users match information given, or spreadsheet doesn't list initials. Try to find a user without using the initials
                $UserAccount = Get-ADUser -Filter {(GivenName -eq $GivenName) -and (Surname -eq $Surname) } -Properties Name, SamAccountName, Initials, company, displayname, title, emailaddress, enabled -ErrorAction Stop


                if (($UserAccount | measure).count -eq 1 -and (-not ($useraccount.Initials)) -and ($noInitials)) { #Only one user account, neither spreadsheet nor AD have initials, so its an exact match
                    $SamAccountName = $UserAccount.SamAccountName
                    $AccountDN = $UserAccount.distinguishedname
                    $UserFound = $true
                } elseif (($UserAccount | measure).count -eq 1) { #Only one user found that matches, Ok to proceed, but warn it wasn't an exact match.
                    
                    if ( (($useraccount.Initials) -and ( -not $Initials )) -or ( -not $useraccount.Initials ) ) {


                        # If the SS has initials and AD does not, I want it logged as user found with a warning
                        # If the SS has initials and AD has DIFFERENT initals, I want it logged as user not found.
                        # If the SS has initials and AD has same initials, it would have been caught above, so this will never be active here
                        # IF the SS has no initials, and AD has initials, I want it logged as user found with a warning.
                        # IF the SS has no initials, and AD has no initials, it is caught above so this will never be active here.
                        # so that means, if AD has no initials OR (AD has initials, but SS does not

                        $SamAccountName = $UserAccount.SamAccountName
                        $AccountDN = $UserAccount.distinguishedname
 
                       

                        # HR said to populate AD with HR's initial in this scenario.

                        $writewarning = "Couldn't find match with Initials, but found: '" + $GivenName + " " + $UserAccount.Initials + " " + $Surname + "'"
                        Write-Warning $writewarning
                        $Failures += $writewarning
                        Remove-Variable writewarning

                        try {
                            if ( $Initials ) {
                                set-aduser -identity $useraccount.samaccountname -Initials $Initials
                            }
                        }
                        catch {
                            $writewarning = "Couldn't set initials"
                            Write-Warning $writewarning
                            $Failures += $writewarning
                            Remove-Variable writewarning
                        }


                        $UserFound = $true
                        $NotExact = $true

                    } else {
                        $writewarning = "No user matches '" + $GivenName + " " + $Surname + "'"
                        Write-Warning $writewarning
                        $Failures += $writewarning
                        Remove-Variable writewarning
                        $UserFound = $false
                    }
                } elseif (($UserAccount | measure).count -gt 1) { #More than one account found that matches. Warn, do nothing with accounts and continue

                    $writewarning = "No Exact Matches, More than one account that matches '" + $GivenName + " " + $Surname + "'"
                    Write-Warning $writewarning
                    $Failures += $writewarning
                    Remove-Variable writewarning
                    $UserFound = $false
                    $multipleusers = $true
                } elseif (($UserAccount | measure).count -eq 0) { #No Matches, unknown user. Warn and move on.
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





        if ( $multipleusers ) { #multiple user accounts is always a problem
            $ProblemAccount = $true        
        } elseif ( ($Status -eq 'Active' -or $Status -eq 'Administrative Leave' -or $Status -eq 'Leave of Absence') -and $UserFound -and -not $NotExact) { #Found exact match for account that should be active
            $ProblemAccount = $false
        } elseif ( -not ($Status -eq 'Active' -or $Status -eq 'Administrative Leave' -or $Status -eq 'Leave of Absence') -and  -not $Userfound) { #Couldn't find an account that for inactive statuses
            $ProblemAccount = $false
        } elseif ( ($Status -eq 'Active' -or $Status -eq 'Administrative Leave' -or $Status -eq 'Leave of Absence') -and -not $UserFound) { #HR active status, but cannot find AD User account
            $ProblemAccount = $true
      

        } else { #otherwise it's going to be a problem account except for a few specific reasons.
         
            # if Inactive on HR Spreadsheet, and account disabled in AD, not a problem (most likely)
            # If Job title is Avid Tutor or Student worker, should never have an account
            # If job title Coach something, fine arts instructor, music instructor, NTS, NTS/CG, 
            # Psych Intern, it doesn't matter if an account exists for ACTIVE users. (Inactive needs account deleted or disabled.)


            $ProblemAccount = $true
            if ($useraccount.enabled -eq $false) { #unless the account is disabled
                $ProblemAccount = $false 
            } else {
                $ProblemAccount = $true
            }
        }





        #OUTPUT for logging
        $Out = '' | Select-Object HRLastName, HRFirstname, HRInitial, HREmploymentStatus, HRSite, HRTitle, ADSamAccountName, ADLastName, ADFirstname, ADInitial, ADSite, ADTitle, ADEmail, ADAccountEnabled, ProblemAccount, warnings
        $OUT.HREmploymentStatus = $Status
        $Out.HRFirstName = $GivenName
        $Out.HRLastName = $Surname
        $Out.HRInitial = $Initials
        $OUT.HRTitle = $PositionTitle
        $OUT.HRSite = $Site
        $Out.ADSamAccountName = $SamAccountName
        $OUT.ADLastname = $UserAccount.Surname
        $OUT.ADFirstname = $userAccount.GivenName
        $OUT.ADInitial = $useraccount.Initials
        $OUT.ADSite = $useraccount.company
        $OUT.ADTitle = $useraccount.title
        $OUT.ADEmail = $useraccount.emailaddress
        $Out.Warnings = $Failures -join ';'
        $out.ADAccountEnabled = $useraccount.enabled
        $OUT.ProblemAccount = $ProblemAccount
        $Out


        #Cleanup Variables so they don't bork us later
        Remove-Variable MoveFailure, DisableFailure, GroupRemoveFailure, UserFound, multipleusers, Failures, noinitials, terminateduser, AccountDN, SamAccountName, TargetOUDN, NotExact, Status, OriginalOU, Company, ProblemAccount -ErrorAction SilentlyContinue

    }
}

# End Region


# Region Execution





    Import-Csv $CSVFile | CheckEmployee | export-csv $outfile -NoTypeInformation

# End Region

