#################################
# SMUSD User Move Script
# Written by Micah Boggs (micah.boggs@gmail.com)
#
# Used to move users between sites
#
#################################

##### Region Module Import ########

Import-module ActiveDirectory

##### End Region ###########

############## Region Configuration #############


    $Version="1.2.1"

    # Uncomment this if testing and you don't want it to send out emails
    # $testing = "y"

    #Confirm Terminations:
    #$Confirm = "Always" #Always ask for confirmation
    $Confirm = "NotExact" #Only ask for confirmation for users where Initials doesn't match
    #$Confirm = "Never" #Never ask for confirmation. Be Very Careful with this


    #scriptpath
    $ScriptRootPath = Split-Path -parent $MyInvocation.MyCommand.Definition

    #The smtp relay address
    $PSEmailServer = 'smusd-relay.smusd.local'

    #results output
    $ResultsFile = Join-Path $ScriptRootPath 'MoveUserResults.csv'

    #CSV file location
    $CSVFile = Join-Path $ScriptRootPath 'MoveUsers.csv'
    If (-not (Test-Path $CSVFile)){
        #File Doesn't Exist, abort
        Write-Error "$CSVFile doesn't exist. This file is required"
        Read-Host -Prompt "Press enter to finish..."
        Exit
    }


    #get info about the user running the script to use later.
    $ScriptRunAs = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.split('\')
    $ScriptRunAsADObject = Get-ADUser $ScriptRunAs[1] -Properties EmailAddress
    $ScriptRunFirstName = $ScriptRunAsADObject.GivenName
    $ScriptRunLastName = $ScriptRunAsADObject.surname


    #Build config hash for splatting
    $MoveUserConfig = @{

        #Email from
        EmailFrom = '"IT Department" <noreply@smusd.org>'
        #EmailFrom = '"' + "$ScriptRunFirstName $ScriptRunLastName" + '"' + '<' + $ScriptRunAsADObject.EmailAddress + '>'

        #Email address for the service desk, used when email is sent as a point of contact if they have question
        ServiceDeskEmail = 'helpdesk@smusd.org'

    }

    ######## Pull in the Email variables from another file. This is just done so I don't sync email addresses into github ################
    ## needs to contain the arrays:   $EmailCC, $ADEmail $CESEmail $DISEmail $DPSEmail $FHSEmail $JAESEmail $KHEmail $LCMEmail $MHHSEmail $MOEmail $PALEmail $RLEmail $SEESEmail 
    ##      $SEMSEmail $SMESEmail $SMMSEmail $SMHSEmail $TOESEmail $TOHSEmail $WPMSEmail $DOEmail $TestEmailAddress 
    $EmailFile =  "..\EmailVariables.ps1"
    If (Test-Path $EmailFile){
        #File exists
        . $EmailFile
    } Else {
        #File Doesn't Exist, abort
        Write-Error "$EmailFile doesn't exist. This file is required"
        Read-Host -Prompt "Press enter to finish..."
        Exit
    }

############### End Region Configuration ###########



############### Region Functions ###################


Function EmailTemplate {
param($GivenName,$Surname,$ServiceDeskEmail,$Company, $Title, $SamAccountName, $oldsite)

#taken and modified from https://github.com/leemunroe/responsive-html-email-template/blob/master/email.html
#thanks!

@"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="viewport" content="width=device-width" />
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>Really Simple HTML Email Template</title>
<style>
/* -------------------------------------
		GLOBAL
------------------------------------- */
* {
	margin: 0;
	padding: 0;
	font-family: "Helvetica Neue", "Helvetica", Helvetica, Arial, sans-serif;
	font-size: 100%;
	line-height: 1.6;
}
img {
	max-width: 100%;
}
body {
	-webkit-font-smoothing: antialiased;
	-webkit-text-size-adjust: none;
	width: 100%!important;
	height: 100%;
}
/* -------------------------------------
		ELEMENTS
------------------------------------- */
.password {
	color: #348eda;
}
.btn-primary {
	text-decoration: none;
	color: #FFF;
	background-color: #348eda;
	border: solid #348eda;
	border-width: 10px 20px;
	line-height: 2;
	font-weight: bold;
	margin-right: 10px;
	text-align: center;
	display: inline-block;
	border-radius: 25px;
}
.btn-secondary {
	text-decoration: none;
	color: #FFF;
	background-color: #aaa;
	border: solid #aaa;
	border-width: 10px 20px;
	line-height: 2;
	font-weight: bold;
	margin-right: 10px;
	text-align: center;
	display: inline-block;
	border-radius: 25px;
}
.last {
	margin-bottom: 0;
}
.first {
	margin-top: 0;
}
.padding {
	padding: 10px 0;
}
/* -------------------------------------
		BODY
------------------------------------- */
table.body-wrap {
	width: 100%;
	padding: 20px;
}
table.body-wrap .container {
	border: 1px solid #f0f0f0;
}
/* -------------------------------------
		FOOTER
------------------------------------- */
table.footer-wrap {
	width: 100%;	
	clear: both!important;
}
.footer-wrap .container p {
	font-size: 12px;
	color: #666;
	
}
table.footer-wrap a {
	color: #999;
}
/* -------------------------------------
		TYPOGRAPHY
------------------------------------- */
h1, h2, h3 {
	font-family: "Helvetica Neue", Helvetica, Arial, "Lucida Grande", sans-serif;
	color: #000;
	margin: 40px 0 10px;
	line-height: 1.2;
	font-weight: 200;
}
h1 {
	font-size: 36px;
}
h2 {
	font-size: 28px;
}
h3 {
	font-size: 22px;
}
h4 {
    font-size: 16px;
}
p, ul, ol {
	margin-bottom: 10px;
	font-weight: normal;
	font-size: 14px;
}
ul li, ol li {
	margin-left: 5px;
	list-style-position: inside;
}
/* ---------------------------------------------------
		RESPONSIVENESS
		Nuke it from orbit. It's the only way to be sure.
------------------------------------------------------ */
/* Set a max-width, and make it display as block so it will automatically stretch to that width, but will also shrink down on a phone or something */
.container {
	display: block!important;
	max-width: 600px!important;
	margin: 0 auto!important; /* makes it centered */
	clear: both!important;
}
/* Set the padding on the td rather than the div for Outlook compatibility */
.body-wrap .container {
	padding: 20px;
}
/* This should also be a block element, so that it will fill 100% of the .container */
.content {
	max-width: 600px;
	margin: 0 auto;
	display: block;
}
/* Let's make sure tables in the content area are 100% wide */
.content table {
	width: 100%;
}
</style>
</head>
<body bgcolor="#f6f6f6">
<!-- body -->
<table class="body-wrap" bgcolor="#f6f6f6">
	<tr>
		<td></td>
		<td class="container" bgcolor="#FFFFFF">
			<!-- content -->
			<div class="content">
			<table>
				<tr>
					<td>
						<p>User site update for</p><h4>Name: $GivenName $Surname</h4>Username: $SamAccountName<br/>Title: $Title<br/>New Site: $Company<br/>Old Site: $oldsite</p>


                        <br/>
                        <p>Technician Notes: Please add the user to the proper groups, and verify login scripts. Additionally, please work with the user to copy their files to the new site. 
                        Once the user's files are moved, notify the site tech for the old site to clean up the data on the old server.</p>

						<p class="padding"></p>
						<p>Thanks,<br/>IT Department</p>
					</td>
				</tr>
			</table>
			</div>
			<!-- /content -->
			
		</td>
		<td></td>
	</tr>
</table>
<!-- /body -->
<!-- footer -->
<table class="footer-wrap">
	<tr>
		<td></td>
		<td class="container">
			
			<!-- content -->
			<div class="content">
				<table>
					<tr>
						<td align="center">
							<p>Want help logging on? Contact us at: <a href="mailto:$ServiceDeskEmail"><unsubscribe>$ServiceDeskEmail</unsubscribe></a>.
							</p>
						</td>
					</tr>
				</table>
			</div>
			<!-- /content -->
			
		</td>
		<td></td>
	</tr>
</table>
<!-- /footer -->
</body>
</html>
"@

}



function Move-User {
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


            #Sanitize the strings
            $pattern ='[^a-zA-Z-.]'
            $namePattern = "[^a-zA-Z.' '`'-]"

            $GivenName = $User.GivenName -replace $namepattern,''
            $GivenName = $GivenName.trim()
            $Surname = $User.Surname -replace $namepattern,''
            $Surname = $Surname.trim()
            $Initials = $User.Initials -replace $pattern,''
            $Company = $User.company -replace $namepattern,''
            $Company = $Company.trim()
            $Title = $User.title -replace $namepattern,''


            if($Initials){
                $moveuser = Get-ADUser -Filter {(GivenName -eq $GivenName) -and (Surname -eq $Surname) -and (Initials -eq $Initials)} -Properties Name, SamAccountName, MemberOf, Initials, company, displayname -ErrorAction Stop
            }else{
                $moveuser = Get-ADUser -Filter {(GivenName -eq $GivenName) -and (Surname -eq $Surname) } -Properties Name, SamAccountName, MemberOf, Initials, company, displayname -ErrorAction Stop
                $noInitials = $true
            }
            
            if (($moveuser | measure).count -eq "1" -and -not $noInitials) { #Only one user found that matches, Ok to proceed.
                
                $SamAccountName = $moveuser.SamAccountName
                if ($moveuser.displayname.contains(',')) {
                    $OriginalOU = $moveuser.DistinguishedName.Substring($moveuser.DistinguishedName.IndexOf(",")+2)
                } else {
                    $OriginalOU = $moveuser.DistinguishedName.Substring($moveuser.DistinguishedName.IndexOf(",")+1)
                }
                $AccountDN = $moveuser.distinguishedname
                $UserFound = $true
            } elseif (($moveuser | measure).count -gt 1) {
                #More than one account found that matches. Warn, do nothing with accounts and continue
                $writewarning = "More than one account that matches '" + $GivenName + " " + $Initials + " " + $Surname + "'"
                Write-Warning $writewarning
                $Failures += $writewarning
                Remove-Variable writewarning
                $UserFound = $false
            } elseif (($moveuser | measure).count -eq 0 -or $noInitials) { #No Users match information given. Try to find a user without using the initials

                $moveuser = Get-ADUser -Filter {(GivenName -eq $GivenName) -and (Surname -eq $Surname)} -Properties Name, SamAccountName, MemberOf, Initials, company, displayname -ErrorAction Stop

                if (($moveuser | measure).count -eq 1) { #Only one user found that matches, Ok to proceed, but warn it wasn't an exact match.

                    $SamAccountName = $moveuser.SamAccountName
                    $AccountDN = $moveuser.distinguishedname
                    if ($moveuser.displayname.contains(',')) {
                        $OriginalOU = $moveuser.DistinguishedName.Substring($moveuser.DistinguishedName.IndexOf(",")+2)
                    } else {
                        $OriginalOU = $moveuser.DistinguishedName.Substring($moveuser.DistinguishedName.IndexOf(",")+1)
                    }
                    $writewarning = "Couldn't find match with Initials, but found: '" + $GivenName + " " + $moveuser.Initials + " " + $Surname + "'"
                    Write-Warning $writewarning
                    $Failures += $writewarning
                    Remove-Variable writewarning
                    $UserFound = $true
                    $NotExact = $true
                } elseif (($moveuser | measure).count -gt 1) { #More than one account found that matches. Warn, do nothing with accounts and continue

                    $writewarning = "No Exact Matches, More than one account that matches '" + $GivenName + " " + $Surname + "'"
                    Write-Warning $writewarning
                    $Failures += $writewarning
                    Remove-Variable writewarning
                    $UserFound = $false
                } elseif (($moveuser | measure).count -eq 0) { #No Matches, unknown user. Warn and move on.
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
            
            #Confirm User move if required
            if ($NotExact -and $Confirm -ne "Never") { 
                $MoveAllowed = read-host -prompt "Initials do not match for Username: $SamAccountName, are you sure you want to move this user?  (y/n)"
            } elseif ($Confirm -eq "Always") {
                $MoveAllowed = read-host -prompt "Are you sure you want to move $GivenName $Initials $Surname , (Username: $SamAccountName)?  (y/n)"
            } elseif ($Confirm -eq "Never") {
                $MoveAllowed = "y"
            } elseif ($Confirm -eq "NotExact") {
                $MoveAllowed = "y"
            }

            if ($MoveAllowed -eq "y") { #Got Confirmation, or no Confirmation required, proceed with Moving account.
                


                $OldSite=$moveuser.company


                #Pick OU to move account to based on company


                #Need to find template user based on site(companty) and position(title)
                #also should set the $OU for district office departments as they are not based on the template
                #
                $CompanySwitchFile = join-path $ScriptRootPath "..\CompanySwitch.ps1"
      
                If (Test-Path $CompanySwitchFile){
                    #File exists
                    . $CompanySwitchFile
                } Else {
                    #File Doesn't Exist, abort
                    Write-Error "$CompanySwitchFile doesn't exist. This file is required"
                    Read-Host -Prompt "Press enter to finish..."
                    Exit
                }

                #Set Department if needed
                if (-not $department) {
                    $department = $Company.trim()
                } 

                #try to get info from template user
        
                try
                {
            
                    $template = get-aduser -Identity $templateuser -Properties HomeDirectory, memberof, scriptpath, homedrive, company, Department, Office -ErrorAction Stop



                }
                catch
                {
                    #Failed to lookup template user
                    $ReadableFailure = "Failed to lookup template for user '$Username'"
                    Write-Error "$ReadableFailure - $_"
                    $Failures += $ReadableFailure + '  -  ' + $_.ToString()
                    remove-variable ReadableFailure
                    logoutput -SamAccountName $Username -HomeDirectory $HomeDirectory -Password $Password -OU $OU -Failures $Failures
                    continue

                }

                #Build OU Path
                if (-not $OU) {
                
                    $OU = $template.DistinguishedName.Substring($template.DistinguishedName.IndexOf(",")+1)
                    if ($Title.contains("Teacher")) {
                        $OU = 'OU=Teachers,' + $OU
                    } elseif ($Title.Contains('Principal')) {
                        $OU = 'OU=AdminStaff,' + $OU
                    } else {
                        $OU = 'OU=Support Staff,' + $OU
                    }
                }

                $TargetOUDN = $OU


                #ok, move them to the OU
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

                if (-not $MoveFailure) { #Account was moved to target OU, Keep Going



                    #Loop through the groups and add the user to them
                    foreach($Group IN $template.memberof){
                        Add-ADGroupMember -Identity $Group -Members $SamAccountName -ErrorAction Stop 
                    }
                    foreach($Group in $AddGroups){
                        Add-ADGroupMember -Identity $Group -Members $SamAccountName -ErrorAction Stop 
                    }



                    #Lets write the title, department, and company

                    Set-ADUser $SamAccountName -title $title -Department $department -Company $template.company


                    try
                    {
                        #All seems great so far so lets send the email


                        ### For testing so it doesn't email everybody. 
                        if ($testing -eq 'y' ) {
                            Write-host "Testing, only send emails to test email address"
                            $EmailTo = $TestEmailAddress
                            $EmailCC = $TestEmailAddress
                        }

                        $EmailSubject = "Site update for $GivenName $Surname"
                        $Body = EmailTemplate -GivenName $GivenName -surname $Surname -Company $Company -ServiceDeskEmail $ServiceDeskEmail -Title $Title -SamAccountName $SamAccountName -oldsite $oldsite
                        Send-MailMessage -To $EmailTo -CC $EmailCC -BodyAsHtml -Body $Body  -From $EmailFrom -Subject $EmailSubject -ErrorAction Stop 
                                    
                    }
                    catch
                    {
                        $writewarning = "Failed to send removal email - "
                        Write-Warning "$writewarning $_"
                        $Failures += $writewarning + $_.ToString()
                        Remove-Variable writewarning
                    }
                }

                

            } else { #confirmation failed

                $writewarning = "Move not confirmed."
                Write-Warning "$writewarning"
                $Failures += $writewarning
                Remove-Variable writewarning
                
            }

        }

        if ($MoveFailure -or $MoveAllowed -ne "y") {
            $Status = "Failed/Warning"
        } else {
            $Status = "Success"
        }


        #OUTPUT for logging
        $Out = '' | Select-Object Status, GivenName, Surname, Initials, SamAccountName, OriginalOU, OU, Warnings
        $OUT.Status = $Status
        $Out.GivenName = $GivenName
        $Out.Surname = $Surname
        $Out.Initials = $Initials
        $Out.SamAccountName = $SamAccountName
        $OUT.OriginalOU = $OriginalOU
        $Out.OU = $TargetOUDN
        $Out.Warnings = $Failures -join ';'
        $Out

        #Cleanup Variables so they don't bork us later
        Remove-Variable MoveFailure, OU, UserFound, Failures, moveuser, AccountDN, SamAccountName, TargetOUDN, NotExact, Status, OriginalOU, noInitials, Company, templateuser -ErrorAction SilentlyContinue

    }
}

############ End Region Functions ############


############# Region Execution  ##############
    Write-Host "Move-Users v$Version"


    if ($testing -eq 'y' ) {
        Write-Warning "System is in test mode!"
        Write-Warning "Accounts will be Moved, but emails only sent to $testemailaddress"
        $ContinueTest = read-host -prompt "Do you want to continue?  (y/n)"
        if ($ContinueTest -ne 'y') {
            Read-Host -Prompt "Aborting... Press enter to finish..."
            exit
        }
    }



    Import-Csv $CSVFile | Move-User @MoveUserConfig | Export-Csv $ResultsFile -NoTypeInformation
    
    $FinalEmailFrom = '"Powershell Script" <noreply@smusd.org>'
    Send-MailMessage -To $ScriptRunAsADObject.EmailAddress -Body "Move User Script Output Report is attached" -From $FinalEmailFrom -Subject "Move User Script Output Report" -Attachments $ResultsFile

    Read-Host -Prompt "Press enter to finish..."

############## End Region Execution #############

