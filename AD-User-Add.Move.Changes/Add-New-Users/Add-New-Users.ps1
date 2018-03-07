#################################
# SMUSD Add New Users Script
# Written by Micah Boggs (micah.boggs@gmail.com)
#
# Used to add new users to AD
#
#################################

##### Region Module Import ########

Import-module ActiveDirectory

##### End Region ###########


####### Region Configuration #########
 
    $Version="1.5"


    # Uncomment this if testing and you don't want it to send out emails
    $testing = "n"

    #scriptpath
    $ScriptRootPath = Split-Path -parent $MyInvocation.MyCommand.Definition


    #The smtp relay address
    $PSEmailServer = 'smusd-relay.smusd.local'

    #results output
    $ResultsFile = Join-Path $ScriptRootPath 'UserResults.csv'

    #CSV file location
    $CSVFile = Join-Path $ScriptRootPath 'Add-New-Users.csv'
    If (-not (Test-Path $CSVFile)){
        #File Doesn't Exist, abort
        Write-Error "$CSVFile doesn't exist. This file is required"
        Read-Host -Prompt "Press enter to finish..."
        Exit
    }

    ######## Pull in the Email variables from another file. This is just done so I don't sync email addresses into github ################
    ## needs to contain the email arrays used in the file CompanySwitch.ps1
    $EmailFile =  join-path $ScriptRootPath "..\EmailVariables.ps1"
    If (Test-Path $EmailFile){
        #File exists
        . $EmailFile
    } Else {
        #File Doesn't Exist, abort
        Write-Error "$EmailFile doesn't exist. This file is required"
        Read-Host -Prompt "Press enter to finish..."
        Exit
    }

    $ScriptRunAs = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.split('\')
    $ScriptRunAsADObject = Get-ADUser $ScriptRunAs[1] -Properties EmailAddress
    $ScriptRunFirstName = $ScriptRunAsADObject.GivenName
    $ScriptRunLastName = $ScriptRunAsADObject.surname


    #Build config hash for splatting
    $NewUserConfig = @{


        #'me@company.com' replace with email address if you want a BCC copy of the email sent
        AdminEmail = $null

        #Email from
        EmailFrom = '"IT Department" <noreply@smusd.org>'
        #EmailFrom = '"' + "$ScriptRunFirstName $ScriptRunLastName" + '"' + '<' + $ScriptRunAsADObject.EmailAddress + '>'

        #Email address for the service desk, used when email is sent as a point of contact if they have question
        ServiceDeskEmail = 'helpdesk@smusd.org'

        #users home drive permissions
        HomePermission = 'Modify' #Modify is recommended to stop permissions being removed and backups failing!

  



    }

########### End region configuration #########



########### Region functions ################

Function New-EmailTemplate {
param($Name,$surname,$LoginName,$Password,$ServiceDeskEmail,$Title,$Site)

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
						<p>A User Account has been created for<br/>Name: $Name $surname<br/>Title: $Title<br/>Site: $Site</p>
						<h4>Username:<br/>$LoginName</h4>
						<h4>Temporary Password:<br/>$Password</h4>
                        <p>Please inform the user of their username and temporary password provided in this email. The user will be prompted to change their 
                        password the first time they log on to a district PC. 
                        The user may also change their password at <a href="https://cloud.smusd.org">https://cloud.smusd.org</a>. The password must be at least 8 characters long.</p>
                        <br/>
                        <p>Technician Notes: Please add the user to the proper groups, and verify login scripts and home directory.</p>

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




function New-Username {
param($FirstName,$LastName)

    #Get basic username first
    $pattern ='[^a-zA-Z0-9-.]'
    $Username = $FirstName + '.' + $LastName
    $Username = $Username -replace $pattern,''
    
    
    
    
    $i = 1



    #while get-aduser returns a result keep trying a different number
    $OrigUsername = $Username
    while((Try-User $Username)){
        $Question = "$Username already exists in Active directory. Do you want me to try " + $OrigUsername + $i + "? (y/n)"
        $tryanotheruser = read-host -Prompt $Question
        if ($tryanotheruser -eq 'y') {
            $Username = $OrigUsername + $i
            $i++
        } else {
            break
        }
    } 

    #unique username returned
    return $Username

}


function New-Password {

    #generate a new password from GUID to make life easy
    $GUID = [guid]::NewGuid().guid.split('-')

    #in rare cases it fails to meet complexity so having to add a $ on the end
    #return (([string](Get-Date).DayOfWeek) + '-' + $GUID[2].ToUpper() + '-' + $GUID[3] + '$')
    return (([string](Get-Date).DayOfWeek) + '-' + $GUID[3] + '$')
    #return ('changemenow')
}

function Try-User {
param($Username)
    try
    {
        Get-ADUser -Identity $Username
    }
    catch
    {
        return $false
    }
}




function New-User {
param(
    [System.Security.AccessControl.FileSystemRights]$HomePermission='Modify',
    [string]$AdminEmail=$null,
    [string]$EmailFrom,
    [string]$ServiceDeskEmail
    )

    #Get domain name
    $DomainName = "smusd.org"

    foreach($User IN $input){
        Write-Host -NoNewline "."
        #Sanitize the strings
        $pattern ='[^a-zA-Z.]'
        $namePattern = "[^a-zA-Z0-9.' '`'-/]"

        $GivenName = $User.GivenName -replace $namepattern,''
        $Surname = $User.Surname -replace $namepattern,''
        $Initials = $User.Initials -replace $pattern,''
        $Company = $User.company -replace $namepattern,''
        $Title = $User.title -replace $namepattern,''



        #Reset the failures or set if first one
        $Failures = @()
        #Reset add to groups or set if first one
        $AddGroups += @()

        #Get username and password
        $Username = New-Username -FirstName $GivenName -LastName $Surname
        $Password = New-Password

        #Need to find template user based on site(companty) and position(title)
        #also should set the $OU for district office departments as they are not based on the template
        # This is sourced from another file as too much crap uses it.

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
            logoutput -SamAccountName $Username -HomeDirectory $HomeDirectory -Password $Password -OU $OU -Failures $Failures -GivenName $GivenName -Surname $Surname -Initials $Initials -Company $Company -Title $Title
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


        $HomeRoot = '\\'
        foreach ($part in $template.HomeDirectory.split("\") ) { 
            if ($part -ne $template.SamAccountName) { 
                if ($part -ne "") { 
                    $HomeRoot = Join-Path $HomeRoot "$part\"
                }  
            } 
        };

        #Build home path
        $HomeDirectory = Join-Path $HomeRoot $Username


        #Set Department if needed
        if (-not $department) {
            $department = $Company.trim()
        } 
        #check to see if company is overridden
        if ($companyoverride ) {
            $company = $companyoverride
        } else {
            $company = $template.company
        }


        try
        {

            if (Try-User $Username) {


                #username already exists
                $ReadableFailure = "$Username already exists. Skipping"
                Write-Warning "$ReadableFailure"
                $Failures += $ReadableFailure
                remove-variable ReadableFailure
                logoutput -SamAccountName $Username -HomeDirectory $HomeDirectory -Password $Password -OU $OU -Failures $Failures
                continue
            }



            if($Username.length -gt 20){
                Write-Warning "$Username is greater than 20 this might not create properly!"
                $SamAccountName = $Username.Substring(0,20)
                $LoginName = "$Username@$DomainName"
            } else {
                $SamAccountName = $Username
                $LoginName = $Username
            }







            #split out username again, why?
            #because if you have a duplicate user you most likely to have duplicae DN!
            $UsernameSplit = $Username.Split('.')

            # write-host "ou $ou ; template $template ; company $company"

            #Create new user account
            $NewUserInfo = @{
                Instance = $template
                Path = $OU
                Name = "$($UsernameSplit[1]), $($UsernameSplit[0])"
                DisplayName = "$($UsernameSplit[1]), $($UsernameSplit[0])"
                SamAccountName = $SamAccountName
                UserPrincipalName = "$Username@$DomainName"
                EmailAddress = "$Username@$DomainName"
                HomeDirectory = $HomeDirectory
                Department = $department
                Company = $company
                Enabled = $true
                ChangePasswordAtLogon = $true
                AccountPassword = (ConvertTo-SecureString -String $Password -AsPlainText -Force)
                ErrorAction = 'Stop'
                GivenName = $GivenName
                Surname = $Surname
                
                Initials = $Initials
            }
            $User | New-ADUser @NewUserInfo
        }
        catch
        {
            #user failed to create
            $ReadableFailure = "Failed to create user '$Username'"
            Write-Error "$ReadableFailure - $_"
            $Failures += $ReadableFailure + '  -  ' + $_.ToString()
            remove-variable ReadableFailure
            logoutput -SamAccountName $SamAccountName -HomeDirectory $HomeDirectory -Password $Password -OU $OU -Failures $Failures -GivenName $GivenName -Surname $Surname -Initials $Initials -Company $Company -Title $Title
            continue
        }

        try
        {
            #Loop through the groups and add the user to them
            foreach($Group IN $template.memberof){
                Add-ADGroupMember -Identity $Group -Members $SamAccountName -ErrorAction Stop 
            }
            foreach($Group in $AddGroups){
                Add-ADGroupMember -Identity $Group -Members $SamAccountName -ErrorAction Stop 
            }

        }
        catch
        {
            $ReadableFailure = "Failed to add groups for user '$SamAccountName'"
            Write-Warning "$ReadableFailure - $_"
            $Failures += $ReadableFailure + '  -  ' + $_.ToString()
            Remove-Variable ReadableFailure
        }


        try
        {
            #if home directory not present create one
            if(-not (Test-Path $HomeDirectory)){
                New-Item -Path $HomeDirectory -ItemType Directory -ErrorAction stop | Out-Null
            }   

            

        }
        catch
        {
            #failed to create home directory, non fatal user can still work so warning only
            $line = $_.InvocationInfo.ScriptLineNumber
            $ReadableFailure =  "Failed to create user home directory '$HomeDirectory' for '$SamAccountName'"
            Write-Warning "$ReadableFailure at line $line - $_"
            $Failures += $ReadableFailure + '  -  ' + $_.ToString()
            remove-variable ReadableFailure
        } 
        try
        {
            start-sleep -Seconds 5
            $ACL = Get-Acl $HomeDirectory -ErrorAction Stop
            $Inherit = [system.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
            $Propagation = [system.security.accesscontrol.PropagationFlags]"None"
            $Rule = New-Object system.security.accesscontrol.filesystemaccessrule($SamAccountName,$HomePermission, $Inherit, $Propagation, "Allow") -ErrorAction Stop
            $ACL.SetAccessRule($Rule)
            Set-Acl $HomeDirectory $ACL -ErrorAction Stop | Out-Null
        }
        catch
        {
            #failed to set permissions on home folder
            $line = $_.InvocationInfo.ScriptLineNumber
            $ReadableFailure =  "Failed to set permissions on user home directory '$HomeDirectory' for '$SamAccountName'"
            Write-Warning "$ReadableFailure at line $line - $_"
            $Failures += $ReadableFailure + '  -  ' + $_.ToString()
            remove-variable ReadableFailure
        }

        try
        {
            #All seems great so far so lets email them the good news
            $EmailSubject = "New User Created for $GivenName $Surname"
            



            ### For testing so it doesn't email everybody. 
            if ($testing -eq 'y' ) {
                Write-host "Testing, only send emails to test email address"
                $EmailTo = $TestEmailAddress
                $EmailCC = $TestEmailAddress
            }

            ########

            $Body = New-EmailTemplate -Name $GivenName -surname $Surname -LoginName $LoginName -Password $Password  -Title $Title -Site $Company -ServiceDeskEmail $ServiceDeskEmail

            if($AdminEmail){
                Send-MailMessage -To $EmailTo -CC $EmailCC -Bcc $AdminEmail -Body $Body -BodyAsHtml -From $EmailFrom -Subject $EmailSubject -ErrorAction Stop
            }else{
                Send-MailMessage -To $EmailTo -CC $EmailCC -Body $Body -BodyAsHtml -From $EmailFrom -Subject $EmailSubject -ErrorAction Stop
            }
        }
        catch
        {
            $ReadableFailure = "Failed to send email with password for user '$Username'"
            
            Write-Warning "$ReadableFailure - $_"
            $Failures += $ReadableFailure + '  -  ' + $_.ToString()
            remove-variable ReadableFailure
        }

        $AccountEmail = "$Username@$DomainName"
        logoutput -SamAccountName $SamAccountName -HomeDirectory $HomeDirectory -Password $Password -OU $OU -Failures $Failures -AccountEmail $AccountEmail -GivenName $GivenName -Surname $Surname -Initials $Initials -Company $Company -Title $Title


        #Remove any variables created incase it causes a duplicate
        Remove-Variable -Name Username,HomeDirectory,HomeRoot,Password,Failures,NewUserInfo,templateuser,department,companyoverride,OU,LoginName,AccountEmail,AddGroups -ErrorAction SilentlyContinue
        

    }#end for each user in CSV

}

function logoutput {
param($SamAccountName,$HomeDirectory,$Password,$OU,$Failures,$AccountEmail)
        $Out = '' | Select-Object Username, Email, HomeDirectory, OU, Password, Failures, givenName, Surname, Initials, Company, Title
        $Out.Username = $SamAccountName
        $Out.Email = $AccountEmail
        $Out.HomeDirectory = $HomeDirectory
        $Out.Password = $Password
        $Out.OU = $OU
        $Out.GivenName = $GivenName
        $Out.Surname = $Surname
        $Out.Initials = $Initials
        $Out.Company = $Company
        $Out.Title = $Title
        $Out.Failures = $Failures -join ';'
        $Out
}

############# End region functions ##########




############# region Execute  ########

    Write-Host "Add-New-Users v$Version"

    

    if ($testing -eq 'y' ) {
        Write-Warning "System is in test mode!"
        Write-Warning "Accounts will be created, but emails only sent to $testemailaddress"
        $ContinueTest = read-host -prompt "Do you want to continue?  (y/n)"
        if ($ContinueTest -ne 'y') {
            Read-Host -Prompt "Aborting... Press enter to finish..."
            exit
        }
    }
    


    

    Import-Csv $CSVFile | New-User @NewUserConfig | Export-Csv $ResultsFile -NoTypeInformation
    $FinalEmailFrom = '"Powershell Script" <noreply@smusd.org>'
    Send-MailMessage -To $ScriptRunAsADObject.EmailAddress -Body "New User Script Output Report is attached" -From $FinalEmailFrom -Subject "New User Script Output Report" -Attachments $ResultsFile

    Read-Host -Prompt "Press enter to finish..."





###### end region execute ##########

