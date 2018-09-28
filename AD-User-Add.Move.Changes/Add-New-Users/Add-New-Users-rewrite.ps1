<#
    SMUSD Add New Users Script
    Written by Micah Boggs (micah.boggs@gmail.com)

    .SYNOPSIS 
    Used to add new users to AD

    .DESCRIPTION
    - Creates new user account
    - Adds new user to same groups as a template user for the site
    - Adds new user to proper email groups
    - Creates Home folder
    - Assignes correct permissions to home folder
    - Sends email with account information
    - Sends email to helpdesk with tasks site tech is required to do
    - Copies input CSV to a processed folder
    - Writes an output CSV with any issues to a logs folder
    - Erase files older than specified date.

    .CONFIGURATION
    Set Variables in the section "Set the below configuration variables."
    Set Email addresses in the file "emailFunction.ps1" which should be located in the parent directory

    .CVSFIELDS
    -GivenName
    -Surname
    -Initials (not required)
    -Company
    -Title

#>

# Set the below configuration variables.

$PSEmailServer = 'smusd-relay.smusd.local'
$CSVFileName = 'Add-New-Users.csv'
$daystokeep = 30 # how long to keep log files and proccessed CSVs


# Do not edit below this line.

$version = "2.0"
$CSVFile = join-path $PSScriptRoot $CSVFileName
$time = Get-Date -format yyyyMMdd.HHmm
$outfilename = $MyInvocation.MyCommand.Name.split('.')[0] + '.' + $time + '.csv'
$logfile = join-path (join-path $PSScriptRoot 'Logs/') $outfilename
$processedFile = join-path (join-path $PSScriptRoot 'Processed/') $outfilename
$HomePermission = 'Modify'
$testing = 'n'


$EmailFile =  join-path $PSScriptRoot "..\EmailVariables.ps1"
If (Test-Path $EmailFile){
    #File exists
    . $EmailFile
} Else {
    #File Doesn't Exist, abort
    Write-Error "$EmailFile doesn't exist. This file is required"
    Read-Host -Prompt "Press enter to finish..."
    Exit
}





# Functions be here



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
						<h4>Temporary Password:<br/>$Password</h4><br/>
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
    # Returns unique username, adds a number at the end until a unique username is found

    $pattern ='[^a-zA-Z0-9-.]'
    $Username = $FirstName + '.' + $LastName
    $Username = $Username -replace $pattern,''
    $t = 1    
    $OrigUsername = $Username
    while((Try-User $Username)){
        $Question = "$Username already exists in Active directory. Do you want me to try " + $OrigUsername + $t + "? (y/n)"
        $tryanotheruser = read-host -Prompt $Question
        if ($tryanotheruser -eq 'y') {
            $Username = $OrigUsername + $t
            $t++
        } else {
            break
        }
    } 
    return $Username
}

function Try-User {
param($Username)

    try {
        Get-ADUser -Identity $Username
    }
    catch {
        return $false
    }
}

function New-Password {
    #use GUID for randomishness, and add a $ as it's possible it won't be complex enough
    $GUID = [guid]::NewGuid().guid.split('-')
    return (([string](Get-Date).DayOfWeek) + '-' + $GUID[3] + '$')
}

function log-output {
param(
    $SamAccountName=$null,
    $HomeDirectory = $null,
    $Password = $null,
    $OU = $null,
    $Failures = $null,
    $AccountEmail = $null
    )
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

function New-User {
param(
    [System.Security.AccessControl.FileSystemRights]$HomePermission='Modify',
    [string]$EmailFrom,
    [string]$ServiceDeskEmail,
    [int]$lines
    )


    $DomainName = "smusd.org"
    $i = 0
    foreach($User IN $input){




        #Remove any variables created incase it causes a duplicate
        Remove-Variable -Name Username,HomeDirectory,HomeRoot,Password,Failures,NewUserInfo,templateuser,department,companyoverride,OU,LoginName,AccountEmail -ErrorAction SilentlyContinue
        $Failures = @()

        $pattern ='[^a-zA-Z.]'
        $namePattern = "[^a-zA-Z0-9.' '`'-/]"

        $GivenName = $User.GivenName -replace $namepattern,''
        $Surname = $User.Surname -replace $namepattern,''
        $Initials = $User.Initials -replace $pattern,''
        $Company = $User.company -replace $namepattern,''
        $Title = $User.title -replace $namepattern,''


        if (-not $GivenName -or -not $surname -or -not $company -or -not $title) {
            $ErrorMessage = "Missing Required Fields. GivenName, Surname, Company, and Title are mandatory"
            write-error "$ErrorMessage"
            $Failures += $ErrorMessage
            log-output -givenname $user.givenname -surname $user.surname -failures $Failures
            continue
        }
        $Username = New-Username -FirstName $GivenName -LastName $Surname
        $Password = New-Password
        $sitedetails = Get-Company $company $Title
        $properties = @('HomeDirectory', 'memberof', 'scriptpath', 'homedrive', 'company', 'Department', 'Office')
        Write-Progress -Activity "Creating User Accounts" -PercentComplete ($i / $lines*100) -status "$givenname $surname"
        $i++

        try {
            $template = Get-ADUser -Identity $sitedetails.templateuser -Properties $properties -ErrorAction Stop
        } 
        catch {
            $ErrorMessage = "Unable to locate template user for site: `"$company`" - "
            write-error "$ErrorMessage $_"
            $Failures += $ErrorMessage + $_.ToString()
            log-output -samaccountname $username -failures $Failures
            continue
        }
        if ($sitedetails.OU) {
            $OU = $sitedetails.OU
        } else {
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
        }        
        $HomeDirectory = Join-Path $HomeRoot $Username
        if ($sitedetails.department) {
            $department = $sitedetails.department
        } else {
            $department = $Company
        } 
        # KOC site users need company to be the site they are at, not KOC/SITENAME
        if ($sitedetails.companyoverride ) {
            $company = $sitedetails.companyoverride
        } else {
            $company = $template.company
        }



        if($Username.length -gt 20){
            Write-Warning "$Username is greater than 20 this might not create properly!"
            $SamAccountName = $Username.Substring(0,20)
            $LoginName = "$Username@$DomainName" #used in email template
        } else {
            $SamAccountName = $Username
            $LoginName = $Username # used in email template
        }




            
        # If a duplicate user, the DN will probably be duplicate as well, so name has to change.
        $UsernameSplit = $Username.Split('.')
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
            Title = $Title
        }
        
        try {
            new-aduser @NewUserInfo
        }
        catch {
            
            $ErrorMessage = "Unable to create user `"$SamAccountName`" - "
            write-error "$ErrorMessage $_"
            $Failures += $ErrorMessage + $_.ToString()
            log-output -samaccountname $SamAccountName -failures $Failures
            continue
        }


        foreach($Group IN $template.memberof){
            try {
                Add-ADGroupMember -Identity $Group -Members $SamAccountName -ErrorAction Stop 
            }
            catch {
                $ErrorMessage = "Unable to add `"$SamAccountName`" to group `"$Group`" - "
                write-warning "$ErrorMessage $_"
                $Failures += $ErrorMessage + $_.ToString()
            }
        }
        foreach($Group in $sitedetails.AddGroups){
            try {
                Add-ADGroupMember -Identity $Group -Members $SamAccountName -ErrorAction Stop 
            }
            catch {
                $ErrorMessage = "Unable to add `"$SamAccountName`" to group `"$Group`" - "
                write-warning "$ErrorMessage $_"
                $Failures += $ErrorMessage + $_.ToString()
            }
        }
        if(-not (Test-Path $HomeDirectory)){
            try {
                New-Item -Path $HomeDirectory -ItemType Directory -ErrorAction stop | Out-Null
            }
            catch {
                $ErrorMessage = "Unable to add `"$SamAccountName`" to group `"$Group`" - "
                write-warning "$ErrorMessage $_"
                $Failures += $ErrorMessage + $_.ToString()
            }
        }  
        try {
            start-sleep -Seconds 5
            $ACL = Get-Acl $HomeDirectory -ErrorAction Stop
            $Inherit = [system.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
            $Propagation = [system.security.accesscontrol.PropagationFlags]"None"
            $Rule = New-Object system.security.accesscontrol.filesystemaccessrule($SamAccountName,$HomePermission, $Inherit, $Propagation, "Allow") -ErrorAction Stop
            $ACL.SetAccessRule($Rule)
            Set-Acl $HomeDirectory $ACL -ErrorAction Stop | Out-Null
        }
        catch {
            $ErrorMessage = "Failed to set permissions on home directory `"$HomeDirectory`" for `"$SamAccountName`" - "
            write-warning "$ErrorMessage $_"
            $Failures += $ErrorMessage + $_.ToString()
        }

        $EmailFrom = get-emailaddresses "emailFrom"
        $helpdeskEmail = get-emailaddresses "helpdeskEmail"
        $Helpdeskemailfrom = get-emailaddresses "Helpdeskemailfrom"
        $EmailTo = $sitedetails.emailto
        $emailCC = get-emailaddresses "EmailCC"

        $EmailSubject = "New User Created for $GivenName $Surname"
 

        if ($testing -eq 'y' ) { # For testing so it doesn't email everybody.
            Write-host "Testing, only send emails to test email address"
            $TestEmailAddress = get-emailaddresses "TestEmailAddress"
            $EmailTo = $TestEmailAddress
            $EmailCC = $TestEmailAddress
        }

        $Body = New-EmailTemplate -Name $GivenName -surname $Surname -LoginName $LoginName -Password $Password  -Title $Title -Site $Company -ServiceDeskEmail $ServiceDeskEmail
        try {
            Send-MailMessage -To $EmailTo -CC $EmailCC -Body $Body -BodyAsHtml -From $EmailFrom -Subject $EmailSubject -ErrorAction Stop
        }
        catch {
            $ErrorMessage = "Failed to send email for user `"$username`" - "
            write-warning "$ErrorMessage $_"
            $Failures += $ErrorMessage + $_.ToString()
        }
        $siteshortname = $OU.Split(",")[-4].split("=")[1]
        $HelpDeskSubject = "$siteshortname - $EmailSubject"
        $HelpDeskBody = "Please add the user to the proper groups, and verify login scripts and home directory."
        try {
            Send-MailMessage -To $ServiceDeskEmail -Body $HelpDeskBody -From $Helpdeskemailfrom -Subject $HelpDeskSubject -ErrorAction Stop
        }
        catch {
            $ErrorMessage = "Failed to send helpdesk email for user `"$username`" - "
            write-warning "$ErrorMessage $_"
            $Failures += $ErrorMessage + $_.ToString()
        }

        $logdetails = @{
            SamAccountName = $SamAccountName
            HomeDirectory = $HomeDirectory
            Password = $Password
            OU = $OU
            AccountEmail = "$username@$DomainName"
            GivenName = $GivenName
            Surname = $Surname
            Initials = $Initials
            Company = $Company
            Title = $Title
            Failures = $Failures
        }
        log-output @logdetails

        
    } # end of main foreach

}
# End Functions

# Execution
write-host ($MyInvocation.MyCommand.Name + " v$version")

$CompanySwitchFile = join-path $PSScriptRoot "..\CompanySwitch2.ps1"
      
If (Test-Path $CompanySwitchFile){
    #File exists
    . $CompanySwitchFile
} Else {
    #File Doesn't Exist, abort
    Write-Error "$CompanySwitchFile doesn't exist. This file is required"
    Read-Host -Prompt "Press enter to finish..."
    Exit
}

If (-not (Test-Path $CSVFile)){
    #File Doesn't Exist, abort
    Write-Error "$CSVFile doesn't exist. This file is required"
    Read-Host -Prompt "Press enter to finish..."
    Exit
}

$lines = Get-Content($CSVFile) | Measure-Object -Line | select lines
$NewUserConfig = @{
    HomePermissions = $HomePermission
    EmailFrom = $EmailFrom
    ServiceDeskEmail = $helpdeskEmail
    lines = $lines.lines
}

import-csv $CSVFile | new-user @NewUserConfig | Export-Csv $logfile -NoTypeInformation
Move-Item -path $CSVFile -Destination $processedFile -Force

$filestodelete = get-childitem (join-path $PSScriptRoot "\logs\") | ?{$_.LastWriteTime -lt (get-date).adddays($daystokeep * -1)}
$filestodelete += Get-ChildItem (join-path $PSScriptRoot "\processed\") | ?{$_.LastWriteTime -lt (get-date).adddays($daystokeep * -1)}
foreach ($file in $filestodelete) {
    $fullfilepath = join-path $file.Directory $file.Name
    try {
        remove-item $fullfilepath -force
    }
    catch {
        write-error "Unable to delete $fullfilepath : $_"
    }
}