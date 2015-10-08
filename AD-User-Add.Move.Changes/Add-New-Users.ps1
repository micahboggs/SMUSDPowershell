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
 
    $Version="1.1.1"

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
    $Username = $FirstName + '.' + $LastName
    $Username = $Username -replace "'"
    
    
    <#  I Want it to error out on duplicate user names, not add a number on the end
    $i = 1

    #Using try catch to supress errors
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

    #while get-aduser returns a result keep trying a different number
    $OrigUsername = $Username
    while((Try-User $Username)){
        $Username = $OrigUsername + $i
        $i++
    } #>

    #unquie username returned
    return $Username

}


function New-Password {

    #generate a new password from GUID to make life easy
    $GUID = [guid]::NewGuid().guid.split('-')

    #in rare cases it fails to meet complexity so having to add a $ on the end
    #return (([string](Get-Date).DayOfWeek) + '-' + $GUID[2].ToUpper() + '-' + $GUID[3] + '$')
    #return (([string](Get-Date).DayOfWeek) + '-' + $GUID[3] + '$')
    return ('changemenow')
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

        $GivenName = $User.GivenName.trim().trim('�')
        $Surname = $User.Surname.trim().trim('�')
        $Initials = $User.Initials.trim().trim('�')
        $Company = $User.company.trim().trim('�')
        $Title = $User.title.trim().trim('�')
        $Title = $Title -replace '�', ' '


        #Reset the failures or set if first one
        $Failures = @()

        #Get username and password
        $Username = New-Username -FirstName $GivenName -LastName $Surname
        $Password = New-Password

        #Need to find template user based on site(companty) and position(title)
        #also should set the $OU for district office departments as they are not based on the template
        #
        switch($Company)
            {
            ("Alvin Dunn Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "ad-teach-template"
                        $AddGroups = "AD Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "ad-teach-template"
                        $AddGroups = "AD Management Email"
                    } else {
                        $templateuser = "ad-ss-template" 
                        $AddGroups = "AD Certificated Email"
                    }
                    $EmailTo = $ADEmail
                }
            ("Carrillo Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "ces-teacher-template"
                        $AddGroups = "CAR Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "ces-teacher-template"
                        $AddGroups = "CAR Management Email"
                    }  else {
                        $templateuser = "ces-ss-template" 
                        $AddGroups = "CAR Classified Email"
                    }
                    $EmailTo = $CESEmail
                }
            ("Double Peak School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "dps-teacher-template"
                        $AddGroups = "DPS Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "dps-teacher-template"
                        $AddGroups = "DPS Management Email"
                    } else {
                        $templateuser = "dps-ss-template" 
                        $AddGroups = "DPS Classified Email"
                    }
                    $EmailTo = $DPSEmail

                }
            ("Discovery Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "dis-teacher-template"
                        $AddGroups = "DIS Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "dis-teacher-template"
                        $AddGroups = "DIS Management Email"
                    
                    } else {
                        $templateuser = "dis-ss-template" 
                        $AddGroups = "DIS Classified Email"
                    }
                    $EmailTo = $DISEmail
                }
            ("Foothills High School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "fhs-teacher-template"
                        $AddGroups = "FH Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "fhs-teacher-template"
                        $AddGroups = "FH Management Email"
                    
                    } else {
                        $templateuser = "fhs-ss-template" 
                        $AddGroups = "FH Classified Email"
                    }
                    $EmailTo = $FHSEmail
                }
            ("Joli Ann Leichtag Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "jaes-teacher-templat"
                        $AddGroups = "JALE Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "jaes-teacher-templat"
                        $AddGroups = "JALE Management Email"
                    
                    } else {
                        $templateuser = "jaes-ss-template" 
                        $AddGroups = "JALE Classified Email"
                    }
                    $EmailTo = $JAESEmail
                }
            ("Knob Hill Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "kh-teacher-template"
                        $AddGroups = "KH Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "kh-teacher-template"
                        $AddGroups = "KH Management Email"
                    
                    } else {
                        $templateuser = "kh-ss-template" 
                        $AddGroups = "KH Classified Email"
                    }
                    $EmailTo = $KHEmail
                }
            ("La Costa Meadows Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "lcm-teacher-template"
                        $AddGroups = "LCM Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "lcm-teacher-template"
                        $AddGroups = "LCM Management Email"
                    
                    } else {
                        $templateuser = "lcm-ss-template" 
                        $AddGroups = "LCM Classified Email"
                    }
                    $EmailTo = $LCMEmail
                }
            ("Mission Hills High School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "mhhs-teacher-templat"
                        $AddGroups = "MHHS Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "mhhs-teacher-templat"
                        $AddGroups = "MHHS Management Email"
                    
                    } else {
                        $templateuser = "mhhs-ss-template" 
                        $AddGroups = "MHHS Classified Email"
                    }
                    $EmailTo = $MHHSEmail
                }
            ("Paloma Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "pal-teacher-template"
                        $AddGroups = "PAL Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "pal-teacher-template"
                        $AddGroups = "PAL Management Email"
                    
                    } else {
                        $templateuser = "pal-ss-template" 
                        $AddGroups = "PAL Classified Email"
                    }
                    $EmailTo = $PALEmail
                }
            ("Richland Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "rl-teacher-template"
                        $AddGroups = "RL Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "rl-teacher-template"
                        $AddGroups = "RL Management Email"
                    
                    } else {
                        $templateuser = "rl-ss-template" 
                        $AddGroups = "RL Classified Email"
                    }
                    $EmailTo = $RLEmail
                }
            ("San Elijo Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "sees-teacher-templat"
                        $AddGroups = "SEES Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "sees-teacher-templat"
                        $AddGroups = "SEES Management Email"
                    
                    } else {
                        $templateuser = "sees-ss-template" 
                        $AddGroups = "SEES Classified Email"
                    }
                    $EmailTo = $SEESEmail
                }
            ("San Elijo Middle School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "sems-teacher-templat"
                        $AddGroups = "SEMS Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "sems-teacher-templat"
                        $AddGroups = "SEMS Management Email"
                    
                    } else {
                        $templateuser = "sems-ss-template" 
                        $AddGroups = "SEMS Classified Email"
                    }
                    $EmailTo = $SEMSEmail
                }
            ("San Marcos Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "smes-teacher-templat"
                        $AddGroups = "SMES Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "smes-teacher-templat"
                        $AddGroups = "SMES Management Email"
                    
                    } else {
                        $templateuser = "smes-ss-template" 
                        $AddGroups = "SMES Classified Email"
                    }
                    $EmailTo = $SMESEmail
                }
            ("San Marcos Middle School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "smms-teacher-templat"
                        $AddGroups = "SMMS Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "smms-teacher-templat"
                        $AddGroups = "SMMS Management Email"
                    
                    } else {
                        $templateuser = "smms-ss-template" 
                        $AddGroups = "SMMS Classified Email"
                    }
                    $EmailTo = $SMMSEmail
                }
            ("San Marcos High School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "smhs-teach-template"
                        $AddGroups = "SMHS Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "smhs-teach-template"
                        $AddGroups = "SMHS Management Email"
                    
                    } else {
                        $templateuser = "smhs-ss-template" 
                        $AddGroups = "SMHS Classified Email"
                    }
                    $EmailTo = $SMHSEmail
                }
            ("Twin Oaks Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "toes-teacher-templat"
                        $AddGroups = "TOE Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "toes-teacher-templat"
                        $AddGroups = "TOE Management Email"
                    
                    } else {
                        $templateuser = "toes-ss-template" 
                        $AddGroups = "TOE Classified Email"
                    }
                    $EmailTo = $TOESEmail
                }
            ("Twin Oaks High School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "tohs-teacher-templat"
                        $AddGroups = "TOHS Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "tohs-teacher-templat"
                        $AddGroups = "TOHS Management Email"
                    
                    } else {
                        $templateuser = "tohs-ss-template" 
                        $AddGroups = "TOHS Classified Email"
                    }
                    $EmailTo = $TOHSEmail
                }
            ("Woodland Park Middle School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "wpms-teacher-templat"
                        $AddGroups = "WPMS Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "wpms-teacher-templat"
                        $AddGroups = "WPMS Management Email"
                    
                    } else {
                        $templateuser = "wpms-ss-template" 
                        $AddGroups = "WPMS Classified Email"
                    }
                    $EmailTo = $WPMSEmail
                }
            ("DO Accounting")
                {
                    $templateuser = "do-ss-template"
                    $department = "Accounting"
                    $OU = "OU=BS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Business Svs.")
                {
                    $templateuser = "do-ss-template"
                    $department = "Business Svs."
                    $OU = "OU=BS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Child Nutrition Svs.")
                {
                    $templateuser = "do-ss-template"
                    $department = "Child Nutrition Svs."
                    $OU = "OU=Users,OU=CNS,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                    $AddGroups = "CNS Classified Email"
                }
            ("DO Curriculum")
                {
                    $templateuser = "do-ss-template"
                    $department = "Curriculum"
                    $OU = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail

                }
            ("DO Human Resources")
                {
                    $templateuser = "do-ss-template"
                    $department = "Human Resources"
                    $OU = "OU=HR&D,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Instructional Svs.")
                {
                    $templateuser = "do-ss-template"
                    $department = "Instructional Svs."
                    $OU = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Kids on Campus")
                {
                    $templateuser = "do-ss-template"
                    $department = "Kids on Campus"
                    $OU = "OU=BS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("Kids on Campus")
                {
                    $templateuser = "do-ss-template"
                    $department = "Kids on Campus"
                    $OU = "OU=KOC,OU=BS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Pupil Personnel Svs.")
                {
                    $templateuser = "do-ss-template"
                    $department = "Pupil Personnel Svs."
                    $OU = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Purchasing")
                {
                    $templateuser = "do-ss-template"
                    $department = "Purchasing"
                    $EmailTo = $DOEmail
                }
            ("DO Special Education")
                {
                    $templateuser = "do-ss-template"
                    $department = "Special Education"
                    $OU = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Technology")
                {
                    $templateuser = "do-ss-template"
                    $department = "Technology"
                    $OU = "OU=IT,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("Facilities Dept.")
                {
                    $templateuser = "do-ss-template"
                    $department = "Facilities Dept."
                    $OU = "OU=Facilities,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                    $AddGroups = "Facilities Staff Email"
                }
            ("Language Assessment Center")
                {
                    $templateuser = "do-ss-template"
                    $OU = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("Maintenance and Operations")
                {
                    $templateuser = "do-ss-template"
                    $EmailTo = $MOEmail
                    $AddGroups = "Maintenance Classified Email"
                }

            }
            if ($Company.contains("DO") -and (-not $Company.contains("Double"))) {
                if ($AddGroups) {
                    $AddGroups += ","
                }
                if ($Title.contains("Director") -or $Title.contains("Principal") -or $Title.contains("Superintendent") -or $Title.contains("Supt.")) {
                    $AddGroups += "DO Management Email"
                } elseif ($Title.contains("Teacher")) {
                    $AddGroups += "DO Certificated Email"
                } else {
                    $AddGroups += "DO Classified Email"
                }
                
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


        try
        {
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
            if (Try-User $Username) {


                #username already exists
                $ReadableFailure = "$Username already exists. Skipping"
                Write-Error "$ReadableFailure"
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
                Company = $template.company
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
            logoutput -SamAccountName $SamAccountName -HomeDirectory $HomeDirectory -Password $Password -OU $OU -Failures $Failures
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
                $ACL = Get-Acl $HomeDirectory -ErrorAction Stop
                $Inherit = [system.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
                $Propagation = [system.security.accesscontrol.PropagationFlags]"None"
                $Rule = New-Object system.security.accesscontrol.filesystemaccessrule($SamAccountName,$HomePermission, $Inherit, $Propagation, "Allow") -ErrorAction Stop
                $ACL.SetAccessRule($Rule)
                Set-Acl $HomeDirectory $ACL -ErrorAction Stop | Out-Null
            }

        }
        catch
        {
            #failed to create home directory, non fatal user can still work so warning only
            $ReadableFailure =  "Failed to create user home directory '$HomeDirectory' for '$SamAccountName'"
            Write-Warning "$ReadableFailure - $_"
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
        logoutput -SamAccountName $SamAccountName -HomeDirectory $HomeDirectory -Password $Password -OU $OU -Failures $Failures -AccountEmail $AccountEmail


        #Remove any variables created incase it causes a duplicate
        Remove-Variable -Name Username,HomeDirectory,HomeRoot,Password,Failures,NewUserInfo,templateuser,department,OU,LoginName,AccountEmail -ErrorAction SilentlyContinue
        

    }#end for each user in CSV

}

function logoutput {
param($SamAccountName,$HomeDirectory,$Password,$OU,$Failures,$AccountEmail)
        $Out = '' | Select-Object Username, Email, HomeDirectory, OU, Password, Failures
        $Out.Username = $SamAccountName
        $Out.Email = $AccountEmail
        $Out.HomeDirectory = $HomeDirectory
        $Out.Password = $Password
        $Out.OU = $OU
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

