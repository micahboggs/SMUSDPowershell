###################################
# Resets passwords, unlocks accounts and requires password 
# change on next login for usernames provided in a CSV 
#
# Written by Micah Boggs. micah.boggs@gmail.com
################
Function Check-RunAsAdministrator()
{
  #Get current user context
  $CurrentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
  
  #Check user is running the script is member of Administrator Group
  if($CurrentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator))
  {
       Write-host "Script is running with Administrator privileges!"
  }
  else
    {
       #Create a new Elevated process to Start PowerShell
       $ElevatedProcess = New-Object System.Diagnostics.ProcessStartInfo "PowerShell";
 
       # Specify the current script path and name as a parameter
       $ElevatedProcess.Arguments = "& '" + $script:MyInvocation.MyCommand.Path + "'"
 
       #Set the Process to elevated
       $ElevatedProcess.Verb = "runas"
 
       #Start the new elevated process
       [System.Diagnostics.Process]::Start($ElevatedProcess)
 
       #Exit from the current, unelevated, process
       Exit
 
    }
}
 
#Check Script is running with Elevated Privileges
Check-RunAsAdministrator

##### Region Module Import ########

Import-module ActiveDirectory

##### End Region ###########

######### Section config ##########


    $Version="1.2"
    #scriptpath
    $ScriptRootPath = Split-Path -parent $MyInvocation.MyCommand.Definition


    #The smtp relay address
    $PSEmailServer = 'smusd-relay.smusd.local'

    #results output
    $ResultsFile = Join-Path $ScriptRootPath 'Results.csv'

    #CSV file location
    $CSVFile = Join-Path $ScriptRootPath 'ResetUsers.csv'

    $ScriptRunAs = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.split('\')
    $ScriptRunAsADObject = Get-ADUser $ScriptRunAs[1] -Properties EmailAddress
    $ScriptRunFirstName = $ScriptRunAsADObject.GivenName
    $ScriptRunLastName = $ScriptRunAsADObject.surname




######## End section config ########


##### Section Functions #########

function New-Password {

    #generate a new password from GUID to make life easy
    $GUID = [guid]::NewGuid().guid.split('-')

    #in rare cases it fails to meet complexity so having to add a $ on the end
    #return (([string](Get-Date).DayOfWeek) + '-' + $GUID[2].ToUpper() + '-' + $GUID[3] + '$')
    return (([string](Get-Date).DayOfWeek) + '-' + $GUID[3] + '$')
    #return ('changemenow')
}


function UnlockUser{


    foreach($User IN $input){

        $SamAccountName = $User.SamAccountName   
        $Status = "Success"
        


        #make sure user exists
        function Try-User {
        param($SamAccountName)
            try
            {
                $Userobject = Get-ADUser -Identity $SamAccountName
                return $Userobject
            }
            catch
            {
                return $false
            }
        }
        $Userobject = Try-User $SamAccountName
        if($Userobject) {
            $givenname = $Userobject.givenname
            $surname = $Userobject.surname
            try{
                $newpassword = New-Password
                Set-ADAccountPassword $SamAccountName -reset -newpassword (ConvertTo-SecureString -String $newpassword -AsPlainText -Force)
            }
            catch{
                $Warnings = "Unable to reset password for $SamAccountName"
                write-host $Warnings
                $Failures += $Warnings
                Write-Warning $Warnings
                Remove-Variable $warnings
                $Status = "Failed"
            }
            try{
                set-aduser -identity $SamAccountName -ChangePasswordAtLogon $true
            }
            catch{
                $Warnings = "Unable to set change password at logon for $SamAccountName"
                $Failures += $Warnings
                Write-Warning $Warnings
                Remove-Variable $warnings
                $Status = "Failed"
            }
            try {
                Unlock-ADAccount $SamAccountName
            }
            catch{
                $Warnings = "Unable to Unlock account for $SamAccountName"
                $Failures += $Warnings
                Write-Warning $Warnings
                Remove-Variable $warnings
                $Status = "Failed"
            }
        } else {
            $Warnings = "Username $SamAccountName not found"
            Write-Error $Warnings
            $Failures += $Warnings
            $Status = "Failed"
        }

       

        #OUTPUT for logging
        $Out = '' | Select-Object Status, Name, SamAccountName, Password, Warnings
        $OUT.Status = $Status
        $OUT.Name = "$surname, $givenname"
        $Out.SamAccountName = $SamAccountName
        $Out.Password = $newpassword
        $Out.Warnings = $Failures -join ';'
        $Out

         Remove-Variable -Name Failures, Warnings, newpassword, surname, givenname, SamAccountName, Status -ErrorAction SilentlyContinue

    }
}





######### End Section Functions ##########





########## Sectin Run #######

Write-host "BatchResetPasswords.ps1 v$version"

Import-Csv $CSVFile | UnlockUser  | Export-Csv $ResultsFile -NoTypeInformation

Read-Host -Prompt "Press enter to finish..."
######### End Section Run ##########