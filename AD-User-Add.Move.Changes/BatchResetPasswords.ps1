###################################
# Resets passwords, unlocks accounts and requires password 
# change on next login for usernames provided in a CSV 
#
# Written by Micah Boggs. micah.boggs@gmail.com
################


##### Region Module Import ########

Import-module ActiveDirectory

##### End Region ###########

######### Section config ##########


    $Version="1.1.1"
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
                Set-ADAccountPassword $SamAccountName -reset -newpassword (ConvertTo-SecureString -String "changemenow" -AsPlainText -Force)
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
        $Out = '' | Select-Object Status, Name, SamAccountName, Warnings
        $OUT.Status = $Status
        $OUT.Name = "$surname, $givenname"
        $Out.SamAccountName = $SamAccountName
        $Out.Warnings = $Failures -join ';'
        $Out

         Remove-Variable -Name Failures, Warnings, surname, givenname, SamAccountName, Status -ErrorAction SilentlyContinue

    }
}

######### End Section Functions ##########





########## Sectin Run #######

Write-host "BatchResetPasswords.ps1 v$version"

Import-Csv $CSVFile | UnlockUser  | Export-Csv $ResultsFile -NoTypeInformation


######### End Section Run ##########