

####### Configuration

#scriptpath
$ScriptRootPath = Split-Path -parent $MyInvocation.MyCommand.Definition


#results output
$ResultsFile = Join-Path $ScriptRootPath 'employeelist.csv'

#CSV file location
$CSVFile = Join-Path $ScriptRootPath 'hrlist.csv'
If (-not (Test-Path $CSVFile)){
    #File Doesn't Exist, abort
    Write-Error "$CSVFile doesn't exist. This file is required"
    Read-Host -Prompt "Press enter to finish..."
    Exit
}

###### Functions

function main  {
    $EmailRegex = '^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,4})$'
    foreach($User IN $input){
        if($user.emailaddress) {
            try {
                $emailaddress = $user.emailaddress
                $userAccount = get-aduser -filter 'emailaddress -like $emailaddress' -properties title, company, initials
            }
            catch{
                
            }


            #OUTPUT for logging
            $Out = '' | Select-Object HRLastName, HRFirstname, HRInitial, HREmploymentStatus, HRSite, HRTitle, ADSamAccountName, ADLastName, ADFirstname, ADInitial, ADSite, ADTitle, ADEmail
            $OUT.HREmploymentStatus = $user.HREmploymentStatus
            $Out.HRFirstName = $user.HRFirstname
            $Out.HRLastName = $user.HRLastName
            $Out.HRInitial = $user.HRInitial
            $OUT.HRTitle = $user.HRTitle
            $OUT.HRSite = $user.HRSite
            $Out.ADSamAccountName = $UserAccount.SamAccountName
            $OUT.ADLastname = $UserAccount.Surname
            $OUT.ADFirstname = $userAccount.GivenName
            $OUT.ADInitial = $useraccount.Initials
            $OUT.ADSite = $useraccount.company
            $OUT.ADTitle = $useraccount.title
            $OUT.ADEmail = $user.emailaddress
            $Out
        }
    }
}


###### Execute

Import-Csv $CSVFile | main | export-csv $ResultsFile -NoTypeInformation