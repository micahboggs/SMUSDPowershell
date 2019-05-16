<#
    BreachedAccountChecker.ps1
    Takes in a csv named accountissues.csv in the same folder the script is ran from
    CSV should contain a column labeled accountname that has samaccountnames
    Script will then lookup AD information for all accounts in the CSV where the email address field is not donotsync@smusd.org
    Will output a csv named BreachedAccounts.csv in the same folder the script is ran from.

#>



function main {
    $BreachedAccounts = import-csv accountissues.csv

    foreach ($account in $BreachedAccounts) {
       $adaccount = get-aduser $account.samaccountname -Properties *
       if ($adaccount.emailaddress -notmatch 'donotsync@smusd.org') {
        $params = @{
                        'adaccount' = $adaccount.SamAccountName
                        'ademailaddress' = $adaccount.EmailAddress
                        'lastset' = $adaccount.PasswordLastSet
                        'lastlogin' = $adaccount.LastLogonDate
                        'enabled' = $adaccount.Enabled
                        'whencreated' = $adaccount.whenCreated
                        'whenchanged' = $adaccount.whenChanged
                        'company' = $adaccount.company
                        'office' = $adaccount.office
                        'department' = $adaccount.Department
                        'title' = $adaccount.Title
                        'DN' = $adaccount.DistinguishedName

                    }

         logoutput @params
       }

    }
}


function logoutput {
param($adaccount,$ademailaddress,$lastset,$lastlogin,$enabled,$whencreated,$whenchanged,$company,$office,$department,$title,$dn)
            #OUTPUT for logging
            $Out = '' | Select-Object SamAccountName, EmailAddress, PasswordLastSet, LastLogin, AccountEnabled, AccountCreationDate, AccountLastChangeDate, Company, Department, Office, Title, DN
            $Out.SamAccountName = $adaccount
            $Out.EmailAddress = $ademailaddress
            $Out.PasswordLastSet = $lastset
            $Out.LastLogin = $lastlogin
            $Out.AccountEnabled = $enabled
            $Out.AccountCreationDate = $whencreated
            $Out.AccountLastChangeDate = $whenchanged
            $Out.Company = $company
            $Out.Department = $department
            $Out.Office = $office
            $Out.Title = $title
            $Out.DN = $dn
            $Out
}

main | export-csv 'BreachedAccounts.csv' -NoTypeInformation
