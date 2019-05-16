#joins two csvs with extra info from active directory



#scriptpath
$ScriptRootPath = Split-Path -parent $MyInvocation.MyCommand.Definition


#csvfiles
$smusdlistcsv = Join-Path $ScriptRootPath 'smusdlist.csv'
$hrlistcsv = Join-Path $ScriptRootPath 'hrlist.csv'
$outputcsv = Join-Path $ScriptRootPath 'output.csv'


function Main {
    $hrlist = import-csv $hrlistcsv
    foreach($User IN $input){
        $match = $false
        $aduser = get-aduser $user.samaccountname -properties emailaddress,initials,title,company

        foreach($employee in $hrlist){
            if($aduser.emailaddress -eq $employee.hremail) {
                $match = $true

                    #Build hash for splatting
                    $outputhash = @{
                        SAMAccountName = $user.samaccountname
                        Firstname = $aduser.givenname
                        Lastname = $aduser.surname
                        Initials = $aduser.initials
                        email = $aduser.emailaddress
                        OU = $user.site
                        adcompany = $aduser.company
                        hrsite = $employee.hrsite
                        adtitle = $aduser.title
                        hrtitle = $employee.hrtitle
                        enabled = $aduser.enabled
                        hrstatus = $employee.hrstatus
                        matched = $true
                    }


                logoutput @outputhash

            } 

        }
        if(-not $match) {
            write-warning ("no email match for: " + $user.samaccountname + ". Trying by firstname, lastname, and initials.")
                foreach($employee in $hrlist){
                    if(($aduser.givenname -eq $employee.givenname) -and ($aduser.surname -eq $employee.surname) -and ($aduser.initials -eq $employee.initials)) {
                        $match = $true

                            #Build hash for splatting
                            $outputhash = @{
                                SAMAccountName = $user.samaccountname
                                Firstname = $aduser.givenname
                                Lastname = $aduser.surname
                                Initials = $aduser.initials
                                email = $aduser.emailaddress
                                OU = $user.site
                                adcompany = $aduser.company
                                hrsite = $employee.hrsite
                                adtitle = $aduser.title
                                hrtitle = $employee.hrtitle
                                enabled = $aduser.enabled
                                hrstatus = $employee.hrstatus
                                matched = "Email Mismatch, Name Match"
                            }
                        logoutput @outputhash

                    } 
                }


            



        }
        if(-not $match) {
            write-warning ("no match for: " + $user.samaccountname + " using email or name.")
            #Build hash for splatting
            $outputhash = @{
                SAMAccountName = $user.samaccountname
                Firstname = $aduser.givenname
                Lastname = $aduser.surname
                Initials = $aduser.initials
                email = $aduser.emailaddress
                OU = $user.site
                adcompany = $aduser.company
                hrsite = $null
                adtitle = $aduser.title
                hrtitle = $null
                enabled = $aduser.enabled
                hrstatus = $null
                matched = $false
            }
            logoutput @outputhash
        }
    }
 }


 function logoutput {
param($SamAccountName,$firstname,$lastname,$initials,$email,$OU,$adcompany,$hrsite,$adtitle,$hrtitle,$enabled,$hrstatus,$matched)
        $Out = '' | Select-Object SamAccountName, Firstname, Lastname, Initials, Email, OU, adcompany, hrsite, adtitle, hrtitle, enabled, hrstatus, matched
        $Out.SamAccountName = $SamAccountName
        $Out.Firstname = $firstname
        $Out.Lastname = $lastname
        $Out.Initials = $initials
        $Out.Email = $email
        $Out.OU = $OU
        $Out.adcompany = $adcompany
        $Out.hrsite = $hrsite
        $Out.adtitle = $adtitle
        $Out.hrtitle = $hrtitle
        $Out.enabled = $enabled
        $Out.hrstatus = $hrstatus
        $Out.matched = $matched
        $Out
}





###   EXECUTE



import-csv $smusdlistcsv | main | Export-Csv $outputcsv -NoTypeInformation
