<# RedirectionGroupChecker.ps1

This script will run through all users in each OU and create a helpdesk
ticket for any that do not have permissions to apply the folder redirction GPO.

by: Micah Boggs
#>

# Set variables to appropriate values


###### Do not modify anything below this line ######

######### Functions ##########

function groupLookup {
param ($ou) 
    switch($ou)
        {
            "AD" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-248491" }
            "CES" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-8148" }
            "DIS" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-50018" }
            "DO" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-48325" }
            "DPS" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-151036" } 
            "FHS" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-49804" }
            "JALE" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-248483" } 
            "KH" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-52299" }
            "LCM" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-51156" }
            "M&O" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-48471" }
            "MHHS" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-40289" }
            "PAL" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-248490" }
            "RL" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-49814" }
            "SEES" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-50020" }
            "SEMS" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-248484" }
            "SMES" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-15108" }
            "SMHS" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-1904" }
            "SMMS" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-27176" }
            "TOE" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-38732" }
            "TOHS" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-86477" }
            "WPMS" { $groupSID="S-1-5-21-3934893901-41520433-3260153597-50022" }
            
        }
        return $groupSID
}


function Test-ADGroupMember {
param ($User,$Group)
  trap {Return "error"}
  if (
    Get-ADUser `
      -Filter "memberOf -RecursiveMatch '$((Get-ADGroup $Group).DistinguishedName)'" `
      -SearchBase $((Get-ADUser $User).DistinguishedName)
    ) {$true}
    else {$false}
}




######### Main Program #########


$oulist = Get-ADOrganizationalUnit -searchbase "OU=smusd,dc=smusd,dc=local" -searchscope onelevel -filter * | select name, distinguishedname

#OUs I don't care about, don't bother checking them at all.
$badOU = "Cisco Applications|Cisco Unified Communications|Disabled Users|ONSSI|Outside Contacts|Sample School|Students|VOIP Accounts"

#go through each valid OU and give me the user accounts for each staff member (must have email address that doesn't = donotsync@smusd.org)
foreach ($ou in $oulist) {
    if (($badOU -notmatch $ou.name)) {
        Remove-Variable -Name groupsid, group -ErrorAction SilentlyContinue
        $groupSID = "S-1-5-21-3934893901-41520433-3260153597-2031"
        if (!($groupsid)) { continue }
        $group = get-adgroup $groupsid
        $siteADUsers = get-aduser -searchbase $ou.distinguishedname -filter * -properties memberof, emailaddress
        foreach ($user in $siteADUsers) {
            if ($user.EmailAddress -and ($user.emailaddress -notlike "donotsync@smusd.org")) {

                if (!(Test-ADGroupMember $user.distinguishedname $group.sid)) {
                    write-warning ($ou.name + ": " + $user.samaccountname + " is NOT member of " + $group.name)
                    $subject = ($ou.name + " - " + $user.samaccountname + ": Folder Redirection Issue")
                    #$body = EmailBody -username $user.samaccountname -groupname $group.name
                    #Send-MailMessage -To $helpdeskemail -Body $body -From $fromemail -Subject $subject -ErrorAction Stop
                    #write-host ("~~~~~~~~~~~`n" + $subject + "`n~~~~~~~~~~~")
                    #write-host ($body + "`n~~~~~~~~~~~")
                } 
            }
        }
    }
}
