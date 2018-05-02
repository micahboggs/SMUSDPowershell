

#scriptpath
$ScriptRootPath = Split-Path -parent $MyInvocation.MyCommand.Definition


#results output
$ResultsFile = Join-Path $ScriptRootPath 'ADUserHomeFolderPermissionsCheckResults.csv'

function main {
    $oulist = Get-ADOrganizationalUnit -searchbase "OU=smusd,dc=smusd,dc=local" -searchscope onelevel -filter * | select name, distinguishedname

    #OUs I don't care about
    $badOU = "Cisco Applications|Cisco Unified Communications|Disabled Users|ONSSI|Outside Contacts|Sample School|Students|VOIP Accounts"

    #go through each valid OU and give me the user accounts for each staff member (must have email address that doesn't = donotsync@smusd.org)
    foreach ($ou in $oulist) {
        if (($badOU -notmatch $ou.name)) {

            $siteADUsers = get-aduser -searchbase $ou.distinguishedname -filter * -properties homedirectory, emailaddress

            foreach ($user in $siteADUsers) {
                if ($user.EmailAddress -and ($user.emailaddress -notlike "donotsync@smusd.org")) {
                    $samaccountname = $user.SamAccountName
                    $homedirectory = $user.HomeDirectory
                    $homedirectorynotexist = $false
                    $homedirectorywrongpermissions = $false
                    $homedirectorynotset = $false
                    #now check their home folder to see if they have correct permissions
                

                    if ($user.HomeDirectory) {
                   
                        if (test-path -path $HomeDirectory) {

                            $acl = get-acl $HomeDirectory
                            $permissions = $acl.access 
                            $correctpermissions = $false
                            $located = $false
                            foreach ($i in $permissions) 
                            {
                                if ($i.IdentityReference -like "smusd\$samaccountname") {
                                    if (($i.FileSystemRights -like "*Modify*") -or ($i.FileSystemRights -like "*FullControl*")) { 
                                        # write-host $homefolder $i.FileSystemRights $i.identityreference
                                        $correctpermissions = $true
                                    }
                                    $located = $true

                                }
                            }
                            if (-not $located) {
                                write-warning "$samaccountname : doesn't have permissions to $homedirectory" 
                                $homedirectorywrongpermissions = $true
                            } else {
                                if (-not $correctpermissions) {
                                    write-warning "$samaccountname : incorrect permissions to $homedirectory"
                                    $homedirectorywrongpermissions = $true
                                }
                            }




                        } else {
                            Write-Warning "$samaccountname : home directory `"$homedirectory`" doesn't exist"
                            $homedirectorynotexist = $true
                        }
                    } else {
                        write-warning "$samaccountname : No home directory set"
                        $homedirectorynotset = $true
                    }






                    #OUTPUT for logging





                    $Out = '' | Select-Object samAccountName, HomeDirectory, Site, HomeDirectoryNotExist, HomeDirectoryWrongPermissions, HomeDirectoryNotSet
                    $OUT.samAccountName = $samaccountname
                    $Out.homedirectory = $homedirectory
                    $Out.site = $ou.name
                    $out.homedirectorynotexist = $homedirectorynotexist
                    $out.homedirectorywrongpermissions = $homedirectorywrongpermissions
                    $out.homedirectorynotset = $homedirectorynotset
                    $out


                    #Cleanup Variables so they don't bork us later
                    Remove-Variable homedirectorynotexist, homedirectorywrongpermissions, homedirectorynotset, samaccountname, homedirectory -ErrorAction SilentlyContinue

                }


            }
        }
    }
}


main | export-csv $ResultsFile
