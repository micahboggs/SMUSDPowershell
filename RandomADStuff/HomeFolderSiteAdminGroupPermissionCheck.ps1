




$unc = read-host "Enter UNC Path for staff home folder locations (i.e. \\do-fs\staff)"
#$unc = "\\do-fs\staff"

$homefolderlisting = Get-ChildItem $unc

foreach ($folder in $homefolderlisting.name) 
{

    $homefolder = join-path (join-path $unc $folder) "/Documents"
    if (-not(test-path $homefolder)) {
        continue
    }
    try {
        $acl = get-acl $homefolder
    }
    catch {
        Write-Warning "$homefolder - Unable to check permissions"
    }

    $permissions = $acl.access 
    $located = $false
    $correctpermissions = $false
    $groupname = "smusd\missionhillsHSSiteAdmins"
    foreach ($i in $permissions) 
    {
        if ($i.IdentityReference -like "$groupname") {
            if (($i.FileSystemRights -like "*FullControl*")) { 
                write-host $homefolder $i.identityreference $i.FileSystemRights
                $correctpermissions = $true
            }
            $located = $true

        }
    }


    if (($located) -and (-not $correctpermissions)) {
        write-warning "Group `"$groupname`" has incorrect permissions on $homefolder"
    }

    if (-not $located) {

            Write-Warning "$groupname has no permissions on $homefolder"

        
    }
}

