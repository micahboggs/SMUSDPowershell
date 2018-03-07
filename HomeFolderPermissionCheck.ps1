




$unc = read-host "Enter UNC Path for staff home folder locations (i.e. \\do-fs\staff)"
#$unc = "\\do-fs\staff"

$homefolderlisting = Get-ChildItem $unc

foreach ($samaccountname in $homefolderlisting.name) 
{

    $homefolder = join-path $unc $samaccountname
    $acl = get-acl $homefolder
    $permissions = $acl.access 
    $located = $false
    $correctpermissions = $false
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


    if (($located) -and (-not $correctpermissions)) {
        write-warning "Account `"$samaccountname`" has incorrect permissions on $homefolder"
    }

    if (-not $located) {
        try {
            Get-ADUser $samaccountname > $null
            Write-Warning "$samaccountname has no permissions on $homefolder"
            }
        catch {
            Write-Warning "Account for `"$samaccountname`" does not exist"
            }
        
    }
}

