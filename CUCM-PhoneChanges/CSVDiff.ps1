<#write-host "Compare 2 different CSVs and output the differences"
$original = Read-Host -prompt "Original filename"
$new = Read-Host -prompt "New Filename"
$differences = Read-Host -prompt "Results filename"
#>

    $siteInitials = Read-Host -prompt "Please enter the site Abbreviation you used in the filenames"
    #$SiteInitials = "JAES" # For testing so I don't have to type the damn thing in every time

    #scriptpath
    $ScriptRootPath = Split-Path -parent $MyInvocation.MyCommand.Definition


    $origfilename = $siteInitials + "-Phones.csv"
    $newfilename = $siteInitials + "-Phones.csv"
    $filename = $siteInitials + "-Phones-Diff.csv"

    $origfilename = Join-Path $ScriptRootPath "originals\$origfilename"
    $newfilename = join-path $ScriptRootPath "returned\$newfilename"
    $filename = Join-Path $ScriptRootPath "Diffd\$filename"

$original = get-content $origfilename | % {$_ -replace '"',"" } 
$new = get-content $newfilename | % {$_ -replace '"',"" } 




Set-Content -path $filename -Value $original[0]
Compare-Object $original $new | Where {$_.SideIndicator -eq '=>'}  | ForEach-Object {$_.InputObject} | Add-Content $filename