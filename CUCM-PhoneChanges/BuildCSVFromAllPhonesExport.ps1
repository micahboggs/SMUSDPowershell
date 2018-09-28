
<# 
#Code from somebody to eliminate duplicate headers.

Function Import-CSVCustom ($csvTemp) {
    $StreamReader = New-Object System.IO.StreamReader -Arg $csvTemp
    [array]$Headers = $StreamReader.ReadLine() -Split "," | % { "$_".Trim() } | ? { $_ }
    $StreamReader.Close()

    $a=@{}; $Headers = $headers|%{
        if($a.$_.count) {"$_$($a.$_.count)"} else {$_}
        $a.$_ += @($_)
    }

    Import-Csv $csvTemp -Header $Headers
}

#>





$csvfile = "\\do-fs\staff\mboggs\desktop\smusdpowershell\CUCM-PhoneChanges\allphones.csv"
$fields = ("location", "Device Name", "Device Type", "Directory Number 1", "Display 1", "Voice Mail Profile 1", "Owner User ID")
$conditions = '$_."device type" -like "Cisco*" -and $_."device name" -like "SEP*"'

remove-variable allphones, locations, location -ErrorAction Ignore

$allphones = import-csv $csvfile | select $fields | where-object {Invoke-Expression $conditions} 
$locations = $allphones | select location | sort location | Get-Unique -AsString
foreach ($location in $locations.location) {
    $locationCSV = @()

    $sitephones = $allphones | ? location -like $location
    foreach ($phone in $sitephones) {
        $out = '' | SELECT location,  extension, Email, displayname, device, type, voicemail
        remove-variable aduser, email, extension -ErrorAction Ignore
        try { 
            $extension = $phone."directory number 1" #| foreach-object {$_.substring($_.length-4)}
        } 
        catch {
            write-warning $phone."device name"
            write-warning $_
            $extension = $null
        }
        if ($phone."owner User ID") {
            $aduser = get-aduser $phone."owner user id" -properties emailaddress
            if ($aduser.emailaddress) {
                $email = $aduser.emailaddress
                if ($email -eq "donotsync@smusd.org") { 
                    Remove-Variable email
                }
                $firstname = $aduser.givenname
                $lastname = $aduser.surname
                $displayname = "$firstname $lastname"
            } 

        }else {
            $displayname = $phone.'Display 1'
        }
        $out.location = $phone.location
        $out.extension = $extension
        $out.email = $email
        $out.displayname = $displayname
        $out.device = $phone."device name"
        $out.type = $phone."device type"
        $out.voicemail = $phone."voice mail profile 1"
        $locationCSV += $out

    }


    $csvoutputfilname = $location + "-phones.csv"
    $csvoutputfile = join-path "\\do-fs\staff\mboggs\desktop\smusdpowershell\phones" $csvoutputfilname
    $locationCSV | sort extension |  export-csv $csvoutputfile -NoTypeInformation
}
 


#$number = "8083924","2943"

#$number | foreach-object {$_.substring($_.length-4)}





#Import-Csv $csvfile | select location, "Device Name","Device Type", "Directory Number 1", "Display 1", "Voice Mail Profile 1", "Owner User ID" | where "device name" -like "SEP*" | where "device type" -like "Cisco*" | Sort-Object "location","Device Type" | format-table