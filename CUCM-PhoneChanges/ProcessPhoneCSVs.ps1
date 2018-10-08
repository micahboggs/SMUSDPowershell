###



# Region Configuration
$ScriptRootPath = Split-Path -parent $MyInvocation.MyCommand.Definition

# End Region

# Region Import Modules

Import-module ActiveDirectory

# End Region


# Region Functions

function masklookup {
param(
    [string]$phoneextension
    )
    $firstdigit = $phoneextension.substring(0,1)
    if ($firstdigit -eq '2' ) {
        return '7602902XXX'
    } elseif ($firstdigit -eq 8) {
        $firstthree = $phoneextension.substring(0,3)
        switch($firstthree)
        {
            ("801") { return '7602902200' }
            ("802") { return '7602902700' }
            ("803") { return '7602902555' }
            ("804") { return '7602902544' }
            ("805") { return '7602902500' }
            ("806") { return '7602902800' }
            ("807") { return '7602902455' }
            ("808") { return '7602902340' }
            ("810") { return '7602902199' }
            ("811") { return '7602902121' }
            ("812") { return '7602902400' }
            ("813") { return '7602902080' }
            ("814") { return '7602902588' }
            ("815") { return '7602902430' }
            ("816") { return '7602902000' }
            ("817") { return '7602902888' }
            ("818") { return '7602902600' }
            ("819") { return '7602902900' }
            ("821") { return '7602902077' }

        }

    }
}

function MainFunction { #The meat of the program
param(
    [string]$SiteInitials
    )





    #setup arrays for objects for eventual export to CSV
    $ManualCSV = @()
    $ProcessedCSV = @()
    $Results = @()


    foreach ($Phone in $input) {
        Write-Host -NoNewline "."
    
        #Reset the failures or set if first one
        $Failures = @()

        $telephoneNumber = $Phone."extension"
        $shortlinenumber = $telephoneNumber.substring($telephoneNumber.length-4)
    
        $name = $Phone."displayname"
        $namesplit = $name.split(" ")
        $givenname = $namesplit[0]
        $surname = $namesplit[1]
        $email = $Phone."email"
        


        #Check for users with old phone number and clear them.
        $OldPhoneUser = Get-ADUser -Filter {(officephone -eq $telephoneNumber)}
        if (($OldPhoneUser | measure).count -ne 0) {
            if (-not ($OldPhoneUser.givenname -eq $namesplit[0] -and $OldPhoneUser.surname -eq $namesplit[1])) { #Make sure that it isn't the same name as before
                try
                {
                    $oldsamaccounts = $oldphoneuser.samaccountname
                    foreach ($i in $oldsamaccounts) {
                        Set-ADUser $i -officephone $null -clear ipphone -ErrorAction Stop
                    }
                }
                catch {

                    $writeerror = "Failed to remove telephone number and/or ipphone from $OldPhoneUser.SamAccountName"
                    Write-Error "$writeerror  -  $_"
                    $Failures += $writeerror + ' - ' + $_.tostring()
                    Remove-Variable writeerror
                }
            }
        }

        #for the next part, I need to check the CUCM Location to see if it is in the active directory DN, but some don't match, so switch it up
        switch($Phone.LOCATION)
        {
            ("LMA")
                {
                    $ADLocation = "La Mirada Academy"
                    break
                }
            ("CES")
                {
                    $ADLocation = "Carrillo Elementary"
                    break
                }
            ("DPS")
                {
                    $ADLocation = "Double Peak"
                    break
                }
            ("DIS")
                {
                    $ADLocation = "Discovery Elementary"
                    break
                }
            ("FHS")
                {
                    $ADLocation = "Foothills High"
                    break
                }
            ("JAES")
                {
                    $ADLocation = "Joli Ann Elementary"
                    break
                }
            ("KH")
                {
                    $ADLocation = "Knob Hill Elementary"
                    break
                }
            ("LCM")
                {
                    $ADLocation = "La Costa Meadows Elementary"
                    break
                }
            ("MHHS")
                {
                    $ADLocation = "Mission Hills High"
                    break
                }
            ("PAL")
                {
                    $ADLocation = "Paloma Elementary"
                    break
                }
            ("RL")
                {
                    $ADLocation = "Richland Elementary"
                    break
                }
            ("SEES")
                {
                    $ADLocation = "San Elijo Elementary"
                    break
                }
            ("SEMS")
                {
                    $ADLocation = "San Elijo Middle"
                    break
                }
            ("SME")
                {
                    $ADLocation = "San Marcos Elementary"
                    break
                }
            ("SMMS")
                {
                    $ADLocation = "San Marcos Middle"
                    break
                }
            ("SMHS")
                {
                    $ADLocation = "San Marcos High"
                    break
                }
            ("TOES")
                {
                    $ADLocation = "Twin Oaks Elementary"
                    break
                }
            ("TOHS")
                {
                    $ADLocation = "Twin Oaks High"
                    break
                }
            ("WPMS")
                {
                    $ADLocation = "Woodland Park Middle"
                    break
                }
        }





        $phonemask = masklookup $telephoneNumber


        #Need this to build the CSVs
        $Out = '' | Select-Object MODEL, "DEVICE NAME", <# DESCRIPTION, "DEVICE POOL",#> LOCATION, "OWNER USER ID", "USER ID", "DIRECTORY NUMBER 1", "DISPLAY 1", "LINE TEXT LABEL 1", "ASCII ALERTING NAME 1", "ALERTING NAME 1", "EXTERNAL PHONE NUMBER MASK 1", "VOICE MAIL PROFILE 1"
        $Out.Model = $Phone."Type"
        $Out."DEVICE NAME" = $Phone."DEVICE"
        # $Out.DESCRIPTION = $Phone.DESCRIPTION  # Don't really care about description, Delete it?
        # $Out."DEVICE POOL" = $Phone."DEVICE POOL" # need to build this from location, or do I not need it?
        $Out.LOCATION = $Phone.LOCATION
        $Out."DIRECTORY NUMBER 1" = $Phone."extension"
        $Out."LINE TEXT LABEL 1" = "Line 1 - $shortlinenumber" 
        $Out."EXTERNAL PHONE NUMBER MASK 1" = $phonemask # not changing this... Can I import without this set?
        $Out."VOICE MAIL PROFILE 1" = $Phone."voicemail"
        


             



        if ($email) { # Try to match on the email before trying firstname lastname as it's more accurate.
            $NewPhoneUser = Get-ADUser -Filter {(emailaddress -eq $email)} -Properties officephone, ipphone
        } 
        
        if ((-not $NewPhoneUser) -and ($surname)) { #No user found yet, try firstname lastname at the given site.
            $NewPhoneUser = Get-ADUser -Filter {(givenname -eq $givenname) -and (surname -eq $surname) -and (company -eq $ADLocation)} -Properties officephone, ipphone
        }

        if (($NewPhoneUser | measure).count -eq 1) { #One Unique User Found
            try
            {
                Set-ADUser $NewPhoneUser.SamAccountName -clear ipphone -ErrorAction Stop  #clear the ipphone field first just in case it has something weird in it
                Set-ADUser $NewPhoneUser.SamAccountName -officephone $telephoneNumber -Add @{ipphone="Unity"}
                
                $givenname = $NewPhoneUser.Givenname
                $surname = $NewPhoneUser.surname
                $displayname = $givenname + " " + $surname

                $Out."DISPLAY 1" = $displayname
                $Out."ASCII ALERTING NAME 1" = $displayname
                $Out."ALERTING NAME 1" = $displayname

                $Out."OWNER USER ID" = $NewPhoneUser.SamAccountName
                $Out."USER ID" = $NewPhoneUser.SamAccountName
                if ($ADUpdateFailure) { 
                    $ManualCSV += $Out 
                } else {
                    $ProcessedCSV += $Out
                }
            }
            catch {
                $writeerror = "Failed to update Active Directory for  $NewPhoneUser.SamAccountName extension: $telephoneNumber"
                Write-Error $writeerror + ' - ' + $_
                $Failures += $writeerror + ' - ' + $_.tostring()
                Remove-Variable writeerror
                $ADUpdateFailure = $true
            }
        } else { #either nobody found, or too many found.
            #output data to manualcsv object
            $Out."DISPLAY 1" = $phone.displayname
            $Out."ASCII ALERTING NAME 1" = $phone.displayname
            $Out."ALERTING NAME 1" = $phone.displayname
            $ManualCSV += $Out
        }








        if (-not $Failures) {
            $Status = "Success"
        } else {
            $Status = "Failed"
        }

        $Out = "" | Select-Object PhoneMac, Status, Warnings
        $Out.PhoneMac = $Phone."DEVICE" 
        $Out.Status = $Status
        $Out.Warnings = $Failures -join ";"
        $Results += $Out


        #remove old variables so they don't bite me later.
        Remove-Variable telephoneNumber, OldPhoneUser, name, namesplit, ADlocation, NewPhoneUser, givenname, surname, ADUpdateFailure -ErrorAction SilentlyContinue
    }
  

   $ResultsFile = $SiteInitials + "-Results.csv"
   $ResultsFile =  Join-Path $ScriptRootPath "Results\$ResultsFile"


   $processedmodels = $ProcessedCSV | select "model" -ExpandProperty model | Get-Unique -AsString 
   foreach ($i in $processedmodels) {
        $modeltrimmed = $i.replace(' ','').replace('Cisco','')
        $ProcessedCSV | where-object {$_.model -eq $i} | select -property * -ExcludeProperty model | export-csv (join-path $ScriptRootPath "Processed\$SiteInitials-$modeltrimmed-Processed.csv") -NoTypeInformation
        
   }
   $manualmodels = $manualCSV | select "model" -ExpandProperty model | Get-Unique -AsString 
   foreach ($i in $manualmodels) {
        $modeltrimmed = $i.replace(' ','').replace('Cisco','')
        $ManualCSV | where-object {$_.model -eq $i} | select -property * -ExcludeProperty model | export-csv (join-path $ScriptRootPath "Processed\$SiteInitials-$modeltrimmed-Manual.csv") -NoTypeInformation    
   }
   $Results | Export-Csv $ResultsFile -NoTypeInformation

    

   <# 
   # Write paths to screen
   write-host "Script Root: $ScriptRootPath"
   $ManualCSVFilename
   $ProcessedFilename
   $ResultsFile
   #>
}




# End Region


# Region Execution


    $siteInitials = Read-Host -prompt "Please enter the site Abbreviation you used in the filenames"
    $siteInitials = $siteInitials.ToUpper()
    #$SiteInitials = "JAES" # For testing so I don't have to type the damn thing in every time

    $filename = $SiteInitials + "-Phones-Diff.csv"
    $filename = Join-Path $ScriptRootPath "Diffd\$filename"


$filename | Import-Csv | MainFunction $SiteInitials 

# Read-Host -Prompt "Press enter to finish..."

# End Region