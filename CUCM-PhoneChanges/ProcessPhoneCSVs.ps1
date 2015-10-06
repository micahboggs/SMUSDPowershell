###



# Region Configuration
$ScriptRootPath = Split-Path -parent $MyInvocation.MyCommand.Definition

# End Region

# Region Import Modules

Import-module ActiveDirectory

# End Region


# Region Functions



function MainFunction { #The meat of the program
param(
    [string]$SiteInitials
    )



    $ManualCSVFilename =  $SiteInitials + "-Manual.csv"
    $ProcessedFilename =  $SiteInitials + "-Processed.csv"
    $ResultsFile = $SiteInitials + "-Results.csv"

    $ManualCSVFilename =  Join-Path $ScriptRootPath $ManualCSVFilename
    $ProcessedFilename =  Join-Path $ScriptRootPath $ProcessedFilename
    $ResultsFile =  Join-Path $ScriptRootPath $ResultsFile


    #setup arrays for objects for eventual export to CSV
    $ManualCSV = @()
    $ProcessedCSV = @()
    $Results = @()


    foreach ($Phone in $input) {
        Write-Host -NoNewline "."
    
        #Reset the failures or set if first one
        $Failures = @()

        $telephoneNumber = $Phone."DIRECTORY NUMBER 1"
    
        $name = $Phone."DISPLAY 1"
        $namesplit = $name.split(" ")
        $givenname = $namesplit[0]
        $surname = $namesplit[1]

        #Check for users with old phone number and clear them.
        $OldPhoneUser = Get-ADUser -Filter {(officephone -eq $telephoneNumber)}
        if (($OldPhoneUser | measure).count -ne 0) {
            if (-not ($OldPhoneUser.givenname -eq $namesplit[0] -and $OldPhoneUser.surname -eq $namesplit[1])) { #Make sure that it isn't the same name as before
                try
                {
                    Set-ADUser $OldPhoneUser.SamAccountName -officephone $null -clear ipphone -ErrorAction Stop
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
            ("AD")
                {
                    $ADLocation = "Alvin Dunn Elementary"
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

        #Need this to build the CSVs
        $Out = '' | Select-Object MODEL, "DEVICE NAME", DESCRIPTION, "DEVICE POOL", LOCATION, "OWNER USER ID", "USER ID", "DIRECTORY NUMBER 1", "DISPLAY 1", "LINE TEXT LABEL 1", "ASCII ALERTING NAME 1", "ALERTING NAME 1", "EXTERNAL PHONE NUMBER MASK 1", "VOICE MAIL PROFILE 1"
        $Out.Model = $phone.MODEL
        $Out."DEVICE NAME" = $Phone."DEVICE NAME"
        $Out.DESCRIPTION = $Phone.DESCRIPTION 
        $Out."DEVICE POOL" = $Phone."DEVICE POOL"
        $Out.LOCATION = $Phone.LOCATION
        $Out."DIRECTORY NUMBER 1" = $Phone."DIRECTORY NUMBER 1"
        $Out."DISPLAY 1" = $PHONE."DISPLAY 1"
        $Out."LINE TEXT LABEL 1" = $Phone."LINE TEXT LABEL 1"
        $Out."ASCII ALERTING NAME 1" = $PHONE."DISPLAY 1"
        $Out."ALERTING NAME 1" = $PHONE."DISPLAY 1"
        $Out."EXTERNAL PHONE NUMBER MASK 1" = $Phone."EXTERNAL PHONE NUMBER MASK 1"
        $Out."VOICE MAIL PROFILE 1" = $Phone."VOICE MAIL PROFILE 1"


        #ok, get the ad user that matches firstname, lastname, and has the location the same as the DN, must only be one match.
        if ($surname) {
            $NewPhoneUser = Get-ADUser -Filter {(givenname -eq $givenname) -and (surname -eq $surname) -and (company -eq $ADLocation)} -Properties officephone, ipphone
            if (($NewPhoneUser | measure).count -eq 1) { #One Unique User Found
                
                
                
                try
                {
                   
                    Set-ADUser $NewPhoneUser.SamAccountName -clear ipphone -ErrorAction Stop  #clear the ipphone field first just in case it has something weird in it
                    Set-ADUser $NewPhoneUser.SamAccountName -officephone $telephoneNumber -Add @{ipphone="Unity"}
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

                $ManualCSV += $Out
            }
        } else {
            #no surname, not an actual user. output data to manual csv object

            $ManualCSV += $Out
         
        }
        if (-not $Failures) {
            $Status = "Success"
        } else {
            $Status = "Failed"
        }

        $Out = "" | Select-Object PhoneMac, Status, Warnings
        $Out.PhoneMac = $Phone."DEVICE NAME" 
        $Out.Status = $Status
        $Out.Warnings = $Failures -join ";"
        $Results += $Out


        #remove old variables so they don't bite me later.
        Remove-Variable telephoneNumber, OldPhoneUser, name, namesplit, ADlocation, NewPhoneUser, givenname, surname, ADUpdateFailure -ErrorAction SilentlyContinue
    }
   # $manualCSV | Format-Table
   # $processedCSV | Format-Table

   $manualCSV | export-csv $ManualCSVFilename
   $ProcessedCSV | Export-Csv  $ProcessedFilename 
   $Results | Export-Csv $ResultsFile 

    


   write-host "Script Root: $ScriptRootPath"
   $ManualCSVFilename
   $ProcessedFilename
   $ResultsFile
}




# End Region


# Region Execution


    $siteInitials = Read-Host -prompt "Please enter the site Abbreviation you used in the filenames"
    #$SiteInitials = "JAES" # For testing so I don't have to type the damn thing in every time

    $filename = $SiteInitials + "Diff.csv"
    $filename = Join-Path $ScriptRootPath $filename


$filename | Import-Csv | MainFunction $SiteInitials 

Read-Host -Prompt "Press enter to finish..."

# End Region