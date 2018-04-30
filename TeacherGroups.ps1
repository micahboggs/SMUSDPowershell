<#
TeacherGroups.ps1

Adds teachers to appropriate grade level email groups.

Written by Micah Boggs
micah.boggs@smusd.org
#>

############# Configuration ############

$CSVFilename = "Teacher_Groups.csv" #name of the import file

$PSEmailServer = 'smusd-relay.smusd.local'  # fqdn of the smtp server
$Fromemail = '"Script Output" <noreply@smusd.org>' # Email address script sends from

[int]$daystokeepprocessed = "30"  # Number of days to keep CSVs that have already been processed.

$logfilename ="TeacherGroups.log"  # Log Filename
$logLevel = "INFO" # Default logging level. ("DEBUG","INFO","WARN","ERROR","FATAL")
$logSize = 1mb # Maximum size of log file before rotation
$logCount = 10 # Number of log files to keep before deletion.

########## End Configuration ##########

##### Region Module Import ########

Import-module ActiveDirectory

##### End Region ###########





######## Functions #######




function Write-Log-Line ($line) {
    Add-Content $logFile -Value $Line
    Write-Host $Line
}

# http://stackoverflow.com/a/38738942
Function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$True)]
    [string]
    $Message,
    
    [Parameter(Mandatory=$False)]
    [String]
    $Level = "INFO"
    )

    $levels = ("DEBUG","INFO","WARN","ERROR","FATAL")
    $logLevelPos = [array]::IndexOf($levels, $logLevel)
    $levelPos = [array]::IndexOf($levels, $Level)
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss:fff")

    if ($logLevelPos -lt 0){
        Write-Log-Line "$Stamp ERROR Wrong logLevel configuration [$logLevel]"
    }
    
    if ($levelPos -lt 0){
        Write-Log-Line "$Stamp ERROR Wrong log level parameter [$Level]"
    }

    # if level parameter is wrong or configuration is wrong I still want to see the 
    # message in log
    if ($levelPos -lt $logLevelPos -and $levelPos -ge 0 -and $logLevelPos -ge 0){
        return
    }
    if ($levelpos -ge 2) {
        $Script:Failures += "$level $message`n"
    }
    $Line = "$Stamp $Level $Message"
    Write-Log-Line $Line
}

# https://gallery.technet.microsoft.com/scriptcenter/PowerShell-Script-to-Roll-a96ec7d4
function Reset-Log 
{ 
    # function checks to see if file in question is larger than the paramater specified 
    # if it is it will roll a log and delete the oldes log if there are more than x logs. 
    param([string]$fileName, [int64]$filesize = 1mb , [int] $logcount = 5) 
     
    $logRollStatus = $true 
    if(test-path $filename) 
    { 
        $file = Get-ChildItem $filename 
        if((($file).length) -ige $filesize) #this starts the log roll 
        { 
            $fileDir = $file.Directory 
            #this gets the name of the file we started with 
            $fn = $file.name
            $files = Get-ChildItem $filedir | ?{$_.name -like "$fn*"} | Sort-Object lastwritetime 
            #this gets the fullname of the file we started with 
            $filefullname = $file.fullname
            #$logcount +=1 #add one to the count as the base file is one more than the count 
            for ($i = ($files.count); $i -gt 0; $i--) 
            {  
                #[int]$fileNumber = ($f).name.Trim($file.name) #gets the current number of 
                # the file we are on 
                $files = Get-ChildItem $filedir | ?{$_.name -like "$fn*"} | Sort-Object lastwritetime 
                $operatingFile = $files | ?{($_.name).trim($fn) -eq $i} 
                if ($operatingfile) 
                 {$operatingFilenumber = ($files | ?{($_.name).trim($fn) -eq $i}).name.trim($fn)} 
                else 
                {$operatingFilenumber = $null} 
 
                if(($operatingFilenumber -eq $null) -and ($i -ne 1) -and ($i -lt $logcount)) 
                { 
                    $operatingFilenumber = $i 
                    $newfilename = "$filefullname.$operatingFilenumber" 
                    $operatingFile = $files | ?{($_.name).trim($fn) -eq ($i-1)} 
                    write-host "moving to $newfilename" 
                    move-item ($operatingFile.FullName) -Destination $newfilename -Force 
                } 
                elseif($i -ge $logcount) 
                { 
                    if($operatingFilenumber -eq $null) 
                    {  
                        $operatingFilenumber = $i - 1 
                        $operatingFile = $files | ?{($_.name).trim($fn) -eq $operatingFilenumber} 
                        
                    } 
                    write-host "deleting " ($operatingFile.FullName) 
                    remove-item ($operatingFile.FullName) -Force 
                } 
                elseif($i -eq 1) 
                { 
                    $operatingFilenumber = 1 
                    $newfilename = "$filefullname.$operatingFilenumber" 
                    write-host "moving to $newfilename" 
                    move-item $filefullname -Destination $newfilename -Force 
                } 
                else 
                { 
                    $operatingFilenumber = $i +1  
                    $newfilename = "$filefullname.$operatingFilenumber" 
                    $operatingFile = $files | ?{($_.name).trim($fn) -eq ($i-1)} 
                    write-host "moving to $newfilename" 
                    move-item ($operatingFile.FullName) -Destination $newfilename -Force    
                } 
            } 
          } 
         else 
         { $logRollStatus = $false} 
    } 
    else 
    { 
        $logrollStatus = $false 
    } 
    $LogRollStatus 
} 


function groupselect {
param ($GradeOrSubject)

    switch ($GradeOrSubject) {    
        "01" { $group="SMUSD-1stGrade-Teachers" }
        "02" { $group="SMUSD-2ndGrade-Teachers" }
        "03" { $group="SMUSD-3rdGrade-Teachers" }
        "04" { $group="SMUSD-4thGrade-Teachers" }
        "05" { $group="SMUSD-5thGrade-Teachers" }
        "English" { $group="SMUSD-English-Teachers" }
        "Foreign Language" { $group="SMUSD-ForeignLanguage-Teachers" }
        "History" { $group="SMUSD-History-Teachers" }
        "KN" { $group="SMUSD-Kindergarten-Teachers" }
        "Mathematics" { $group="SMUSD-Math-Teachers" }
        "Science" { $group="SMUSD-Science-Teachers" }
        "TK" { $group="SMUSD-TK-Teachers" }
        default { return $null }
    }
    return $group
}


function emailerrors {
param ($failures)

    
    if ($failures) {
        $emailto = get-content (join-path $script:rootpath "/script/EmailTo.txt")
        $subject = "Teacher Group Script Issues"
        $body = @"
The TeacherGroups.ps1 script has generated the below errors. 

$failures
"@    
        
        try {
            Send-MailMessage -To $emailto -Body $body -From $Script:FromEmail -Subject $subject
            write-log "Sent email to $emailto"
        }
        catch {
            write-log "Unable to send email: $_" "ERROR"
        }
    }
}


function AddGroupMember {
param ($group, $user)
    try {
        add-adgroupmember $group -members $user
        write-log "Added $user to group $group" 
    }
    catch {
        write-log "Unable to add $user to $group : $_" "ERROR"
    }
}

function RemoveGroupMember {
param ($group, $user)
    try {
        remove-adgroupmember $group -members $user -confirm:$false
        write-log "Removed $user from group $group" 
    }
    catch {
        write-log "Unable to remove $user from $group : $_" "ERROR"
    }
}


function main {
param ($list, $gradeorsubject)
    
    
    $uniquelist = $list | select t_email | Sort-Object T_EMail | Get-Unique -AsString


    
    $groupname = groupselect $gradeorsubject
    try {
        $ADGroup = get-adgroupmember $groupname
    }
    catch {
        #Unexpected subject. Log error. then return

        write-log "Unknown Subject `"$gradeorsubject`": $_" "ERROR"

        return
    }
    $SamAccountList = @()
    foreach ($user in $uniquelist) {
        $emailaddress = $user.T_Email
        if (!($emailaddress -like '*_*')) {
            try {
                $userAccount = get-aduser -filter 'emailaddress -like $emailaddress' -properties emailaddress
            }
            catch {
                #cannot find useraccount with specified email address
                Write-log "No User with email address $emailaddress, $user" "ERROR"
                continue
            }
            if ($useraccount.measuer -gt 1) {
                #more than one account with specified emailaddress
                Write-Log "More than one user with email $emailaddress, $($useraccount.samaccountname)"
                continue
            } elseif ($useraccount.measure -lt 1) {
                #cannot find useraccount with specified email address
                Write-Log "No User with email address $emailaddress, $user" "ERROR"
                continue
            } else {
                $SamAccountList += $useraccount.SamAccountName
            }
        }
        
    }
    if ($adgroup -and $SamAccountList) {
        $differences = compare-object $adgroup.samaccountname @($SamAccountList | select-object)
        if (!$differences) {
            write-log "CSV and ADGroupMembers for $groupname are the same. No changes applied"
        } else {
            foreach ($difference in $differences) {
                if ($difference.SideIndicator -like '=>') {
                    #User is in CSVList, but not in Group. Add user to group
                    addgroupmember $groupname $difference.InputObject
                } elseif ($difference.SideIndicator -like '<=') {
                    #user is in group, but not in CSVList. Remove user from group
                    RemoveGroupMember $groupname $difference.InputObject
                }
            
            }
        }
    } elseif ($adgroup -and !$SamAccountList) {
        #### No users in csvlist, delete all users from group.
        write-log "CSV lists no users in $groupname, all users will be removed from group." "WARNING"
        foreach ($user in $adgroup.samaccountname) {
            
        }
    } elseif (!$adgroup -and $SamAccountList) {
        #### No users in group, add all csv users to group
        foreach ($user in $SamAccountList) {
            addGroupMember $groupname $user
        }
    }
}



###### Execute #####

$RootPath = split-path -parent (Split-Path -parent $MyInvocation.MyCommand.Definition)
$logfile = join-path (join-path $rootpath "/logs/") $logfilename


# Rotate log if needed.
$Null = @(
    Reset-Log -fileName $logFile -filesize $logSize -logcount $logCount
) 

    


write-log "#########################"
write-log "Script Started"    

#Reset the failures or set if first one
$Failures = @()
$csvfile = join-path $rootpath $CSVFilename
If (-not (Test-Path $CSVFile)){
    #File Doesn't Exist, abort
    Write-Log "$CSVFile does not exist" "FATAL"
    emailerrors $Failures
    write-log "Script Finished"
    write-log "#########################"
    exit
}


$input = import-csv $CSVFile
$GradesAndSubjectsArray = $input | Sort-Object Grade_or_subject | select Grade_or_subject | get-unique -AsString | foreach {"$($_.Grade_or_subject)"}

foreach ($gradeorsubject in $GradesAndSubjectsArray){ 
    #everything but Science as there are different types of science that needs to be treated as a single subject
    if (!($gradeorsubject -like "*science")) {
        main $($input | ?{$_.Grade_or_Subject -eq $gradeorsubject } ) $gradeorsubject
    }
}
#just science as there are multiple Sciences.  Will need to deal with Grade_or_subject not being Science exactly
main $( $input | ?{$_.Grade_or_Subject -like "*science"} ) "Science"


#ok, done, so move csv to processed folder.

$time=Get-Date -format yyyyMMdd.HHmm
$processedFile = join-path $rootpath "\processed\Teacher_Groups.$time.csv"




try {
    move-item -path $CSVFile -Destination $processedFile -Force
    write-log "Moved `"$CSVFile`" to `"$processedFile`""
}
catch {
    write-log "Unable to move `"$CSVFile`" to `"$processedFile`" : $_" "ERROR"
}


$filestodelete = get-childitem "\\do-fs\Department Shares\SynergyExports\Processed" | ?{$_.LastWriteTime -lt (get-date).adddays($daystokeepprocessed * -1)}
foreach ($file in $filestodelete) {
    $fullfilepath = join-path $file.directory $file.Name
    try {
        remove-item $fullfilepath -force
        write-log "Deleted $fullfilepath"
    }
    catch {
        write-log "Unable to delete $fullfilepath : $_"
    }
}




emailerrors $Failures
write-log "Script Finished"
write-log "#########################"

# SIG # Begin signature block
# MIIVGQYJKoZIhvcNAQcCoIIVCjCCFQYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUDTmT2Mt6IAKLzqtcnQ6rpb9M
# JLOgghA7MIIEmTCCA4GgAwIBAgIPFojwOSVeY45pFDkH5jMLMA0GCSqGSIb3DQEB
# BQUAMIGVMQswCQYDVQQGEwJVUzELMAkGA1UECBMCVVQxFzAVBgNVBAcTDlNhbHQg
# TGFrZSBDaXR5MR4wHAYDVQQKExVUaGUgVVNFUlRSVVNUIE5ldHdvcmsxITAfBgNV
# BAsTGGh0dHA6Ly93d3cudXNlcnRydXN0LmNvbTEdMBsGA1UEAxMUVVROLVVTRVJG
# aXJzdC1PYmplY3QwHhcNMTUxMjMxMDAwMDAwWhcNMTkwNzA5MTg0MDM2WjCBhDEL
# MAkGA1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4GA1UE
# BxMHU2FsZm9yZDEaMBgGA1UEChMRQ09NT0RPIENBIExpbWl0ZWQxKjAoBgNVBAMT
# IUNPTU9ETyBTSEEtMSBUaW1lIFN0YW1waW5nIFNpZ25lcjCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAOnpPd/XNwjJHjiyUlNCbSLxscQGBGue/YJ0UEN9
# xqC7H075AnEmse9D2IOMSPznD5d6muuc3qajDjscRBh1jnilF2n+SRik4rtcTv6O
# KlR6UPDV9syR55l51955lNeWM/4Og74iv2MWLKPdKBuvPavql9LxvwQQ5z1IRf0f
# aGXBf1mZacAiMQxibqdcZQEhsGPEIhgn7ub80gA9Ry6ouIZWXQTcExclbhzfRA8V
# zbfbpVd2Qm8AaIKZ0uPB3vCLlFdM7AiQIiHOIiuYDELmQpOUmJPv/QbZP7xbm1Q8
# ILHuatZHesWrgOkwmt7xpD9VTQoJNIp1KdJprZcPUL/4ygkCAwEAAaOB9DCB8TAf
# BgNVHSMEGDAWgBTa7WR0FJwUPKvdmam9WyhNizzJ2DAdBgNVHQ4EFgQUjmstM2v0
# M6eTsxOapeAK9xI1aogwDgYDVR0PAQH/BAQDAgbAMAwGA1UdEwEB/wQCMAAwFgYD
# VR0lAQH/BAwwCgYIKwYBBQUHAwgwQgYDVR0fBDswOTA3oDWgM4YxaHR0cDovL2Ny
# bC51c2VydHJ1c3QuY29tL1VUTi1VU0VSRmlyc3QtT2JqZWN0LmNybDA1BggrBgEF
# BQcBAQQpMCcwJQYIKwYBBQUHMAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20w
# DQYJKoZIhvcNAQEFBQADggEBALozJEBAjHzbWJ+zYJiy9cAx/usfblD2CuDk5oGt
# Joei3/2z2vRz8wD7KRuJGxU+22tSkyvErDmB1zxnV5o5NuAoCJrjOU+biQl/e8Vh
# f1mJMiUKaq4aPvCiJ6i2w7iH9xYESEE9XNjsn00gMQTZZaHtzWkHUxY93TYCCojr
# QOUGMAu4Fkvc77xVCf/GPhIudrPczkLv+XZX4bcKBUCYWJpdcRaTcYxlgepv84n3
# +3OttOe/2Y5vqgtPJfO44dXddZhogfiqwNGAwsTEOYnB9smebNd0+dmX+E/CmgrN
# Xo/4GengpZ/E8JIh5i15Jcki+cPwOoRXrToW9GOUEB1d0MYwggW9MIIEpaADAgEC
# AhNvAAAAxnGItq0OVJ8XAAAAAADGMA0GCSqGSIb3DQEBBQUAMEkxFTATBgoJkiaJ
# k/IsZAEZFgVsb2NhbDEVMBMGCgmSJomT8ixkARkWBXNtdXNkMRkwFwYDVQQDExBz
# bXVzZC1ETy1EQzAyLUNBMB4XDTE3MTIxMTIxMDIxMloXDTE4MTIwNTAwNDA1MFow
# gZExFTATBgoJkiaJk/IsZAEZFgVsb2NhbDEVMBMGCgmSJomT8ixkARkWBXNtdXNk
# MQ4wDAYDVQQLEwVTTVVTRDENMAsGA1UECxMEVEVDSDEOMAwGA1UECxMFVXNlcnMx
# CzAJBgNVBAsTAklUMQ4wDAYDVQQLEwVNaWNhaDEVMBMGA1UEAxMMQm9nZ3MsIE1p
# Y2FoMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEArs8PbHNuwrG8IkYW
# wl2NFs0cnhCiVLuDo9DHe3Ln1WrGJBBPEEZFdLfiRodWYh5oe2xMf/z9H3rPzsWS
# uj63H6RhfyahiLMe6gyTBMSfmeAkTE0kw5iFeF0DEK1d4WIq+ql5QOkwptGnSC79
# lyul+XWA0F6SeWVXO3zRacrmon6jIRTWHnArytD/y8g5arRGZ3i5k19T1LmPmTQF
# mo3iHToMXUFaQCjoHys2kXOkVm81rlchcR71KI3o/iO+H7kJhqM5tNCVcyZCBszE
# /aO6o/S62oQmTYcNiPHczxgldsP9TR3DbJldTLdRM3jZQY7JhTasKyiJbeSQ9IgE
# 4aNqtwIDAQABo4ICUzCCAk8wJQYJKwYBBAGCNxQCBBgeFgBDAG8AZABlAFMAaQBn
# AG4AaQBuAGcwEwYDVR0lBAwwCgYIKwYBBQUHAwMwDgYDVR0PAQH/BAQDAgeAMB0G
# A1UdDgQWBBSvymSOCU+421ba01RgSZcjRLD8WTAfBgNVHSMEGDAWgBTxukjV5Brr
# JCVfIJmW2SGDTDLhETCBzgYDVR0fBIHGMIHDMIHAoIG9oIG6hoG3bGRhcDovLy9D
# Tj1zbXVzZC1ETy1EQzAyLUNBLENOPURPLURDMDIsQ049Q0RQLENOPVB1YmxpYyUy
# MEtleSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9
# c211c2QsREM9bG9jYWw/Y2VydGlmaWNhdGVSZXZvY2F0aW9uTGlzdD9iYXNlP29i
# amVjdENsYXNzPWNSTERpc3RyaWJ1dGlvblBvaW50MIHCBggrBgEFBQcBAQSBtTCB
# sjCBrwYIKwYBBQUHMAKGgaJsZGFwOi8vL0NOPXNtdXNkLURPLURDMDItQ0EsQ049
# QUlBLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNv
# bmZpZ3VyYXRpb24sREM9c211c2QsREM9bG9jYWw/Y0FDZXJ0aWZpY2F0ZT9iYXNl
# P29iamVjdENsYXNzPWNlcnRpZmljYXRpb25BdXRob3JpdHkwKwYDVR0RBCQwIqAg
# BgorBgEEAYI3FAIDoBIMEG1ib2dnc0BzbXVzZC5vcmcwDQYJKoZIhvcNAQEFBQAD
# ggEBACrzCI5amXQShNCkIMF6FOCZKoG9e910WxLL6TumPW8u0zzBMobjO4lOew1l
# FRQZV+o+GihJWFJFN9plChZAN1y7Bs1rN4K4hzQC/T+Wbdft1NwdTc2vK/UMX6fp
# L/EV9RhuiTQsu58jGO5+U5kWP8ktG7kq0CjAp5ZGHLxdPJALQGLOV+HefIJGBdwU
# sAA6gowWDylKRQQRTmtmSuOi6TNcP+UiAFhw9pVNEzFCyoU2iHA3m3Er0h0E33vc
# PusFHLPLbWyyyfSs68+pP0ocmMfJzKqtg0bEk89lI6BIgVvnkmCq08g2EaOyU5ms
# Ta7K/Knf+yPZ2cbtZucmkxyLrbYwggXZMIIEwaADAgECAgpSySgeAAEAAAAMMA0G
# CSqGSIb3DQEBBQUAMCMxITAfBgNVBAMTGFNNVVNELVN0YW5kYWxvbmUtUm9vdC1D
# QTAeFw0xMzEyMjMyMDA2NDRaFw0xODEyMDUwMDQwNTBaMEkxFTATBgoJkiaJk/Is
# ZAEZFgVsb2NhbDEVMBMGCgmSJomT8ixkARkWBXNtdXNkMRkwFwYDVQQDExBzbXVz
# ZC1ETy1EQzAyLUNBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA7IFw
# qdY/70IMjfLLeQKQZ86blKpvQd0ICzDFBPmq2tamRmCD/pZXdyMyizzKXl6gocvT
# xpdLgfbuPeFOVt9NbqVblCT1tLgZvbPl+MLc0By07PlMoOvVaythEp/BZdebvLt2
# 3prClZeiI5g2tX9lyTvNwEU7KTod1At72Do7bnXN7dIFpUCcXO2b6AMvYJXqSO4p
# Dw9SmeNYng5rcHgElyZgSVVuG7dBQc7cK9cru5feQzqjWKcY3XMPcZy8/xMZO3Qa
# W6vGq8gwPizLbbnT6wfrPu1sJa56SrSyaGwBFSqE29FJX8Qd1LcQ4YPYG+MeTy5Z
# ABcOP+GKvZ+CyAI+WwIDAQABo4IC5zCCAuMwDwYDVR0TAQH/BAUwAwEB/zAdBgNV
# HQ4EFgQU8bpI1eQa6yQlXyCZltkhg0wy4REwCwYDVR0PBAQDAgGGMBAGCSsGAQQB
# gjcVAQQDAgEAMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMB8GA1UdIwQYMBaA
# FK9Om4cqt1xrxno8ENiEJfqgkahBMIIBJAYDVR0fBIIBGzCCARcwggEToIIBD6CC
# AQuGR2h0dHA6Ly9jZXJ0aWZpY2F0ZXMuc211c2QubG9jYWwvQ2VydEVucm9sbC9T
# TVVTRC1TdGFuZGFsb25lLVJvb3QtQ0EuY3JshoG/bGRhcDovLy9DTj1TTVVTRC1T
# dGFuZGFsb25lLVJvb3QtQ0EsQ049RE8tREMwMCxDTj1DRFAsQ049UHVibGljJTIw
# S2V5JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049Q29uZmlndXJhdGlvbixEQz1z
# bXVzZCxkYz1sb2NhbD9jZXJ0aWZpY2F0ZVJldm9jYXRpb25MaXN0P2Jhc2U/b2Jq
# ZWN0Q2xhc3M9Y1JMRGlzdHJpYnV0aW9uUG9pbnQwggEsBggrBgEFBQcBAQSCAR4w
# ggEaMF4GCCsGAQUFBzAChlJodHRwOi8vY2VydGlmaWNhdGVzLnNtdXNkLmxvY2Fs
# L0NlcnRFbnJvbGwvRE8tREMwMF9TTVVTRC1TdGFuZGFsb25lLVJvb3QtQ0EoMSku
# Y3J0MIG3BggrBgEFBQcwAoaBqmxkYXA6Ly8vQ049U01VU0QtU3RhbmRhbG9uZS1S
# b290LUNBLENOPUFJQSxDTj1QdWJsaWMlMjBLZXklMjBTZXJ2aWNlcyxDTj1TZXJ2
# aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPXNtdXNkLGRjPWxvY2FsP2NBQ2VydGlm
# aWNhdGU/YmFzZT9vYmplY3RDbGFzcz1jZXJ0aWZpY2F0aW9uQXV0aG9yaXR5MA0G
# CSqGSIb3DQEBBQUAA4IBAQA27CsxnCPJSiQUEHdfX3wxhH18OqGV0GORtRkZpLRF
# loL12/MSLMHTvFyVBnmyDAoeMk749+eFFQzwuKlMIZcwY8QntwHcIcTgoqM78Hfv
# +VtUfXu0jq3kkOSiN33DRQltdlkEPFCN7N9HU83cefMpc1vA+QwxWUClXvRxhnT0
# vsmKGAbxZImK2WRYDRdNJSLm210bXSan2LH6q8tUy1TJ3HZKFsPXqKFDFefO7jK8
# dSUD/JVcKedHh1SWEUW+6UGV6EtxSPtmZrbRfd3TGCuYXupbBVlONGQi54kbM7o9
# /9TCG2Ej6PXRkVIkbzT16ovjcdpkuqUfEyOVLVlYBWXfMYIESDCCBEQCAQEwYDBJ
# MRUwEwYKCZImiZPyLGQBGRYFbG9jYWwxFTATBgoJkiaJk/IsZAEZFgVzbXVzZDEZ
# MBcGA1UEAxMQc211c2QtRE8tREMwMi1DQQITbwAAAMZxiLatDlSfFwAAAAAAxjAJ
# BgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAj
# BgkqhkiG9w0BCQQxFgQUwvh8l0oiUx/4V+DN1CIHYdU9A9cwDQYJKoZIhvcNAQEB
# BQAEggEAeLFM9zZGcYxdciCE2CU7MDgUdxU1bqb5fWbDrLuygs3IQvnJAzcuDV3g
# rQSilMUc0ZqM8L7QtHqGu141jx7WmlX5/ilWAMDqM8Op/JWf4zh4sanjLj5VpyRq
# RkS8b5LdnbNifNSGHq3A1HZ3cOCeJbXcvsqsbxVLhQ7QUe/KAtV5ltizEeSq3I4C
# +yQ3sEJrk6XZyOmq2P/0QP0FklWfqmKCMvLwRtXhR+Nr0oIqodv7ccw9+mgPYPPH
# owJcfj+0a2VbocZds1pVScHv2P09zvRJPF1KArmzXp3DgbSlGBieKriAX3mfrybk
# jBa6vfYnOqdIDndyLHGWyKTay/m8YaGCAkMwggI/BgkqhkiG9w0BCQYxggIwMIIC
# LAIBATCBqTCBlTELMAkGA1UEBhMCVVMxCzAJBgNVBAgTAlVUMRcwFQYDVQQHEw5T
# YWx0IExha2UgQ2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMSEw
# HwYDVQQLExhodHRwOi8vd3d3LnVzZXJ0cnVzdC5jb20xHTAbBgNVBAMTFFVUTi1V
# U0VSRmlyc3QtT2JqZWN0Ag8WiPA5JV5jjmkUOQfmMwswCQYFKw4DAhoFAKBdMBgG
# CSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE4MDQzMDE1
# NTQzM1owIwYJKoZIhvcNAQkEMRYEFBva33nMrTLDHKMYFFoaiyzMgpUUMA0GCSqG
# SIb3DQEBAQUABIIBAEMhXoXEsvs4W5Wwe1iZcQ/gqUyEqk4KEfJXa/7UP5bQvOFd
# 3B+Ot7KjBBSXKgkz0tABoHoiPCiBw6pvTKxxc6GPkjyiTCMMm/R7H7/p+DrlPZiQ
# dZLvvkTIzRyLM3QGtd2qXxDXqcxmdhoeHgdGgJuTpfUvops0XRh/l+3himHAbkEQ
# OS4Wd6cQPWwrHI2FYgYMlQ3vr2IoK5rG5YJN84du4TziGRzFDHHI441TOW+W+BuR
# BFVTNNoauJ0P7XjMbnQ0f1T7oZRJTTy3tEFSE/zbpREBXzP1PYXVLAzcn1wTpTFi
# 1xFX7KWpYMa9hitnsghqB/6Cy6bukm3dM/TMk4E=
# SIG # End signature block
