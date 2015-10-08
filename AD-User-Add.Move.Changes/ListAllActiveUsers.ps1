 ###################################
 # Lists all active users in the SMUSD Domain that have real email addresses, and are not disabled
 #
 # Written by Micah Boggs. micah.boggs@gmail.com
 ################

 
 ##### Region Module Import ########

Import-module ActiveDirectory

##### End Region ###########
 
 
 ########## Region Configuration #########
 
 $version = "1.0"


 #scriptpath
 $ScriptRootPath = Split-Path -parent $MyInvocation.MyCommand.Definition
 $CSVFile = Join-Path $ScriptRootPath 'AllUsers.csv'

 # OU names in active directory
 $sitearray = @("ad","atp","ces","cns","dis","disabled users","do","dps","FHS","JALE","KH","LCM","M&O","MHHS","MultiSite","PAL","RL","SEES","SEMS","SMES","SMHS","SMMS","Tech","toe","tohs","wpms")


 ############# End Region Configuration ###############


 ############ Region Execute ################
 Write-host "ListAllActiveUsers.ps1 v$version"
 foreach ($site in $sitearray) {
    
    $allusers += get-aduser -filter {(emailaddress -ne "donotsync@smusd.org") -and (enabled -eq $true) -and (distinguishedname -like "*") } -SearchBase "ou=$site,ou=smusd,dc=smusd,dc=local" -properties department, company, emailaddress, title, LastLogonDate, created, givenname, surname
    Write-Host -NoNewline "."
 }


 $allusers | select name, givenname, surname, emailaddress, title, department, company, LastLogonDate, created | Sort-Object Name | export-csv $CSVFile



 Remove-Variable allusers

 ############ End Region Execute ##########
