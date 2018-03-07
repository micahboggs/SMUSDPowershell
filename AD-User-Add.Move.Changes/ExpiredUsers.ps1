#################################
# SMUSD Expired Users Script
# Written by Micah Boggs (micah.boggs@gmail.com)
#
# Used to export list of expired users to csv
#
#################################

##### Region Module Import ########

Import-module ActiveDirectory

##### End Region ###########
    #scriptpath
    $ScriptRootPath = Split-Path -parent $MyInvocation.MyCommand.Definition
    #results output
    $ResultsFile = Join-Path $ScriptRootPath 'ExpiredUsers.csv'

search-adaccount -accountexpired | select samaccountname | export-csv $resultsfile -NoTypeInformation