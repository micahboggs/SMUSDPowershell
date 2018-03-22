#script to remove old dmp files from VMware Update Manager Logs

# How many days of dump files should be retained
$daystokeep = 30


######### Do not edit below this line ###########
$limit = (get-date).adddays(-$daystokeep)
$path = "C:\ProgramData\VMware\VMware Update Manager\Logs\*.dmp"




# Execute
get-childitem -path $path | Where-Object { $_.LastWriteTime -lt $limit } | Remove-Item -whatif

