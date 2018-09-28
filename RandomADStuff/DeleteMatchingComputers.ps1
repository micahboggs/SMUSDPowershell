# DeleteMatchingComputers.ps1
#
# Reads file 'computers.txt' and deletes computers with partial name match. Will ask for confirmation by default.


$list = get-content computers.txt
foreach ( $line in $list ) {
    get-adcomputer -filter "name -like '*$line*'" | Remove-ADComputer -confirm
} 