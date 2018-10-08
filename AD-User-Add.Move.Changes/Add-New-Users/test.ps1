

function test {
    return @("maribel.palacios@smusd.org", "AlvinDunnElemSiteAdmins@smusd.org")

}

$a = test
$a -is [array]

