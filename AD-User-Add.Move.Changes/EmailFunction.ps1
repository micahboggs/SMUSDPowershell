

function get-emailaddresses {
param (
    [string]$emaillist=$null
    )

    switch ($emaillist) {

        "TestEmailAddress" { return @("micah.boggs@smusd.org") }
        "EmailFrom" { return 'IT Department <noreply@smusd.org>' }
        "helpdeskEmail" { return 'helpdesk@smusd.org' }
        "Helpdeskemailfrom" { return 'sami.belmonte@smusd.org' }

        "EmailCC" {return @("hector.lopez@smusd.org", "alek.torres-reyes@smusd.org", "kitty.ross@smusd.org", "janine.clark@smusd.org", "tony.cabral@smusd.org", "stephanie.casperson@smusd.org", "sami.belmonte@smusd.org", "micah.boggs@smusd.org" ,"hector.velarde@smusd.org", "daniel.perez@smusd.org") }
        "ADEmail" { return @("maribel.palacios@smusd.org", "AlvinDunnElemSiteAdmins@smusd.org") }
        "CESEmail" { return @("jennifer.smith@smusd.org", "CarrilloElemSiteAdmins@smusd.org") }
        "DISEmail" { return @("Terri.pecoraro@smusd.org", "DiscoveryElemSiteAdmins@smusd.org") }
        "DPSEmail" { return @("maria.ortiz@smusd.org", "DoublePeakSiteAdmins@smusd.org") }
        "FHSEmail" { return @("joanne.wimsatt@smusd.org", "FHSiteAdmins@smusd.org") }
        "JAESEmail" { return @("pat.walker@smusd.org", "JoliAnnSiteAdmins@smusd.org") }
        "KHEmail" { return @("maria.torres@smusd.org", "KnobHillElemSiteAdmins@smusd.org") }
        "LCMEmail" { return @("heather.cooper@smusd.org", "LaCostaMeadowsElemSiteAdmins@smusd.org") }
        "MHHSEmail" { return @("pat.rodriguez-myers@smusd.org", "MissionHillsHSSiteAdmins@smusd.org", "Denise.Le@smusd.org") }
        "MOEmail" { return @("debbie.keenan@smusd.org", "MOSiteAdmins@smusd.org") }
        "PALEmail" { return @("vivian.brix@smusd.org", "PalomaElemSiteAdmins@smusd.org") }
        "RLEmail" { return @("gaby.dellamary@smusd.org", "RichlandElemSiteAdmins@smusd.org") }
        "SEESEmail" { return @("Jessica.raya@smusd.org", "SanElijoESSiteAdmins@smusd.org") }
        "SEMSEmail" { return @("shelly.valentine@smusd.org", "SanElijoMSSiteAdmins@smusd.org") }
        "SMESEmail" { return @("lupe.escobedo@smusd.org", "SanMarcosElemSiteAdmins@smusd.org") }
        "SMMSEmail" { return @("karen.margis@smusd.org", "SanMarcosMSSiteAdmins@smusd.org") }
        "SMHSEmail" { return @("enda.davis@smusd.org", "SanMarcosHSSiteAdmins@smusd.org") }
        "TOESEmail" { return @("liliana.garcia@smusd.org", "TwinOaksElemSiteAdmins@smusd.org") }
        "TOHSEmail" { return @("joanne.wimsatt@smusd.org", "TwinOaksHSSiteAdmins@smusd.org") }
        "WPMSEmail" { return @("debra.weaver@smusd.org", "WoodlandParkMSSiteAdmins@smusd.org") }
        "DOEmail" { return @("DistrictOfficeSiteAdmins@smusd.org") }
        "CNSEmail" { return @("cns-management@smusd.org", "DistrictOfficeSiteAdmins@smusd.org") }
        "KOCEmail" { return @("teigynn.knight@smusd.org", "DistrictOfficeSiteAdmins@smusd.org") }

        default { return $null }
    }
}


