#Note: Email Arrays used in this file include:arrays:   $EmailCC, $ADEmail $CESEmail $DISEmail $DPSEmail $FHSEmail $JAESEmail $KHEmail $LCMEmail $MHHSEmail $MOEmail $PALEmail $RLEmail $SEESEmail 
    ##      $SEMSEmail $SMESEmail $SMMSEmail $SMHSEmail $TOESEmail $TOHSEmail $WPMSEmail $DOEmail $KOCEmail $CNSEmail $TestEmailAddress 
    ## They must all be defined in the EmailVariables.ps1 File

function Get-Company {
param ([string]$company, [string]$Title)

    $properties = @{
        templateuser = $false
        addgroups = @()
        emailto = $false
        department = $false
        ou = $false
        companyoverride = $false
    }


    switch($Company) {
        ("Adult Transition Program")
            {
            $properties['templateuser'] = "ATP-TEMPLATE"
                
                if ($Title.contains("Teacher"))
                {
                        
                    $properties['addgroups'] += "ATP Certificated Email"
                } elseif ($Title.contains('Principal')) {
                        
                    $properties['addgroups'] += "ATP Management Email"
                } else {
                         
                    $properties['addgroups'] += "ATP Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "DOEmail"
            }
        ("Alvin Dunn Elementary School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "ad-teach-template"
                    $properties['addgroups'] += "AD Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "ad-teach-template"
                    $properties['addgroups'] += "AD Management Email"
                } else {
                    $properties['templateuser'] = "ad-ss-template" 
                    $properties['addgroups'] += "AD Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "ADEmail"
            }
        ("Carrillo Elementary School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "ces-teacher-template"
                    $properties['addgroups'] += "CAR Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "ces-teacher-template"
                    $properties['addgroups'] += "CAR Management Email"
                }  else {
                    $properties['templateuser'] = "ces-ss-template" 
                    $properties['addgroups'] += "CAR Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "CESEmail"
            }
        ("Double Peak School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "dps-teacher-template"
                    $properties['addgroups'] += "DPS Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "dps-teacher-template"
                    $properties['addgroups'] += "DPS Management Email"
                } else {
                    $properties['templateuser'] = "dps-ss-template" 
                    $properties['addgroups'] += "DPS Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "DPSEmail"

            }
        ("Discovery Elementary School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "dis-teacher-template"
                    $properties['addgroups'] += "DIS Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "dis-teacher-template"
                    $properties['addgroups'] += "DIS Management Email"
                    
                } else {
                    $properties['templateuser'] = "dis-ss-template" 
                    $properties['addgroups'] += "DIS Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "DISEmail"
            }
        ("Foothills High School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "fhs-teacher-template"
                    $properties['addgroups'] += "FH Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "fhs-teacher-template"
                    $properties['addgroups'] += "FH Management Email"
                    
                } else {
                    $properties['templateuser'] = "fhs-ss-template" 
                    $properties['addgroups'] += "FH Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "FHSEmail"
            }
        ("Joli Ann Leichtag Elementary School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "jaes-teacher-templat"
                    $properties['addgroups'] += "JALE Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "jaes-teacher-templat"
                    $properties['addgroups'] += "JALE Management Email"
                    
                } else {
                    $properties['templateuser'] = "jaes-ss-template" 
                    $properties['addgroups'] += "JALE Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "JAESEmail"
            }
        ("Knob Hill Elementary School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "kh-teacher-template"
                    $properties['addgroups'] += "KH Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "kh-teacher-template"
                    $properties['addgroups'] += "KH Management Email"
                    
                } else {
                    $properties['templateuser'] = "kh-ss-template" 
                    $properties['addgroups'] += "KH Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "KHEmail"
            }
        ("La Costa Meadows Elementary School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "lcm-teacher-template"
                    $properties['addgroups'] += "LCM Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "lcm-teacher-template"
                    $properties['addgroups'] += "LCM Management Email"
                    
                } else {
                    $properties['templateuser'] = "lcm-ss-template" 
                    $properties['addgroups'] += "LCM Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "LCMEmail"
            }
        ("Mission Hills High School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "mhhs-teacher-templat"
                    $properties['addgroups'] += "MHHS Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "mhhs-teacher-templat"
                    $properties['addgroups'] += "MHHS Management Email"
                    
                } else {
                    $properties['templateuser'] = "mhhs-ss-template" 
                    $properties['addgroups'] += "MHHS Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "MHHSEmail"
            }
        ("Paloma Elementary School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "pal-teacher-template"
                    $properties['addgroups'] += "PAL Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "pal-teacher-template"
                    $properties['addgroups'] += "PAL Management Email"
                    
                } else {
                    $properties['templateuser'] = "pal-ss-template" 
                    $properties['addgroups'] += "PAL Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "PALEmail"
            }
        ("Richland Elementary School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "rl-teacher-template"
                    $properties['addgroups'] += "RL Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "rl-teacher-template"
                    $properties['addgroups'] += "RL Management Email"
                    
                } else {
                    $properties['templateuser'] = "rl-ss-template" 
                    $properties['addgroups'] += "RL Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "RLEmail"
            }
        ("San Elijo Elementary School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "sees-teacher-templat"
                    $properties['addgroups'] += "SEES Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "sees-teacher-templat"
                    $properties['addgroups'] += "SEES Management Email"
                    
                } else {
                    $properties['templateuser'] = "sees-ss-template" 
                    $properties['addgroups'] += "SEES Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "SEESEmail"
            }
        ("San Elijo Middle School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "sems-teacher-templat"
                    $properties['addgroups'] += "SEMS Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "sems-teacher-templat"
                    $properties['addgroups'] += "SEMS Management Email"
                    
                } else {
                    $properties['templateuser'] = "sems-ss-template" 
                    $properties['addgroups'] += "SEMS Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "SEMSEmail"
            }
        ("San Marcos Elementary School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "smes-teacher-templat"
                    $properties['addgroups'] += "SMES Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "smes-teacher-templat"
                    $properties['addgroups'] += "SMES Management Email"
                    
                } else {
                    $properties['templateuser'] = "smes-ss-template" 
                    $properties['addgroups'] += "SMES Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "SMESEmail"
            }
        ("San Marcos Middle School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "smms-teacher-templat"
                    $properties['addgroups'] += "SMMS Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "smms-teacher-templat"
                    $properties['addgroups'] += "SMMS Management Email"
                    
                } else {
                    $properties['templateuser'] = "smms-ss-template" 
                    $properties['addgroups'] += "SMMS Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "SMMSEmail"
            }
        ("San Marcos High School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "smhs-teach-template"
                    $properties['addgroups'] += "SMHS Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "smhs-teach-template"
                    $properties['addgroups'] += "SMHS Management Email"
                    
                } else {
                    $properties['templateuser'] = "smhs-ss-template" 
                    $properties['addgroups'] += "SMHS Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "SMHSEmail"
            }
        ("Twin Oaks Elementary School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "toes-teacher-templat"
                    $properties['addgroups'] += "TOE Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "toes-teacher-templat"
                    $properties['addgroups'] += "TOE Management Email"
                    
                } else {
                    $properties['templateuser'] = "toes-ss-template" 
                    $properties['addgroups'] += "TOE Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "TOESEmail"
            }
        ("Twin Oaks High School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "tohs-teacher-templat"
                    $properties['addgroups'] += "TOHS Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "tohs-teacher-templat"
                    $properties['addgroups'] += "TOHS Management Email"
                    
                } else {
                    $properties['templateuser'] = "tohs-ss-template" 
                    $properties['addgroups'] += "TOHS Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "TOHSEmail"
            }
        ("Woodland Park Middle School")
            {
                if ($Title.contains("Teacher"))
                {
                    $properties['templateuser'] = "wpms-teacher-templat"
                    $properties['addgroups'] += "WPMS Certificated Email"
                } elseif ($Title.contains('Principal')) {
                    $properties['templateuser'] = "wpms-teacher-templat"
                    $properties['addgroups'] += "WPMS Management Email"
                    
                } else {
                    $properties['templateuser'] = "wpms-ss-template" 
                    $properties['addgroups'] += "WPMS Classified Email"
                }
                $properties['emailto'] = get-emailaddresses "WPMSEmail"
            }
        ("DO Accounting")
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Accounting"
                $properties['ou'] = "OU=BS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "DOEmail"
            }
        ("DO Business Svs.")
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Business Svs."
                $properties['ou'] = "OU=BS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "DOEmail"
            }

        ("Pace Promise")
            {
                $templateuser = "do-ss-template"
                $department = "San Marcos Promise"
                $OU = "OU=San Marcos Promise,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                $EmailTo = $DOEmail
            }


        ("DO Child Nutrition Svs.")
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Child Nutrition Svs."
                $properties['ou'] = "OU=CNS District Office Staff,OU=Users,OU=CNS,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "DOEmail"
                $properties['addgroups'] += "CNS Classified Email"
            }
        ("Child Nutrition Services")
            {
                $properties['templateuser'] = "CNS-Template"
                $properties['department'] = "Child Nutrition Svs."
                $properties['ou'] = "OU=CNS Asst.,OU=Users,OU=CNS,OU=SMUSD,DC=smusd,DC=local"
                $properties['addgroups'] += "CNS Classified Email"
                $properties['emailto'] = get-emailaddresses "CNSEmail"
            }
        ("DO Curriculum")
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Curriculum"
                $properties['ou'] = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "DOEmail"

            }
        ("DO Human Resources")
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Human Resources"
                $properties['ou'] = "OU=HR&D,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "DOEmail"
            }
        ("DO Instructional Svs.")
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Instructional Svs."
                $properties['ou'] = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "DOEmail"
            }
        ("DO Kids on Campus")
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Kids on Campus"
                $properties['ou'] = "OU=BS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "KOCEmail"
                $properties['addgroups'] += "KOC Classified Email"
                $properties['addgroups'] += "DO Classified Email"
            }
        ('KOC/Carrillo Elementary')
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Kids on Campus"
                $properties['ou'] = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "KOCEmail"
                $properties['companyoverride'] = "Carrillo Elementary"
                $properties['addgroups'] += "KOC Classified Email"
                $properties['addgroups'] += "koc-car-email"
            }
        ('KOC/Discovery Elementary')
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Kids on Campus"
                $properties['ou'] = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "KOCEmail"
                $properties['companyoverride'] = "Discovery Elementary"
                $properties['addgroups'] += "KOC Classified Email"
                $properties['addgroups'] += "koc-dsc-email"
            }
        ('KOC/Double Peak School')
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Kids on Campus"
                $properties['ou'] = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "KOCEmail"
                $properties['companyoverride'] = "Double Peak"
                $properties['addgroups'] += "KOC Classified Email"
                $properties['addgroups'] += "koc-dps-email"
            }
        ('KOC/Knob Hill Elementary')
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Kids on Campus"
                $properties['ou'] = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "KOCEmail"
                $properties['companyoverride'] = "Knob Hill Elementary"
                $properties['addgroups'] += "KOC Classified Email"
                $properties['addgroups'] += "koc-kh-email"
            }
        ('KOC/La Costa Meadows Elementary')
            {
                
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Kids on Campus"
                $properties['ou'] = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "KOCEmail"
                $properties['companyoverride'] = "La Costa Meadows Elementary"
                $properties['addgroups'] += "KOC Classified Email"
                $properties['addgroups'] += "koc-lcm-email"
            }
        ('KOC/Paloma Elementary')
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Kids on Campus"
                $properties['ou'] = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "KOCEmail"
                $properties['companyoverride'] = "Paloma Elementary"
                $properties['addgroups'] += "KOC Classified Email"
                $properties['addgroups'] += "koc-pal-email"
            }
        ('KOC/Richland Elementary')
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Kids on Campus"
                $properties['ou'] = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "KOCEmail"
                $properties['companyoverride'] = "Richland Elementary"
                $properties['addgroups'] += "KOC Classified Email"
                $properties['addgroups'] += "koc-rl-email"
            }
        ('KOC/San Elijo Elementary')
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Kids on Campus"
                $properties['ou'] = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "KOCEmail"
                $properties['companyoverride'] = "San Elijo Elementary"
                $properties['addgroups'] += "KOC Classified Email"
                $properties['addgroups'] += "koc-sees-email"
            }
        ('KOC/Twin Oaks Elementary')
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Kids on Campus"
                $properties['ou'] = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "KOCEmail"
                $properties['companyoverride'] = "Twin Oaks Elementary"
                $properties['addgroups'] += "KOC Classified Email"
                $properties['addgroups'] += "koc-toes-email"
            }
        ("Kids on Campus")
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Kids on Campus"
                $properties['ou'] = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "KOCEmail"
                $properties['addgroups'] += "KOC Classified Email"
            }
        ("DO Pupil Personnel Svs.")
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Pupil Personnel Svs."
                $properties['ou'] = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "DOEmail"
            }
        ("DO Purchasing")
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Purchasing"
                $properties['ou'] = "OU=BS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "DOEmail"
            }
        ("DO Special Education")
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Special Education"
                $properties['ou'] = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "DOEmail"
            }
        ("DO Special Programs")
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Special Programs"
                $properties['ou'] = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "DOEmail"
            }
        ("DO Technology")
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Technology"
                $properties['ou'] = "OU=IT,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "DOEmail"
            }
        ("Facilities Dept.")
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['department'] = "Facilities Dept."
                $properties['ou'] = "OU=Facilities,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "DOEmail"
                $properties['addgroups'] += "Facilities Staff Email"
            }
        ("Language Assessment Center")
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['ou'] = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "DOEmail"
            }
        ("Maintenance and Operations")
            {
                $properties['templateuser'] = "mo-ss-template"
                $properties['emailto'] = get-emailaddresses "MOEmail"
                $properties['addgroups'] = "Maintenance Classified Email"
                if ($Title.contains("grounds"))
                    {
                        $properties['ou'] = "OU=Users,OU=Grounds,OU=M&O,OU=SMUSD,DC=smusd,DC=local"
                          
                    } else {
                        $properties['ou'] = "OU=Users,OU=Maint,OU=M&O,OU=SMUSD,DC=smusd,DC=local"
                    }

            }
        ("Transportation")
            {
                    
                $properties['emailto'] = get-emailaddresses "MOEmail"
                $properties['addgroups'] += "Transportation Classified Email"
                if ($Title.contains("driver"))
                    {
                        $properties['templateuser'] = "transdrivertemplate"
                        $properties['ou'] = "OU=Drivers,OU=Users,OU=TRANS,OU=M&O,OU=SMUSD,DC=smusd,DC=local"
                          
                    } elseif ($Title.contains('Mechanic')) {
                        $properties['templateuser'] = "transdrivertemplate"
                        $properties['ou'] = "OU=Mechanics,OU=Users,OU=TRANS,OU=M&O,OU=SMUSD,DC=smusd,DC=local"
                    } elseif ($Title.contains('aide')) {
                        $properties['templateuser'] = "transdrivertemplate"
                        $properties['ou'] = "OU=Support Staff,OU=Users,OU=TRANS,OU=M&O,OU=SMUSD,DC=smusd,DC=local"
                    } else {
                        $properties['templateuser'] = "transdrivertemplate"
                        $properties['ou'] = "OU=Admin,OU=Users,OU=TRANS,OU=M&O,OU=SMUSD,DC=smusd,DC=local"
                    }
            }
        ("Multisite")
            {
                $properties['templateuser'] = "do-ss-template"
                $properties['ou'] = "OU=Users,OU=MultiSite,OU=SMUSD,DC=smusd,DC=local"
                $properties['emailto'] = get-emailaddresses "DOEmail"
            }

    }
    if ($Company.contains("DO") -and (-not $Company.contains("Double"))) {
        if ($Title.contains("Director") -or $Title.contains("Principal") -or $Title.contains("Superintendent") -or $Title.contains("Supt.")) {
            $properties['addgroups'] += "DO Management Email"
        } elseif ($Title.contains("Teacher")) {
            $properties['addgroups'] += "DO Certificated Email"
        } else {
            $properties['addgroups'] += "DO Classified Email"
        }
                
    }
    new-object -Property $properties -TypeName psobject
}