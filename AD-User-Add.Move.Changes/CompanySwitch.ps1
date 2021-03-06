﻿#Note: Email Arrays used in this file include:arrays:   $EmailCC, $ADEmail $CESEmail $DISEmail $DPSEmail $FHSEmail $JAESEmail $KHEmail $LCMEmail $MHHSEmail $MOEmail $PALEmail $RLEmail $SEESEmail 
    ##      $SEMSEmail $SMESEmail $SMMSEmail $SMHSEmail $TOESEmail $TOHSEmail $WPMSEmail $DOEmail $KOCEmail $CNSEmail $TestEmailAddress 
    ## They must all be defined in the EmailVariables.ps1 File

switch($Company)
            {
            ("Adult Transition Program")
                {
                $templateuser = "ATP-TEMPLATE"
                
                    if ($Title.contains("Teacher"))
                    {
                        
                        $AddGroups += "ATP Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        
                        $AddGroups += "ATP Management Email"
                    } else {
                         
                        $AddGroups += "ATP Classified Email"
                    }
                    $EmailTo = $DOEmail
                }
            ("La Mirada Academy")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "lma-teach-template"
                        $AddGroups += "LMA Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "lma-teach-template"
                        $AddGroups += "LMA Management Email"
                    } else {
                        $templateuser = "lma-ss-template" 
                        $AddGroups += "LMA Classified Email"
                    }
                    $EmailTo = $LMAEmail
                }
            ("Carrillo Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "ces-teacher-template"
                        $AddGroups += "CAR Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "ces-teacher-template"
                        $AddGroups += "CAR Management Email"
                    }  else {
                        $templateuser = "ces-ss-template" 
                        $AddGroups += "CAR Classified Email"
                    }
                    $EmailTo = $CESEmail
                }
            ("Double Peak School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "dps-teacher-template"
                        $AddGroups += "DPS Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "dps-teacher-template"
                        $AddGroups += "DPS Management Email"
                    } else {
                        $templateuser = "dps-ss-template" 
                        $AddGroups += "DPS Classified Email"
                    }
                    $EmailTo = $DPSEmail

                }
            ("Discovery Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "dis-teacher-template"
                        $AddGroups += "DIS Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "dis-teacher-template"
                        $AddGroups += "DIS Management Email"
                    
                    } else {
                        $templateuser = "dis-ss-template" 
                        $AddGroups += "DIS Classified Email"
                    }
                    $EmailTo = $DISEmail
                }
            ("Foothills High School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "fhs-teacher-template"
                        $AddGroups += "FH Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "fhs-teacher-template"
                        $AddGroups += "FH Management Email"
                    
                    } else {
                        $templateuser = "fhs-ss-template" 
                        $AddGroups += "FH Classified Email"
                    }
                    $EmailTo = $FHSEmail
                }
            ("Joli Ann Leichtag Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "jaes-teacher-templat"
                        $AddGroups += "JALE Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "jaes-teacher-templat"
                        $AddGroups += "JALE Management Email"
                    
                    } else {
                        $templateuser = "jaes-ss-template" 
                        $AddGroups += "JALE Classified Email"
                    }
                    $EmailTo = $JAESEmail
                }
            ("Knob Hill Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "kh-teacher-template"
                        $AddGroups += "KH Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "kh-teacher-template"
                        $AddGroups += "KH Management Email"
                    
                    } else {
                        $templateuser = "kh-ss-template" 
                        $AddGroups += "KH Classified Email"
                    }
                    $EmailTo = $KHEmail
                }
            ("La Costa Meadows Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "lcm-teacher-template"
                        $AddGroups += "LCM Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "lcm-teacher-template"
                        $AddGroups += "LCM Management Email"
                    
                    } else {
                        $templateuser = "lcm-ss-template" 
                        $AddGroups += "LCM Classified Email"
                    }
                    $EmailTo = $LCMEmail
                }
            ("Mission Hills High School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "mhhs-teacher-templat"
                        $AddGroups += "MHHS Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "mhhs-teacher-templat"
                        $AddGroups += "MHHS Management Email"
                    
                    } else {
                        $templateuser = "mhhs-ss-template" 
                        $AddGroups += "MHHS Classified Email"
                    }
                    $EmailTo = $MHHSEmail
                }
            ("Paloma Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "pal-teacher-template"
                        $AddGroups += "PAL Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "pal-teacher-template"
                        $AddGroups += "PAL Management Email"
                    
                    } else {
                        $templateuser = "pal-ss-template" 
                        $AddGroups += "PAL Classified Email"
                    }
                    $EmailTo = $PALEmail
                }
            ("Richland Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "rl-teacher-template"
                        $AddGroups += "RL Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "rl-teacher-template"
                        $AddGroups += "RL Management Email"
                    
                    } else {
                        $templateuser = "rl-ss-template" 
                        $AddGroups += "RL Classified Email"
                    }
                    $EmailTo = $RLEmail
                }
            ("San Elijo Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "sees-teacher-templat"
                        $AddGroups += "SEES Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "sees-teacher-templat"
                        $AddGroups += "SEES Management Email"
                    
                    } else {
                        $templateuser = "sees-ss-template" 
                        $AddGroups += "SEES Classified Email"
                    }
                    $EmailTo = $SEESEmail
                }
            ("San Elijo Middle School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "sems-teacher-templat"
                        $AddGroups += "SEMS Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "sems-teacher-templat"
                        $AddGroups += "SEMS Management Email"
                    
                    } else {
                        $templateuser = "sems-ss-template" 
                        $AddGroups += "SEMS Classified Email"
                    }
                    $EmailTo = $SEMSEmail
                }
            ("San Marcos Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "smes-teacher-templat"
                        $AddGroups += "SMES Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "smes-teacher-templat"
                        $AddGroups += "SMES Management Email"
                    
                    } else {
                        $templateuser = "smes-ss-template" 
                        $AddGroups += "SMES Classified Email"
                    }
                    $EmailTo = $SMESEmail
                }
            ("San Marcos Middle School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "smms-teacher-templat"
                        $AddGroups += "SMMS Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "smms-teacher-templat"
                        $AddGroups += "SMMS Management Email"
                    
                    } else {
                        $templateuser = "smms-ss-template" 
                        $AddGroups += "SMMS Classified Email"
                    }
                    $EmailTo = $SMMSEmail
                }
            ("San Marcos High School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "smhs-teach-template"
                        $AddGroups += "SMHS Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "smhs-teach-template"
                        $AddGroups += "SMHS Management Email"
                    
                    } else {
                        $templateuser = "smhs-ss-template" 
                        $AddGroups += "SMHS Classified Email"
                    }
                    $EmailTo = $SMHSEmail
                }
            ("Twin Oaks Elementary School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "toes-teacher-templat"
                        $AddGroups += "TOE Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "toes-teacher-templat"
                        $AddGroups += "TOE Management Email"
                    
                    } else {
                        $templateuser = "toes-ss-template" 
                        $AddGroups += "TOE Classified Email"
                    }
                    $EmailTo = $TOESEmail
                }
            ("Twin Oaks High School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "tohs-teacher-templat"
                        $AddGroups += "TOHS Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "tohs-teacher-templat"
                        $AddGroups += "TOHS Management Email"
                    
                    } else {
                        $templateuser = "tohs-ss-template" 
                        $AddGroups += "TOHS Classified Email"
                    }
                    $EmailTo = $TOHSEmail
                }
            ("Woodland Park Middle School")
                {
                    if ($Title.contains("Teacher"))
                    {
                        $templateuser = "wpms-teacher-templat"
                        $AddGroups += "WPMS Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $templateuser = "wpms-teacher-templat"
                        $AddGroups += "WPMS Management Email"
                    
                    } else {
                        $templateuser = "wpms-ss-template" 
                        $AddGroups += "WPMS Classified Email"
                    }
                    $EmailTo = $WPMSEmail
                }
            ("DO Accounting")
                {
                    $templateuser = "do-ss-template"
                    $department = "Accounting"
                    $OU = "OU=Business Services,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Business Svs.")
                {
                    $templateuser = "do-ss-template"
                    $department = "Business Svs."
                    $OU = "OU=Business Services,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("San Marcos Promise")
                {
                    $templateuser = "do-ss-template"
                    $department = "San Marcos Promise"
                    $OU = "OU=San Marcos Promise,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Child Nutrition Svs.")
                {
                    $templateuser = "do-ss-template"
                    $department = "Child Nutrition Svs."
                    $OU = "OU=CNS District Office Staff,OU=Users,OU=CNS,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                    $AddGroups += "CNS Classified Email"
                }
            ("Child Nutrition Services")
                {
                    $templateuser = "CNS-Template"
                    $department = "Child Nutrition Svs."
                    $OU = "OU=CNS Asst.,OU=Users,OU=CNS,OU=SMUSD,DC=smusd,DC=local"
                    $AddGroups += "CNS Classified Email"
                    $EmailTo = $CNSEmail
                }
            ("DO Curriculum")
                {
                    $templateuser = "do-ss-template"
                    $department = "Curriculum"
                    $OU = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail

                }
            ("DO Human Resources")
                {
                    $templateuser = "do-ss-template"
                    $department = "Human Resources"
                    $OU = "OU=HR&D,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Instructional Svs.")
                {
                    $templateuser = "do-ss-template"
                    $department = "Instructional Svs."
                    $OU = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Kids on Campus")
                {
                    $templateuser = "do-ss-template"
                    $department = "Kids on Campus"
                    $OU = "OU=Business Services,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $KOCDOEmail
                    $AddGroups += "KOC Classified Email"
                    $AddGroups += "DO Classified Email"
                }
            ('KOC/Carrillo Elementary')
                {
                    $templateuser = "do-ss-template"
                    $department = "Kids on Campus"
                    $OU = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $KOCEmail
                    $companyoverride = "Carrillo Elementary"
                    $AddGroups += "KOC Classified Email"
                    $AddGroups += "koc-car-email"
                }
            ('KOC/Discovery Elementary')
                {
                    $templateuser = "do-ss-template"
                    $department = "Kids on Campus"
                    $OU = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $KOCEmail
                    $companyoverride = "Discovery Elementary"
                    $AddGroups += "KOC Classified Email"
                    $AddGroups += "koc-dsc-email"
                }
            ('KOC/Double Peak School')
                {
                    $templateuser = "do-ss-template"
                    $department = "Kids on Campus"
                    $OU = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $KOCEmail
                    $companyoverride = "Double Peak"
                    $AddGroups += "KOC Classified Email"
                    $AddGroups += "koc-dps-email"
                }
            ('KOC/Knob Hill Elementary')
                {
                    $templateuser = "do-ss-template"
                    $department = "Kids on Campus"
                    $OU = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $KOCEmail
                    $companyoverride = "Knob Hill Elementary"
                    $AddGroups += "KOC Classified Email"
                    $AddGroups += "koc-kh-email"
                }
            ('KOC/La Costa Meadows Elementary')
                {
                    $templateuser = "do-ss-template"
                    $department = "Kids on Campus"
                    $OU = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $KOCEmail
                    $companyoverride = "La Costa Meadows Elementary"
                    $AddGroups += "KOC Classified Email"
                    $AddGroups += "koc-lcm-email"
                }
            ('KOC/Paloma Elementary')
                {
                    $templateuser = "do-ss-template"
                    $department = "Kids on Campus"
                    $OU = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $KOCEmail
                    $companyoverride = "Paloma Elementary"
                    $AddGroups += "KOC Classified Email"
                    $AddGroups += "koc-pal-email"
                }
            ('KOC/Richland Elementary')
                {
                    $templateuser = "do-ss-template"
                    $department = "Kids on Campus"
                    $OU = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $KOCEmail
                    $companyoverride = "Richland Elementary"
                    $AddGroups += "KOC Classified Email"
                    $AddGroups += "koc-rl-email"
                }
            ('KOC/San Elijo Elementary')
                {
                    $templateuser = "do-ss-template"
                    $department = "Kids on Campus"
                    $OU = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $KOCEmail
                    $companyoverride = "San Elijo Elementary"
                    $AddGroups += "KOC Classified Email"
                    $AddGroups += "koc-sees-email"
                }
            ('KOC/Twin Oaks Elementary')
                {
                    $templateuser = "do-ss-template"
                    $department = "Kids on Campus"
                    $OU = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $KOCEmail
                    $companyoverride = "Twin Oaks Elementary"
                    $AddGroups += "KOC Classified Email"
                    $AddGroups += "koc-toes-email"
                }
            ("Kids on Campus")
                {
                    $templateuser = "do-ss-template"
                    $department = "Kids on Campus"
                    $OU = "OU=Users,OU=KOC,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $KOCEmail
                    $AddGroups += "KOC Classified Email"
                }
            ("DO Purchasing")
                {
                    $templateuser = "do-ss-template"
                    $department = "Purchasing"
                    $OU = "OU=Business Services,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Special Education")
                {
                    $templateuser = "do-ss-template"
                    $department = "Special Education"
                    $OU = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Special Programs")
                {
                    $templateuser = "do-ss-template"
                    $department = "Special Programs"
                    $OU = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Stu Ser")
                {
                    $templateuser = "do-ss-template"
                    $department = "Student Services"
                    $OU = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Subs")
                {
                    $templateuser = "do-ss-template"
                    $OU = "OU=Do Subs,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Superintendent")
                {
                    $templateuser = "do-ss-template"
                    $department = "Superintendent"
                    $OU = "OU=Superintendent,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("DO Technology")
                {
                    $templateuser = "do-ss-template"
                    $department = "Technology"
                    $OU = "OU=IT,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("Facilities Dept.")
                {
                    $templateuser = "do-ss-template"
                    $department = "Facilities Dept."
                    $OU = "OU=Facilities,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                    $AddGroups += "Facilities Staff Email"
                }
            ("Language Assessment Center")
                {
                    $templateuser = "do-ss-template"
                    $OU = "OU=IS,OU=Users,OU=DO,OU=SMUSD,DC=smusd,DC=local"
                    $EmailTo = $DOEmail
                }
            ("Maintenance and Operations")
                {
                    $templateuser = "mo-ss-template"
                    $EmailTo = $MOEmail
                    $AddGroups = "Maintenance Classified Email"
                    if ($Title.contains("grounds"))
                        {
                            $OU = "OU=Users,OU=Grounds,OU=M&O,OU=SMUSD,DC=smusd,DC=local"
                          
                        } else {
                            $OU = "OU=Users,OU=Maint,OU=M&O,OU=SMUSD,DC=smusd,DC=local"
                        }

                }
            ("Transportation")
                {
                    
                    $EmailTo = $MOEmail
                    $AddGroups += "Transportation Classified Email"
                    if ($Title.contains("driver"))
                        {
                            $templateuser = "transdrivertemplate"
                            $OU = "OU=Drivers,OU=Users,OU=TRANS,OU=M&O,OU=SMUSD,DC=smusd,DC=local"
                          
                        } elseif ($Title.contains('Mechanic')) {
                            $templateuser = "transdrivertemplate"
                            $OU = "OU=Mechanics,OU=Users,OU=TRANS,OU=M&O,OU=SMUSD,DC=smusd,DC=local"
                        } elseif ($Title.contains('aide')) {
                            $templateuser = "transdrivertemplate"
                            $OU = "OU=Support Staff,OU=Users,OU=TRANS,OU=M&O,OU=SMUSD,DC=smusd,DC=local"
                        } else {
                            $templateuser = "transdrivertemplate"
                            $OU = "OU=Admin,OU=Users,OU=TRANS,OU=M&O,OU=SMUSD,DC=smusd,DC=local"
                        }
                }
            ("Multisite")
                {
                    $templateuser = "multisite-ss-templat"
                    if ($Title.contains("sub"))
                    {
                        $OU = "OU=subs,OU=Users,OU=MultiSite,OU=SMUSD,DC=smusd,DC=local"
                    } elseif ($title.contains("Social Work Intern")) {
                        $OU = "OU=Social Work Interns,OU=Users,OU=MultiSite,OU=SMUSD,DC=smusd,DC=local"
                    } else {
                        $OU = "OU=Users,OU=MultiSite,OU=SMUSD,DC=smusd,DC=local"
                    }
                    $EmailTo = $MultiSiteEmail
                    if ($Title.contains("Teacher")-or $Title.contains("psychologist"))
                    {
                        $AddGroups += "Multisite Certificated Email"
                    } elseif ($Title.contains('Principal')) {
                        $AddGroups += "Multisite Management Email"
                    
                    } else {
                        $AddGroups += "Multisite Classified Email"
                    }
                }

            }
            if ($Company.contains("DO") -and (-not $Company.contains("Double"))) {
                if ($Title.contains("Director") -or $Title.contains("Principal") -or $Title.contains("Superintendent") -or $Title.contains("Supt.")) {
                    $AddGroups += "DO Management Email"
                } elseif ($Title.contains("Teacher")) {
                    $AddGroups += "DO Certificated Email"
                } else {
                    $AddGroups += "DO Classified Email"
                }
                
            }
            if ($title.contains("Intern")) {
                $AddGroups += "SMUSD-Interns"
            }
            if ($title.contains("Student Teacher")) {
                $AddGroups += "SMUSD-StudentTeachers"
            }