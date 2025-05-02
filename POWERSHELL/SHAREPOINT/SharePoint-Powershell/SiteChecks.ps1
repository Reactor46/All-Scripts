$ContentDBs = @(
"SP2013_UAT_KSC_Intranet_Depts",
"SP2013_UAT_KSC_Intranet_Depts_BusinessDevelopment",
"SP2013_UAT_KSC_Intranet_Depts_HealthcareFinance",
"SP2013_UAT_KSC_Intranet_Depts_HR",
"SP2013_UAT_KSC_Intranet_Depts_HumanResources",
"SP2013_UAT_KSC_Intranet_Depts_InfoSystems",
"SP2013_UAT_KSC_Intranet_Depts_KCA", 
"SP2013_UAT_KSC_Intranet_Depts_Marketing", 
"SP2013_UAT_KSC_Intranet_Depts_Operations",
"SP2013_UAT_KSC_Intranet_Depts_Place"
)
$Sites = @(
"https://uat-apps1.kscpulse.com/",
"https://uat-archive1.kscpulse.com/",
"https://uat-bi1.kscpulse.com/",
"https://uat-contenthub1.kscpulse.com/",
"https://uat-depts1.kscpulse.com/",
"https://uat-eflipchart1.kscpulse.com/",
"https://uat-my1.kscpulse.com/",
"https://uat-pulse1.kscpulse.com/",
"https://uat-search1.kscpulse.com/",
"https://uat-teams1.kscpulse.com/",
"https://uat-wiki1.kscpulse.com/"
)
ForEach($DB in $ContentDBs){
ForEach($Site in $Sites){
Test-SPContentDatabase -name $DB -webapplication $Site | Select Category, Error, UpgradeBlocking, Message, Remedy, Locations | Export-Csv -NoTypeInformation D:\Reports\SiteReports.csv -Append

 }
    }

