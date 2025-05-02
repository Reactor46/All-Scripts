# #############################################################################
# CompanyD - SCRIPT - POWERSHELL
# NAME: DailyHealthScript.ps1
# 
# AUTHOR:  OffHand
# DATE:  2017.03.21
# EMAIL: OffHand
# 
# COMMENT:  This script will run across several critical systems to query 
#           their health status and send a report to technical SysAdmins.
#
# VERSION HISTORY
# 1.0 2017.03.21 Initial Version.
# 
#
# TO ADD
# -Add a Function to email
# -Fix the...
# #############################################################################



$reportpath = "\\fs1\departments\Information Systems\Infrastructure\Systems\"
$date = get-date -Format yyyyMMdd
$time = get-date -Format HHmm

# Exchange
#get-serverhealth -id dmgiovexch03 | ? {$_.alertvalue -ne "Healthy" -and $_.alertvalue -ne "Disabled" }

# Active Directory
dcdiag /e /v /c /s:IODC1 /f:"$reportpath\active directory\$date dcdiag.txt"
repadmin /queue * > "$reportpath\active directory\$date repadmin.txt"
repadmin /replsummary >> "$reportpath\active directory\$date repadmin.txt"

# Group Policy
#Get-GPOReport -All -Path "$reportpath\active directory\$date FullGPOReport.html" -ReportType Html  -Server IODC1 
