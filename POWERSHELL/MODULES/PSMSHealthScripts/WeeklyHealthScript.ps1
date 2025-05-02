# #############################################################################
# CompanyD - SCRIPT - POWERSHELL
# NAME: WeeklyHealthScript.ps1
# 
# AUTHOR:  OffHand
# DATE:  2017/05/31
# EMAIL: Offhand
# 
# COMMENT:  This script will perform health checks and tasks on a weekly basis
#           to maintain CompanyD's environment
#
# VERSION HISTORY
# 1.0 20170531 Initial Version.
# 1.1 
#
# TO ADD
# -Add a Function to ...
# -Fix the...
# #############################################################################

$systemspath = "\\fs1\departments\Information Systems\Infrastructure\Systems\"
$date = get-date -Format yyyyMMdd

#Backup Group Policies
$gpodescription = "Weekly backup of Group Policies for comparision or small restores"
$gpopath = $systemspath + "Active Directory\GroupPolicy\Group Policy Backup $date\"
New-Item $gpopath -ItemType Directory
Backup-GPO -All -Comment $gpodescription -Path $gpopath -Server iodc1 

