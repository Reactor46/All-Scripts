# #############################################################################
# DMG - SCRIPT - POWERSHELL
# NAME: CheckServerServicesandUptime.ps1
# 
# AUTHOR:  Jayson Cox, DMG
# DATE:  2017/05/11
# EMAIL: jayson_cox@dmgaz.org
# 
# COMMENT:  This script will compare services status with startup type and 
#           make sure everything is running. It will also check uptime to make
#           sure the server rebooted.
#
# VERSION HISTORY
# 1.0 2017.05.15 Initial Version.
# 1.1 2017.05.18 Added $results Splat, and foreach services for cleaner output
# 
#
# TO ADD
# -add date/time stamp of script run to log file
# -Add a Function to check system uptime
# -Incorporate date/time into $serviceresults file name
# -Function to compare against previous results
# -Function to run external script of custom validation per server
# -Restart machine (maybe)
# -Change name of script and turn into module
# #############################################################################


$computers = import-csv c:\temp\computers.csv

$serviceresults = "c:\temp\serviceresults.csv"
foreach ($computer in $computers) {
    
    $cname = $computer.computername
    $services = Get-WmiObject win32_service -ComputerName $cname | 
        where {$_.state -ne "Running" -and $_.startmode -eq "Auto"}
        
        # need a foreach loop on the $services so that each service gets its own line. Much easier to read.
        foreach ($service in $services) {

            $results = [ordered]@{'ComputerName'=$cname;
                         'ServiceName'=$service.Name;
                         'ServiceState'=$service.state;
                         'ServiceStartMode'=$service.startmode
                         #'ComputerUptime'=
                        }

            $obj = New-Object -TypeName psobject -Property $results

             
            Write-Output -InputObject $obj |
                Export-Csv $serviceresults -NoTypeInformation -append

        }
    #Add-Content -Value $results -Path $serviceresults 

    #$uptime = (Get-WmiObject Win32_OperatingSystem).LastBootUpTime 
    #$update = $uptime.Substring(0,8)
    #$up = $uptime.Substring(8,4)
    #Write-Verbose -Message -

}

