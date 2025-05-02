<# 
.SYNOPSIS 
Using this script you can restart/shutdown and check server/Services health status and also capture IP Details for Bulk Servers. 
 
.DESCRIPTION 
This is script is added with below options
1.Returns IP Details of the Servers along with DNSHostName,Description,DHCPEnabled,IPAddress,IpSubnet,DefaultIPGateway,MACAddress,DNSServerSearchOrder. 
2. Check server reachablities and displays the status of the server.
3. Checks all the Automatic Services status and saves the stopped services in the CSV Format output file
4. Shutdown/Reboot the bulk Servers at once.. and you can check server status
* Note: PLease do test bebfore you run the script in the production

 
.PARAMETER ComputerName 
You can add FQDN,Hostname as input in the text file and save in the location. Change the Inoutfile location in the script based on your Inout file location

 
 
.INPUTS 
Input file should be modified  as per the your input location and saved the script before you execute the script. 
 
.EXAMPLE 
PS C:\> .\Servers_ServiceandRestart.ps1 
 
Prompts for the further actions and follows your answers and does all the functions 
installed applications on SVR001, SVR002, SVR003 and displays any errors encountered. 
 

#> 
 
#*============================================================================= 
#* Name:    Windowstechpro.Com
#* Created: 16/04/2017 
#* Author:     Radhakrishnan Govindan
#* Email:     Radhakrishnan.G@windowstechpro.com 
#* 
#* Returns:    All installed software for the computer(s) 
#*----------------------------------------------------------------------------- 
#* Purpose:    Quickly returns information about all installed software on 
#* computers regardless of whether it was installed by Windows Installer. 
#* 
#*============================================================================= 
 
#*============================================================================= 
#* REVISION HISTORY 
#*----------------------------------------------------------------------------- 
#* Version:        1.0 
#* Date:         16/04/2017 
#* Time:         4:28 PM 
#* Issue:         IP Cature sends System.Strin[] values for the IP Details for CSV File
#* Solution:    Fixed System.Strin[] errors in the outfile for IP Cature using Expression Cmdlets
#* 
#*============================================================================= 
 
#*============================================================================= 
#* SCRIPT BODY 
#*=================================
# Hostnames TXT Location - Edit this line to fit your install path
$hostnamestxt = "C:\Scripts\hosts.txt"
$servers = get-content "$hostnamestxt"
# Main Menu Function
Function Main_Menu {
         Write-Host " Welcome to the Server Shutdown/Reboot Script!                                " -foregroundcolor white -backgroundcolor blue
         Write-Host "                                                                              " -foregroundcolor gray -backgroundcolor blue
         Write-Host " This script will allow you to mass shutdown/reboot and check the status of   " -foregroundcolor gray -backgroundcolor blue
         Write-Host " all servers found in the hosts.txt                                           " -foregroundcolor gray -backgroundcolor blue
         Write-Host "                                                                              " -foregroundcolor gray -backgroundcolor blue
         Write-Host " "
         Write-Host " Current location of the hosts.txt file: $hostnamestxt" -foregroundcolor black -backgroundcolor green
         $choices = [Management.Automation.Host.ChoiceDescription[]] ( `
      (new-Object Management.Automation.Host.ChoiceDescription "&1 List of Servers","List of Servers"),
         (new-Object Management.Automation.Host.ChoiceDescription "&2 Check Server Status","Check Server Status"),
         (new-Object Management.Automation.Host.ChoiceDescription "&3 Shutdown/Reboot Servers","Shutdown/Reboot Servers."),
         (new-Object Management.Automation.Host.ChoiceDescription "&4 Service Check","Service Check"),
      (new-Object Management.Automation.Host.ChoiceDescription "&4 IPCapture","IPCapture"),
      (new-Object Management.Automation.Host.ChoiceDescription "&Exit.","Exit"));
         $answer = $host.ui.PromptForChoice("Server Shutdown-Reboot Tool", "Which action would you like to perform?", $choices, 4)
if($answer -eq 0){
  Write-Host "Servers in host.txt:" -foregroundcolor white -backgroundcolor blue
  Write-Host "---------------------------"
  $servers
  Write-Host "---------------------------"
  Pause
  #cls
  Main_Menu
} 

elseif ($answer -eq 1) {
    Server_Status_Check
  Pause
  #cls
  Main_Menu
} 
elseif ($answer -eq 2){
  Server_Reboot_Shutdown_Menu
  #cls
  Main_Menu
} 
elseif ($answer -eq 3){

  Server_Service_Status_Check
  #cls
} 
elseif ($answer -eq 4){
IPDetails_Hunt
#cls
} 
elseif ($answer -eq 5){

  Write-Host " "
  Write-Host "Exiting Application...Goodbye!" -foregroundcolor white -backgroundcolor blue
  Write-Host " "
} 
}

# Pause Function
Function Pause {
  Write-Host "Press Any Key To Continue..." -foregroundcolor gray -backgroundcolor blue
# $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") >$null
  }

# Server_Reboot_Shutdown_Menu - Prompts for Shutdown or Reboot
Function Server_Reboot_Shutdown_Menu {
       $choices1 = [Management.Automation.Host.ChoiceDescription[]] ( `
    (new-Object Management.Automation.Host.ChoiceDescription "&Reboot Servers","Reboot Servers"),
    (new-Object Management.Automation.Host.ChoiceDescription "&Shutdown Servers","Shutdown Servers"),
       (new-Object Management.Automation.Host.ChoiceDescription "&Exit to Main Menu","&Exit to Main Menu"));
       $answer1 = $host.ui.PromptForChoice("Server Reboot/Shutdown", "Which action would you like to perform?", $choices1, 2)
       if ($answer1 -eq 0) {
              Server_Reboot
              } elseif ($answer1 -eq 1) {
        Server_Shutdown
              } elseif ($answer1 -eq 2) {
        
              }
}

# Server_Shutdown Function - Pings and then Shuts Down All Servers in hosts.txt
Function Server_Shutdown {
    Write-Host "                                                                "
    Write-Host "                                                                " -foregroundcolor black -backgroundcolor red
    Write-Host " CAUTION! All host names in the host.txt file will be SHUTDOWN! " -foregroundcolor black -backgroundcolor red
    Write-Host "                                                                " -foregroundcolor black -backgroundcolor red
       $choices2 = [Management.Automation.Host.ChoiceDescription[]] ( `
    (new-Object Management.Automation.Host.ChoiceDescription "&Yes","Yes"),
       (new-Object Management.Automation.Host.ChoiceDescription "&No","No"));
       $answer2 = $host.ui.PromptForChoice("Server Shutdown", "Are you sure you want to continue?", $choices2, 1)
       if ($answer2 -eq 0) {
              write-host "Server Shutdown begin..." -foregroundcolor white -backgroundcolor blue
              Write-Host "---------------------------"
              $shutdown_reason = Read-Host "Server shutdown reason/comment"
        foreach($server in $servers){
                     ping -n 2 $server >$null
                     if($lastexitcode -eq 0){
                           write-host "Shutting Down $server..." -foregroundcolor black -backgroundcolor green
                           shutdown /s /f /m \\$server /d p:1:1 /t 01 /c "$shutdown_reason"
                     } else {
                           write-host "$server is OFFLINE/UNREACHABLE" -foregroundcolor black -backgroundcolor red
                           }
                     }
                     Write-Host "---------------------------"
                     Pause
              } elseif ($answer2 -eq 1) {

                     } 
}

# Server_Reboot Function - Pings and then Shuts Down All Servers in hosts.txt
Function Server_Reboot {
    Write-Host "                                                                "
    Write-Host "                                                                " -foregroundcolor black -backgroundcolor red
    Write-Host " CAUTION! All host names in the host.txt file will be REBOOTED! " -foregroundcolor black -backgroundcolor red
    Write-Host "                                                                " -foregroundcolor black -backgroundcolor red
       $choices3 = [Management.Automation.Host.ChoiceDescription[]] ( `
    (new-Object Management.Automation.Host.ChoiceDescription "&Yes","Yes"),
       (new-Object Management.Automation.Host.ChoiceDescription "&No","No"));
       $answer3 = $host.ui.PromptForChoice("Server Reboot", "Are you sure you want to continue?", $choices3, 1)
       if ($answer3 -eq 0) {
              write-host "Server Reboot begin..." -foregroundcolor white -backgroundcolor blue
              Write-Host "---------------------------"
        $reboot_reason = Read-Host "Server reboot reason/comment"
              foreach($server in $servers){
                     ping -n 2 $server >$null
                     if($lastexitcode -eq 0){
                           write-host "Rebooting $server..." -foregroundcolor black -backgroundcolor green
                           shutdown /r /f /m \\$server /d p:1:1 /t 01 /c "$reboot_reason"
                     } else {
                           write-host "$server is OFFLINE/UNREACHABLE" -foregroundcolor black -backgroundcolor red
                           }
                     }
                     Write-Host "---------------------------"
                     Pause
              } elseif ($answer3 -eq 1) {

                     } 
}

# Server_Status_Check Function - Pings Servers from hosts.txt and then shows Online/Offline
Function Server_Status_Check {
write-host "Checking Status of Servers..." -foregroundcolor white -backgroundcolor blue
Write-Host "---------------------------"
foreach($server in $servers){
              ping -n 3 $server >$null
              if($lastexitcode -eq 0) {
                     write-host "$server is ONLINE" -foregroundcolor black -backgroundcolor green
              } else {
                     write-host "$server is OFFLINE/UNREACHABLE" -foregroundcolor black -backgroundcolor red
              }
       }
       Write-Host "---------------------------"
}

# Server_Service_Status_Check - Check the Automatic Services and saves the stopped services in CSV File
Function Server_Service_Status_Check
{
write-host "Checking Services Status of Servers..." -foregroundcolor white -backgroundcolor blue
Write-Host "---------------------------"
foreach($server in $servers)
{
if (Test-Connection -ComputerName $server -Count 1 -TimeToLive 10 -Quiet) 
{
Get-WmiObject Win32_Service -ComputerName $server| Where-Object {$_.StartMode -eq 'Auto'  -and  $_.State -ne 'running'}`
 | select PSComputername,Name,Startmode,State|Export-Csv C:\scripts\Servicestatus.csv -NoTypeInformation -Append
}
else
{ # This server is not online
      $Computer= $server
              Write-Warning -Message "Unable to connect - $Computer"
} #end else
} 

}

Function IPDetails_Hunt
{
write-host "Checking IP Details of Servers..." -foregroundcolor white -backgroundcolor blue
Write-Host "---------------------------"
foreach($server in $servers)
{
ping -n 2 $server >$null
                     if($lastexitcode -eq 0){
  $newINFO=Get-WmiObject -computername $server Win32_NetworkAdapterConfiguration|`
   Where-Object { $_.IPAddress -ne $null }`
  |Select-Object DNSHostName,Description,DHCPEnabled,@{Name='IPAddress';Expression={$_.IPAddress}},`
  @{Name='Subnet';Expression={$_.IpSubnet -join '; '}},`
  @{Name='DefaultIPgateway';Expression={$_.DefaultIPgateway -join '; '}},`
   @{Name='DNSServerSearchOrder';Expression={$_.DNSServerSearchOrder -join '; '}},`
   MACAddress,WinsPrimaryServer, WINSSecindaryServer|`
    Export-Csv -Path C:\scripts\IPCapture.csv -NoTypeInformation -Append #location to save the output file
                     
} else {
write-host "$server is OFFLINE/UNREACHABLE" -foregroundcolor black -backgroundcolor red
}
} 

}
cls
Main_Menu
#Script ends here.. Thank you for using this script.. 
