<#  
.Author
    Jim Adkins
    Winsys
    Windows System Administrator II    
    
    EZAdmin - This is a simple tool for CreditOneBank Sys Admins to streemline a number of common tasks.

    Useage : 
        To use this you will need to import it:
        Import-Modual C:\Scripts\EzAdmin.psm1 -Force -Verbose

        Type Ezhelp for command usage and list.


#>

function Get-OS {
  Param([string]$computername=$(Throw "You must specify a computername."))
  $wmi=Get-WmiObject Win32_OperatingSystem -computername $computername | Select Caption, Version, BuildNumber
  
  write $wmi

}

Function Get_Logged_In_User{
#Needs Error handling code - will return an error if no user is logged in.
    [CmdletBinding()]
        
        Param(
             [Parameter(Mandatory=$True, Position=0)]       
             [String] $Target_Host
        )

            $FullUser= Get-WmiObject –ComputerName $Target_Host –Class Win32_ComputerSystem | Select-Object UserName     
            $UserDomain , $LoggedIN = $FullUser.username.split("\")
            $Logged_In_User = Get-ADUser -Identity $LoggedIN | Select Name
            Write-Host "Full UserName $FullUser"
            Write-Host $Logged_In_User.Name -ForegroundColor Green
            }

function Get-UDVariable {
  get-variable | where-object {(@(
    "FormatEnumerationLimit",
    "MaximumAliasCount",
    "MaximumDriveCount",
    "MaximumErrorCount",
    "MaximumFunctionCount",
    "MaximumVariableCount",
    "PGHome",
    "PGSE",
    "PGUICulture",
    "PGVersionTable",
    "PROFILE",
    "PSSessionOption",
    "moduleBrowser",
    "psISE",
    "psUnsupportedConsoleApplications"
    

    ) -notcontains $_.name) -and `
    (([psobject].Assembly.GetType('System.Management.Automation.SpecialVariables').GetFields('NonPublic,Static') | Where-Object FieldType -eq ([string]) | ForEach-Object GetValue $null)) -notcontains $_.name
    }
}

Function Get_DNS_Servers{

    [CmdletBinding()]
    
    Param(
        [Parameter(Mandatory=$False, Position=0)]
        [String] $HostName = $Env:ComputerName
    )
    
    Get-DnsClientServerAddress -AddressFamily IPv4 -CimSession $HostName | Select PSComputerName, ServerAddresses,InterfaceAlias | Where{$_.InterfaceAlias -notlike "*Loopback*"}

     

}

Function Get_Last_Reboot{

    [CmdletBinding()]
    
    Param(
        [Parameter(Mandatory=$false)]
        [String] $HostName = $env:COMPUTERNAME
    
    )

        $OS=Get-WmiObject Win32_OperatingSystem -ComputerName $HostName | Select  CSName, LastBootupTime, LocalDateTime
        $BootTime=[System.Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootupTime)
        $Uptime=[System.Management.ManagementDateTimeConverter]::ToDateTime($os.LocalDateTime) - $BootTime

        #Write-Host "`n" $Uptime.Days "Day(s) Since last Reboot" -ForegroundColor Green
        Write-Host ("'n {0} Day(s) since last Reboot" -f $Uptime.Days) -ForegroundColor Green
}

Function Do_IIS_Restart{

    [CmdletBinding()]
   
    Param(
        [Parameter(Mandatory=$True, Position=0)]
        [String]$Webserver 
    )
        
        Write-Debug "Do_IIS_Restart - Webserver = $webserver"
        Invoke-Command -ComputerName $Webserver {cd C:\Windows\System32\; ./cmd.exe /c "iisreset /restart" }
    
        $Status=Get-Service -ComputerName $Webserver -Name W3SVC | select Status
    
        #Console output
        Write-Host "W3SVC Status is" $status.Status -ForegroundColor Green

}

Function Start_Contoso{
  #This function will restart DataLayer service on a specified Server.
  #Added code to Also start the NCach Service.
    Param(
        [Parameter(Mandatory=$True, Position=0)]
        [String] $Target_Host
        
    )
    Write-Host "Starting ContosoDataLayerService on $Target_Host" -ForegroundColor Green

    Get-Service -ComputerName $Target_Host -Name 'ContosoDataLayerService' | Start-Service  #Starts Contoso Service
    Get-Service -ComputerName $Target_Host -Name 'W3SVC'  | Start-Service      #Starts Web Service
    Get-Service -ComputerName $Target_Host -Name 'NCacheSvc' | Start-Service #Starts NCach Service
    Get-Service -ComputerName $Target_Host -Name 'ContosoDataLayerService'
    Get-Service -ComputerName $Target_Host -Name 'W3SVC'
    Get-Service -ComputerName $Target_Host -Name 'NCacheSvc' 
    


}

Function cPing{


    Param(
        [Parameter(Mandatory=$true, Position=0)]
        [String] $Target_Host

    )


        Ping -t $Target_Host


}

Function MT_Recycle{
    Param(
      [Parameter(Mandatory=$true, Position=0)]  
      [String]$Host_List
      )
      
      #This will take a server, Reboot it,  Wait for the reboot and then it will start the Contoso Service 
  ForEach($Target_Host in $Host_List)    
  {

        Get-VM -name $Target_Host | Restart-VMGuest -Confirm:$false
        Write-Host "Rebootin $Target_Host"
        While($ToolStatus -ne 'toolsNotRunning'){$ToolStatus = (Get-VM $Target_Host | Get-View).guest.toolsStatus}


          $ToolStatus = (Get-VM $Target_Host | Get-View).guest.toolsStatus
        Write-Host "Booting $vm"
        While($ToolStatus -eq 'toolsNotRunning'){ 
          $ToolStatus = (Get-VM $Target_Host | Get-View).guest.toolsStatus
            Write-Host '.' -NoNewline
        } Write-Host "`nBoot Up Complete" 

        
        
        Start_Contoso -Target_Host $Target_Host

    }}

Function EzAdmin_Help{

  Clear-Host
  Write-Host "Get-LoggedInUser - Outputs the Username of the logged in user."
  Write-Host 'Usage: Get-LoggedInUser -Target_Host <HOSTNAME>' -ForegroundColor Yellow
  Write-Host "Get-DNS-Servers"
  Write-Host "Get-Reboot - Gets the Last time the server was rebooted."
  Write-Host "IIS-Restart -  Restarts IIS on the Specified Server"
  Write-Host "StartContoso - Start Contoso Data Layer Service on Specified Host. Usage: StartContoso lasmt01"
  Write-Host "MT-Recycle - Reboots Server and Re-Starts the DataLayerService, NCach and W3SVC services after the server reboots."
  Write-Host "Get-OS - Returns the Windows Operating system on the specified server"
  Write-Host 'Get-Vars - List the local variable values' 
  Write-Host 'Get-AccountStatus - Show Lockout Status/ Is it expierd and last pasword set date'
  Write-Host "EzHelp - This listing."
  Write-Host "Version: EzAdmin v1.0"
  
}

      function Get-AccountStatus{

       Param (
        [String] $UserName
        )

        $ADUser = Get-ADUser -Identity $UserName -Properties *
       
        Write-Host ("Password Last set on {0}." -f ([datetime]::FromFileTime($ADUser.pwdLastSet))) -ForegroundColor Yellow
        Write-Host ('Is account Expierd? {0}.' -f $ADUser.PasswordExpired) -ForegroundColor Yellow
        Write-Host ('Current Locout Status {0}.' -f $aduser.LockedOut) -ForegroundColor Yellow

        }

Set-Alias -Name Get-LoggedInUser -Value Get_Logged_In_User
Set-Alias -Name Get-DNSServers -Value Get_DNS_Servers
Set-Alias -Name Get-Reboot -Value Get_Last_Reboot
Set-Alias -Name IIS-Restart -Value Do_IIS_Restart
Set-Alias -Name StartContoso -Value Start_Contoso
Set-Alias -Name Get-Vars -Value Get-UDVariable
Set-Alias -Name EzHelp -Value EzAdmin_Help
Set-Alias -Name MT-Recyle -Value MT_Recycle