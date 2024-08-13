function Get-LoggedOnUser 
{
    [CmdletBinding()]
    Param
    (
        [Parameter(
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$True)]
        [String]$ComputerName = $env:COMPUTERNAME
    )

     begin
     {
        $ReportArray = @()

        $UserRegex = '.+Domain="(.+)",Name="(.+)"$' 
        $LogonRegex = '.+LogonId="(\d+)"$' 
 
        $LogonType = @{ "0"="Local System" 
                        "2"="Interactive" #(Local logon) 
                        "3"="Network" # (Remote logon) 
                        "4"="Batch" # (Scheduled task) 
                        "5"="Service" # (Service account logon) 
                        "7"="Unlock" #(Screen saver) 
                        "8"="NetworkCleartext" # (Cleartext network logon) 
                        "9"="NewCredentials" #(RunAs using alternate credentials) 
                        "10"="RemoteInteractive" #(RDP\TS\RemoteAssistance) 
                        "11"="CachedInteractive" #(Local w\cached credentials) 
                      } 
     }
     process
     {
        try
        {
            #Get WMI objects
            $WMISessions = Get-WmiObject win32_logonsession -ComputerName $ComputerName
            $WMILoggedUsers = Get-WmiObject win32_loggedonuser -ComputerName $ComputerName

            foreach($User in $WMILoggedUsers)
            {
                #RegEx to extract User Name and Domain Name
                $User.Antecedent -match $UserRegex | Out-Null

                $UserName = $Matches[1] + "\" + $Matches[2]

                #RegEx to extract LogonID
                $User.Dependent -match $LogonRegex | Out-Null
                $LogonID = $Matches[1]

                $UserSession = $WMISessions | Where-Object {$_.LogonID -eq $LogonID}

                #setting object properties

                $Properties = @{"Server Name" = $ComputerName
                                User = $UserName
                                Auth = $UserSession.AuthenticationPackage
                                Type = $LogonType[$UserSession.LogonType.tostring()]
                                "Start Time" = [management.managementdatetimeconverter]::todatetime($UserSession.StartTime)
                                "Logon ID" = $LogonID
                               }
            
                #Creating object
                $Obj = New-Object psobject -Property $Properties

                #Adding to array
                $ReportArray += $Obj
            }
        }
        catch
        {
            throw
        }
   
    } 
    end
    {
        #Returning array
        $ReportArray
    }
 
}

$ComputerName | Get-LoggedOnUser | Select "Server Name", User, Type, Auth, "Logon ID", "Start Time" | Where-Object  {$_.Type -ne "Network"}