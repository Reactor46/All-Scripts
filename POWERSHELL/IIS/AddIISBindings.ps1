# Microsoft
# Peter Heese
# Email: pheese@hotmail.com
# 
# Adds IIS Web-Bindings and SSL Bindings
#

$Erroractionpreference = "SilentlyContinue"

## Function display help
Function Help{
	write-host "Run script with the following parameters:"  -ForegroundColor yellow
	write-host " -w      Web Site Name" -ForegroundColor yellow
	Write-host " -ip     IP Address (use 0.0.0.0 to assign *)" -ForegroundColor yellow
	Write-host " -p      Port" -ForegroundColor yellow
	Write-host "Example:" -ForegroundColor yellow
	Write-Host $($MyInvocation.MyCommand.Name + " -w " + [char]34 + "Default Site" + [char]34 + " -ip 192.168.1.23 -p 8080" + [char]10) -ForegroundColor yellow
	Exit 1}

## Function Returns true if admin privileges are used
function CheckAdminPrivileges(){
     $WindowsIdentity=[System.Security.Principal.WindowsIdentity]::GetCurrent()
     $SecurityPrincipal=new-object System.Security.Principal.WindowsPrincipal($WindowsIdentity)
     $Administrator=[System.Security.Principal.WindowsBuiltInRole]::Administrator
     $bIsAdmin=$SecurityPrincipal.IsInRole($Administrator)
     if ($bIsAdmin -eq $False) {
		write-host $("The scripts needs to run with administrative privileges to deploy windows roles in Azure VM.") -foregroundcolor red
		Exit 1
	 }
}

Function AddIISBinding{
	param($WebSite, $IP, $Port)
	
	# Add Binding for WebSite
	$WebBinding = get-webbinding -port $Port -name $WebSite
	If ($WebBinding -ne $null) {
		write-host @("Web Binding for IP Address: " + $IP + " port: " + $Port + " WebSite: " + $WebSite +" already exists.") -Foregroundcolor green}	
	Else
	{
        write-host @("Create Web Binding for IP Address: " + $IP + " and port: " + $Port + " WebSite: " + $WebSite) -Foregroundcolor green
        if ($IP.Equals("0.0.0.0"))
        {
		    New-WebBinding -name $WebSite -port $Port -protocol HTTP
        }
        else
        {
		    New-WebBinding -name $WebSite -IP $IP -port $Port -protocol HTTP
        }
    }
}

## Function main 
$WebSite = ""
$IP = ""
$Port = 0
# $bError = 0 No Errors occured
# $bError = 1 Display Help
# $bError >= 2 Errors occured
$bError = 0
# Parse command line
for ($i = 0; $i -le $Args.count; $i++)
	{
	If ($Args[$i].Length -gt 0)
		{
		Switch ($Args[$i].ToLower())
			{
			"-w" {$i = $i + 1
					$WebSite =  [String]$Args[$i].ToLower().Trim()}
			"-ip" {$i = $i + 1
					$IP =  [String]$Args[$i].ToLower().Trim()}
			"-p" {$i = $i + 1
					$Port = [String]$Args[$i]
					$Port = $Port.Trim()}
			default {$bError = 1}
			}
		}
	}

# Check if all parameters are set
If ($Args.Count -ne 6) {$bError = 1}
ElseIf ($WebSite.Length -eq 0){$bError = 1}
ElseIf ($IP.Length -eq 0) {$bError = 1}
ElseIf ($Port -eq 0) {$bError = 1}
Else {$bError = 0}

# Check administrative privileges
CheckAdminPrivileges

If ($bError -eq 0)
	# Import web administration module
	{import-module WebAdministration
	IF ($error.count -ne 0) 
		{$bError = 2
		$Msg = "TargetObject:          " + [string]$Error[0].TargetObject + [char]10
		$Msg = $Msg + "CategoryInfo:          " + [string]$Error[0].CategoryInfo + [char]10
		$Msg = $Msg + "InvocationInfo:        " + [string]$error[0].InvocationInfo + [char]10
		$Msg = $Msg + "Errordetails:          " + [string]$Error[0].ErrorDetails + [char]10
		$Msg = $Msg + "FullyQualifiedErrorID: " + [string]$error[0].FullyQualifiedErrorID + [char]10
		$Msg = $Msg + "Exception:             " + [string]$Error[0].Exception + [char]10
		write-host $Msg -Foregroundcolor red
		import-module WebAdministration
		$Error.Clear()}
	# Call Function to assign certificate
	If ($bError -eq 0)
		{
            $IISWebSite = Get-WebSite -name $WebSite
            if ($IISWebSite -eq $null) 
            {
                write-host @("Website: " +  $WebSite + " not found.") -ForegroundColor Red
                Exit 1
            }
            else
            {
                $bError = AddIISBinding $WebSite $IP $Port
            }
        }
	ElseIf ($bError -eq 1)
		{Help}
	}
Else {Help}