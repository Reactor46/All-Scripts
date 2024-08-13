function WMIDateStringToDate($Bootup) {    
    [System.Management.ManagementDateTimeconverter]::ToDateTime($Bootup)    
} 
 
function Test-Port{    
[cmdletbinding(    
    DefaultParameterSetName = '',    
    ConfirmImpact = 'low'    
)]    
    Param(    
        [Parameter(    
            Mandatory = $True,    
            Position = 0,    
            ParameterSetName = '',    
            ValueFromPipeline = $True)]    
            [array]$computer,    
        [Parameter(    
            Position = 1,    
            Mandatory = $True,    
            ParameterSetName = '')]    
            [array]$port,    
        [Parameter(    
            Mandatory = $False,    
            ParameterSetName = '')]    
            [int]$TCPtimeout=1000,    
        [Parameter(    
            Mandatory = $False,    
            ParameterSetName = '')]    
            [int]$UDPtimeout=1000,               
        [Parameter(    
            Mandatory = $False,    
            ParameterSetName = '')]    
            [switch]$TCP,    
        [Parameter(    
            Mandatory = $False,    
            ParameterSetName = '')]    
            [switch]$UDP                                      
        )    
    Begin {    
        If (!$tcp -AND !$udp) {$tcp = $True}    
        #Typically you never do this, but in this case I felt it was for the benefit of the function    
        #as any errors will be noted in the output of the report            
        $ErrorActionPreference = "SilentlyContinue"    
        $report = @()    
    }    
    Process {       
        ForEach ($c in $computer) {    
            ForEach ($p in $port) {    
                If ($tcp) {      
                    #Create temporary holder     
                    $temp = "" | Select Server, Port, TypePort, Open, Notes    
                    #Create object for connecting to port on computer    
                    $tcpobject = new-Object system.Net.Sockets.TcpClient    
                    #Connect to remote machine's port                  
                    $connect = $tcpobject.BeginConnect($c,$p,$null,$null)    
                    #Configure a timeout before quitting    
                    $wait = $connect.AsyncWaitHandle.WaitOne($TCPtimeout,$false)    
                    #If timeout    
                    If(!$wait) {    
                        #Close connection    
                        $tcpobject.Close()    
                        Write-Verbose "Connection Timeout"    
                        #Build report    
                        $temp.Server = $c    
                        $temp.Port = $p    
                        $temp.TypePort = "TCP"    
                        $temp.Open = "False"    
                        $temp.Notes = "Connection to Port Timed Out"    
                    } Else {    
                        $error.Clear()    
                        $tcpobject.EndConnect($connect) | out-Null    
                        #If error    
                        If($error[0]){    
                            #Begin making error more readable in report    
                            [string]$string = ($error[0].exception).message    
                            $message = (($string.split(":")[1]).replace('"',"")).TrimStart()    
                            $failed = $true    
                        }    
                        #Close connection        
                        $tcpobject.Close()    
                        #If unable to query port to due failure    
                        If($failed){    
                            #Build report    
                            $temp.Server = $c    
                            $temp.Port = $p    
                            $temp.TypePort = "TCP"    
                            $temp.Open = "False"    
                            $temp.Notes = "$message"    
                        } Else{    
                            #Build report    
                            $temp.Server = $c    
                            $temp.Port = $p    
                            $temp.TypePort = "TCP"    
                            $temp.Open = "True"      
                            $temp.Notes = ""    
                        }    
                    }       
                    #Reset failed value    
                    $failed = $Null        
                    #Merge temp array with report                
                    $report += $temp    
                }        
                If ($udp) {    
                    #Create temporary holder     
                    $temp = "" | Select Server, Port, TypePort, Open, Notes                                       
                    #Create object for connecting to port on computer    
                    $udpobject = new-Object system.Net.Sockets.Udpclient  
                    #Set a timeout on receiving message   
                    $udpobject.client.ReceiveTimeout = $UDPTimeout   
                    #Connect to remote machine's port                  
                    Write-Verbose "Making UDP connection to remote server"   
                    $udpobject.Connect("$c",$p)   
                    #Sends a message to the host to which you have connected.   
                    Write-Verbose "Sending message to remote host"   
                    $a = new-object system.text.asciiencoding   
                    $byte = $a.GetBytes("$(Get-Date)")   
                    [void]$udpobject.Send($byte,$byte.length)   
                    #IPEndPoint object will allow us to read datagrams sent from any source.    
                    Write-Verbose "Creating remote endpoint"   
                    $remoteendpoint = New-Object system.net.ipendpoint([system.net.ipaddress]::Any,0)   
                    Try {   
                        #Blocks until a message returns on this socket from a remote host.   
                        Write-Verbose "Waiting for message return"   
                        $receivebytes = $udpobject.Receive([ref]$remoteendpoint)   
                        [string]$returndata = $a.GetString($receivebytes)  
                        If ($returndata) {  
                           Write-Verbose "Connection Successful"    
                            #Build report    
                            $temp.Server = $c    
                            $temp.Port = $p    
                            $temp.TypePort = "UDP"    
                            $temp.Open = "True"    
                            $temp.Notes = $returndata     
                            $udpobject.close()     
                        }                         
                    } Catch {   
                        If ($Error[0].ToString() -match "\bRespond after a period of time\b") {   
                            #Close connection    
                            $udpobject.Close()    
                            #Make sure that the host is online and not a false positive that it is open   
                            If (Test-Connection -comp $c -count 1 -quiet) {   
                                Write-Verbose "Connection Open"    
                                #Build report    
                                $temp.Server = $c    
                                $temp.Port = $p    
                                $temp.TypePort = "UDP"    
                                $temp.Open = "True"    
                                $temp.Notes = ""   
                            } Else {   
                                <#   
                                It is possible that the host is not online or that the host is online,    
                                but ICMP is blocked by a firewall and this port is actually open.   
                                #>   
                                Write-Verbose "Host maybe unavailable"    
                                #Build report    
                                $temp.Server = $c    
                                $temp.Port = $p    
                                $temp.TypePort = "UDP"    
                                $temp.Open = "False"    
                                $temp.Notes = "Unable to verify if port is open or if host is unavailable."                                   
                            }                           
                        } ElseIf ($Error[0].ToString() -match "forcibly closed by the remote host" ) {   
                            #Close connection    
                            $udpobject.Close()    
                            Write-Verbose "Connection Timeout"    
                            #Build report    
                            $temp.Server = $c    
                            $temp.Port = $p    
                            $temp.TypePort = "UDP"    
                            $temp.Open = "False"    
                            $temp.Notes = "Connection to Port Timed Out"                           
                        } Else {                        
                            $udpobject.close()   
                        }   
                    }       
                    #Merge temp array with report                
                    $report += $temp    
                }                                    
            }    
        }                    
    }    
    End {    
        #Generate Report    
        $report   
    }  
} 

#### Symantec Enterprise Protection ####
function Get-SEPVersion {
<#
.SYNOPSIS
Retrieve Symantec Endpoint Version, Definition Date and Sylink Group

.DESCRIPTION
Retrieve Symantec Endpoint Version, Definition Date and Sylink Group

.PARAMETER  ComputerName
Name of the computer to query SEP info for

.EXAMPLE
PS C:\> Get-SEPVersion -ComputerName Server01

.EXAMPLE
PS C:\> $servers | Get-SEPVersion

.NOTES
Author: Jonathan Medd
Date: 23/12/2011
#>

[CmdletBinding()]
param(
[Parameter(Position=0,Mandatory=$true,HelpMessage="Name of the computer to query SEP for",
ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
[Alias('CN','__SERVER','IPAddress','Server')]
[System.String]
$ComputerName
)

begin {
# Create object to enable access to the months of the year
$DateTimeFormat = New-Object System.Globalization.DateTimeFormatInfo

# Set Registry keys to query
$SMCKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC" 
$AVKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\AV" 
$SylinkKey = "SOFTWARE\\Wow6432Node\\Symantec\\Symantec Endpoint Protection\\SMC\\SYLINK\\SyLink" 
}


process {


try {

# Connect to Registry
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$ComputerName)

# Obtain Product Version value
$SMCRegKey = $reg.opensubkey($SMCKey)
$SEPVersion = $SMCRegKey.GetValue('ProductVersion')

# Obtain Pattern File Date Value
$AVRegKey = $reg.opensubkey($AVKey)
$AVPatternFileDate = $AVRegKey.GetValue('PatternFileDate')

# Convert PatternFileDate to readable date
$AVYearFileDate = [string]($AVPatternFileDate[0] + 1970)
$AVMonthFileDate = $DateTimeFormat.MonthNames[$AVPatternFileDate[1]]
$AVDayFileDate = [string]$AVPatternFileDate[2]
$AVFileVersionDate = $AVDayFileDate + " " + $AVMonthFileDate + " " + $AVYearFileDate

# Obtain Sylink Group value
$SylinkRegKey = $reg.opensubkey($SylinkKey)
$SylinkGroup = $SylinkRegKey.GetValue('CurrentGroup')

}

catch [System.Management.Automation.MethodInvocationException]

{
$SEPVersion = "Unable to connect to computer"
$AVFileVersionDate = ""
$SylinkGroup = ""
}

$MYObject = “” | Select-Object ComputerName,SEPProductVersion,SEPDefinitionDate,SylinkGroup
$MYObject.ComputerName = $ComputerName
$MYObject.SEPProductVersion = $SEPVersion
$MYObject.SEPDefinitionDate = $AVFileVersionDate
$MYObject.SylinkGroup = $SylinkGroup
$MYObject

}
}