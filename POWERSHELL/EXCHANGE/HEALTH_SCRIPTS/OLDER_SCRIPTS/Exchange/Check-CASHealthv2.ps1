##### Check-CASHealthv2.ps1
#####
##### FinickyAdmin.com
##### Created Date:    2015/08/17
##### Modified Date:   2014/08/17
#####
##### Using a preset list of Exchange 2013 virtual directories, this script checks for the Exchange Managed Availability
##### Health Probes on a per-server basis.
#####
##### Reference: http://blogs.technet.com/b/exchange/archive/2014/03/05/load-balancing-in-exchange-2013.aspx

# Parameter handler

<#
.SYNOPSIS
Check-CASHealthv2

.DESCRIPTION
Check-CASHealthv2 is a multifunctional tool to check 1 or all Exchange servers in the environment for properly functioning IIS sites.

.PARAMETER ServerName
ServerName

.PARAMETER DAGName
DAGName

.PARAMETER AllServers
AllServers

.EXAMPLE

C:\PS> Check-CASHealthv2 -ServerName "Server01"

.EXAMPLE

C:\PS> Check-CASHealthv2 -DAGName "DAG01"

.EXAMPLE

C:\PS> Check-CASHealthv2 -AllServers $true

#>

Param
(
    [parameter(Mandatory=$true,ParameterSetName="Server")]
    [string] $ServerName,

    [parameter(Mandatory=$false,ParameterSetName="DAG")]
    [string] $DAGName,

    [parameter(Mandatory=$false,ParameterSetName="BoolAllServers")]
    [boolean] $AllServers
)


if ($Server -eq "" -and $DAGName -eq "" -and $AllServers -eq $false) {Write-Host -ForegroundColor yellow "Parameters missing.  See help for usage.";break}

# Healthcheck Site Hash Table - Exchange 2013 SP1 (CU8)

$Sites=@{"Exchange ActiveSync"="Microsoft-Server-ActiveSync"
         "Exchange Admin Center"="ECP"
         "Exchange Autodiscover"="Autodiscover"
         "Offline Address Book"="oab"
         "Outlook Anywhere"="rpc"
         "Outlook Web Access"="owa"
         "Exchange Web Services"="EWS"
        }


# Function DisplayResults
# Create an object containing the results for future use

function DisplayResults
{
    param([string]$Site,[string]$URL,[string]$Value)
  
    $output = New-Object PSObject
    $output | Add-Member NoteProperty Test($Site)
    $output | Add-Member NoteProperty URL($URL)
    $output | Add-Member NoteProperty Value($Value)
    Write-Output $output
}  

# Function RunTests
# Performs the tests for each member of the Sites hashtable.

function RunTests
{
    param([string]$Server)

    # For each object in the hash table, run the site healthcheck
    foreach($Site in $Sites.GetEnumerator())
    {
        # Define the test
        $Test = $($Site.Key)

        # Define the test URL
        $URL = "https://$Server/$($Site.Value)/healthcheck.htm"
        
        # Prepare the response
        $Response = $null
    
        # Try/Catch to handle web exceptions
        Try
        {
            # Create a new WebClient object which will handle getting the web page
            $webclient = New-Object System.Net.WebClient
            # Add the user-agent header to make searching through logs easier.
            $webclient.Headers.Add("user-agent","Check-CASHealth.ps1 script check")
            
            # Declare object for webpage and set it to null
            $webpage = $null

            # Use the WebClient.DoanloadString method to download the page as a string for parsing.
            $webpage = $webclient.DownloadString($URL)
            
            # Handles the results of a $null (empty) return.
            if($webpage -ne $null)
            {
                # Since the grab was successful, create the Response value by replacing the HTML line break
                $Response = $webpage.Replace("<br/>"," ")
            }
        }
        Catch [System.Net.WebException]
        {
            # Since the grab failed, set the Response value to the exception message
            $Response = "Error:`t" + $_.Exception.Message
        }
        Finally
        {
            # Finally, send the results to the object builder to handle returning the object
            DisplayResults $Test $URL $Response
        }
    }
}

function DetermineTests
{
    if($ServerName.Length -gt 0)
    {
        Write-Host "Server name provided, going to run with 1 server."
        
        ## Call RunTests($ServerName)
        RunTests($ServerName)
    }
    elseif ($DAGName.Length -gt 0)
    {
        Write-Host "DAG name provided, enumerating DAG servers."

        ## Enumerate DAG servers and call RunTests() with them
        $DAG = Get-DatabaseAvailabilityGroup $DagName -Status
        foreach($DAGMember in $DAG.StartedMailboxServers)
        {
            RunTests([string]$DAGMember)
        }
    }
    elseif ($AllServers -ne $false)
    {
        Write-Host "Status for all Exchange servers requested, enumerating Exchange servers."
        
        ## Enumerate all Exchange servers with CAS role and call RunTests() on them
        #$CASServers = Get-ExchangeServer | where {$_.ServerRole -cmatch "ClientAccess"}
        $CASServers = Get-ClientAccessServer
        foreach($CASServer in $CASServers)
        {
            RunTests($CASServer)
        }
    }
    else
    {
        Write-Host "Options were not selected.  Reference the help if you need to."
        break;
    }
}

function IgnoreIISCert
{

    # IIS Ignore Certificate Type Definition
    # Reference: http://stackoverflow.com/questions/11696944/powershell-v3-invoke-webrequest-https-error
    # Reference: http://connect.microsoft.com/PowerShell/feedback/details/419466/new-webserviceproxy-needs-force-parameter-to-ignore-ssl-errors
    # Checks for type definition and if missing adds it
    # Handles untrusted certs (useful for self-signed certs and test environments)

    if(-not ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type)
    {
        add-type @"
            using System.Net;
            using System.Security.Cryptography.X509Certificates;
            public class TrustAllCertsPolicy : ICertificatePolicy
            {
                public bool CheckValidationResult(ServicePoint srvPoint,X509Certificate certificate,WebRequest request, int certificateProblem)
                {
                    return true;
                }
            }
"@ # It's incredibly annoying that you can't allign this with a tab!
    }

    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

}

##  Run the function IgnoreIISCert to prevent cert issues (from locally generated or outdated certificates)
IgnoreIISCert

##  Run the function DetermineTests to determine which cmdlet parameters are going to be used, and launch the RunTests function accordingly
DetermineTests
