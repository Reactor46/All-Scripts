<#

Version:  1.0

Created:  February 16, 2014

Summary:  

This script uses the nslookup.exe utility to query the DNS servers listed in the
$DNSservers variable for the DNS A record of the host listed in the $Arecord
variable.

Here are the steps to run this script.

1.  Open the script in PowerShell ISE.
2.  Modify the $DNSservers array so that each DNS server to query is listed in
double quotes and that there is also a comma separating each one.
3.  Modify the  $Arecord variable to be the host name to lookup on the DNS servers 
listed in the $DNSservers array.
4.  Execute the script.

This script can also take parameters from the command line.  Here are the steps.

1.  Open the script in PowerShell ISE.
2.  Modify the $DNSservers array so that each DNS server to query is listed in
double quotes and that there is also a comma separating each one.
3.  Save the script, and close the PowerShell ISE.
4.  From the PowerShell CLI, change to the directory where the script is saved.
Then, execute the script as per the following.

.\LookupARecord.ps1 www.example.com

Where www.example.com is the fully qualified domain name to lookup.  If no 
host name is entered, then www.microsoft.com will be queried.
#>

# The default host record to query is www.microsoft.com.  This script can be run from
# the PowerShell ISE by changing this value, or from the command line by specifying the
# host name as a parameter.
Param (
$Arecord = "www.microsoft.com"
)

# Enter the DNS servers to query.
$DNSservers = "8.8.8.8", "8.8.4.4"

# Query each DNS server for the A record.
Foreach ($DNSServer in $DNSservers)
{
    $ErrorActionPreference = "SilentlyContinue"
    Write-Host "
Querying $DNSserver for $Arecord"
    $QueryDNSServer = nslookup.exe $Arecord $DNSServer
    # $Error[1].Exception
    $DNSQueryResults = $QueryDNSServer | Select-Object -skip 3
    $DNSQueryResults
    $ErrorActionPreference = "Continue"
}
