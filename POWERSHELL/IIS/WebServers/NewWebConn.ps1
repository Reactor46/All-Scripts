## This script gets the current connections of all WEB and CAS servers

##1/3/17 - Changed script drasticly to improve performance combined Web and CAS servers changed to 1 foreach loop.  channged to a looping website write.

function getData  
{
$webServers = "phxweb07.phx.Contoso.CORP","phxweb08.phx.Contoso.corp","phxweb09.phx.Contoso.corp"#,"lasweb01", "lasweb02", "lasweb03", "lasweb04", "lasweb05","lasweb06","lasweb07","lasweb08","lasweb10","lasweb11","lasweb12","lascas01","lascas02","lascas03","lascas04","lascas05","lascas06","lascas07","lascas08","lasdmzweb01","lasdmzweb02","lasdmzweb03","lasdmzweb04","lasdmzweb05","lasdmzweb06","lasdmzweb07","lasdmzweb08","lasdmzweb09","lasdmzweb10","lasdmzweb11","lasdmzweb12"
# , "lasweb05", "lasweb09", "lasweb13".phx.Contoso.corp

$webServersUserCount = New-Object System.Collections.Generic.List[System.Object]
$webServersConnectionCount = New-Object System.Collections.Generic.List[System.Object]
    
$currentDateTime = Get-Date
$bColor= New-Object System.Collections.Generic.List[System.Object]
##New-Object System.Collections.ArrayList

foreach ($webServer in $webServers)
{
    
    $namespace = "root\CIMV2"
    $userSessions = Get-wmiObject -class Win32_PerfRawData_W3SVC_WebService -computername $webServer -namespace $namespace | select-object -expand currentanonymoususers

    $connectioncount = $userSessions[0]
    $WebServersUserCount.Add($connectioncount)

    $connectioncount = 0

    $Sessions = Get-WmiObject -Class Win32_PerfFormattedData_W3SVC_WebService -ComputerName $webServer | Where {$_.Name -eq "_Total"} | % {$_.CurrentConnections}

    $connectioncount = $Sessions[0]
    $WebServersConnectionCount.Add($connectioncount)

    if($connectioncount -gt 1699 -and $connectioncount -lt 1800){$bColor.add('yellow')}  
    if($connectioncount -gt 1799){$bColor.add('red')}       
    if($connectioncount -lt 1699) {$bColor.Add('white')}


}

$Outputreport = "
<HTML>
    <TITLE> CreditOne Bank Total Web Connections </TITLE> 
        <header>
            <meta http-equiv=""refresh"" content=""5"" >
        </header>
    <BODY background-color:white> 
        <font color =""#0099ff"" face=""Microsoft Tai le"" > 
       <H2> Total Web Connections </H2></font> 
        

        
        <Table border=1 cellpadding=0 cellspacing=0 width=""350"" font face=""Microsoft Tai le"" size=""6""> 
            <TR bgcolor=white align=center> 
            <TD><B>Last Polling Time</B></TD> 
            <TD><B>" + $currentDateTime + "</B></TD>
       </Table>
       <br>&nbsp;
       
        <Table border=1 cellpadding=0 cellspacing=0 width=""600"" font face=""Microsoft Tai le"" size=""6""> 
             <TR bgcolor=white align=center> 
                <TD><B>Server Name</B></TD> 
                <TD><B>Active Users</B></TD>
                <TD><B>Active Connections</B></TD>
            </TR>"
   
   
    for($i=0; $i -le $webServers.Count ; $i++){
    $WebBody = $WebBody +  "
            <TR bgcolor=white align=center> 
                <TD bgcolor="+ $bColor[$i] +"><B>"+ $webServers[$i].toupper() +"</B></TD>
                <TD bgcolor="+ $bColor[$i] +"><B>" + $webServersUserCount[$i] + "</B></TD>
                <TD bgcolor="+ $bColor[$i] +"><B>" + $webServersConnectionCount[$i] + "</B></TD>  
            </TR>"
   }
              

$website = $Outputreport + $WebBody 

$website | out-file 'C:\temp\WebServerConnectionsNEW.html'
#|out-file '\\lasinfra02\c$\WebConnections\WebServerConnections2.html'  
#out-file '\\lasinfra02\c$\WebConnections\WebServerConnections2.html'  
#out-file 'C:\temp\WebServerConnectionsNEW.html'
#out-file '\\lasinfra02\c$\WebConnections\WebServerConnections.html'  

}

while($true)
{
   getdata
   
}