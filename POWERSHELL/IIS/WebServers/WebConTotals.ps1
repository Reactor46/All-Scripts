
$WeboutFile = 'C:\inetpub\wwwroot\TotalsOnly.html'
$PollingTimer=5 #Number of seconds to wait before polling the servers again.

function PollServers{
$WebServers=Get-Content C:\LazyWinAdmin\WebServers\Configs\webservers.txt
$TimeToRun=Measure-Command{
$currentDateTime = Get-Date
$HTMLReport="
<HTML>
    <TITLE> CreditOne Bank Total Web Connections </TITLE> 
        <header>
            <meta http-equiv=""refresh"" content=""5"" >
        </header>
    <BODY background-color:white> 
        <font color =""#0099ff"" face=""Microsoft Tai le"" > 
       <H2> Total Web Connections </H2></font> 
        <Body><Div id=""wrapper""><div id=""content"">
        <Table border=1 cellpadding=0 cellspacing=0 width=""350"" font face=""Microsoft Tai le"" size=""6""> 
            <TR bgcolor=white align=center> 
            <TD><B>Last Polling Time</B></TD> 
            <TD><B>" + $currentDateTime + "</B></TD>
       </Table>
       <br>&nbsp;
       <Table border=1 cellpadding=0 cellspacing=0 width=""600"" font face=""Microsoft Tai le"" size=""6""> 
       <TR bgcolor=white align=center> 
       <TD><B>Server Name</B></TD> 
       <TD><B>Users</B></TD>
       <TD><B>Connections</B></TD></TR>
 "






ForEach($WebServer in $WebServers){


$SiteData = Get-WmiObject -Class Win32_PerfFormattedData_W3SVC_WebService -ComputerName $WebServer |Where {$_.Name -eq "_Total"} | select  CurrentAnonymousUsers,CurrentConnections,CurrentNonAnonymousUsers
$TotalUsers=$SiteData.CurrentAnonymousUsers + $SiteData.CurrentNonAnonymousUsers

if($SiteData.CurrentConnections -ige 1700 -and $SiteData.CurrentConnections -ile 1799) {$BackColor = "Yellow"}
            if($SiteData.CurrentConnections -ige 1800){$BackColor="Red"}
$HTML_Add = "<tr><td><center>" + $Webserver.ToUpper() + "</td><td bgcolor="+$BackColor+"><center>" + $TotalUsers + "</td><td bgcolor="+$BackColor+"><center>" + $SiteData.CurrentConnections +"</td></center></tr>"
$HTMLReport=$HTMLReport+$HTML_Add
$BackColor=""
$TotalUsers=0
}}
$HTMLReport=$HTMLReport+"</table><Div id=""footer"">Script Time " + $TimeToRun.TotalSeconds + " Seconds</div></div></div>"
$HtmlReport | Out-File $WeboutFile
Start-Sleep -Seconds $PollingTimer
}

while($true){
    PollServers
}