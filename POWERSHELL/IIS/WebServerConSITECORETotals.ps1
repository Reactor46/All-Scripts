Clear-Host
Write-Host "Running Connection Monitoring Scripts" -ForegroundColor Yellow
$TimeToRun=$null
$WeboutFile2 = 'D:\SCRIPTS\laspshost-web-files\laspshost-web-files\Web\reports\TotalWebConSC.html'
$bColor= New-Object System.Collections.Generic.List[System.Object]
$bColor.Add('#42C8DA')
$bColor.Add('#42DAB7')
$PollingTimer=5 #Number of seconds to wait before polling the servers again.
$currentDateTime=Get-Date
#Invoke-Expression -Command D:\Scripts\WebServerConSITECORETotals.ps1  #Runs the Totals Only Script in the background.


function PollServers{
$i=0
$x=0
$currentDateTime=Get-Date
$ExcludeList=""
$WebServers= 'fbv-scrcd10-p01','fbv-scrcd10-p02','fbv-scrcd10-p03','fbv-scrcd10-p04'
$HTMLReport="
<HTML>
    <TITLE> SiteCore 10.3 PROD Total Web Connections </TITLE> 
        <header>
            <meta http-equiv=""refresh"" content=""5"" >
        </header>
    <BODY background-color:white> 
        <font color =""#0099ff"" face=""Microsoft Tai le"" > 
       <H2> Total Web Connections </H2></font> 
        <Body><Div id=""wrapper""><div id=""content"">
        <Table {width: auto; border-collapse: collapse;} font face=""Microsoft Tai le"" size=""6""> 
            <TR bgcolor=white align=center> 
            <TD><B>Last Polling Time</B></TD> 
            <TD><B>" + $currentDateTime + "</B></TD>
       </Table>
       <br>&nbsp;
       <Table {width: auto; border-collapse: collapse;} font face=""Microsoft Tai le"" size=""6""> 
       <TR bgcolor=white align=center> 
       <TD><B>Server Name</B></TD> 
       "

$htmlTable="       
                
				<TD><B>Site</B></TD>
                <TD><B>Users</B></TD>
                <TD><B>Connections</B></TD>
 "
            
$TimeToRun=Measure-Command{

    foreach($Webserver in $Webservers){
        
        if(Test-Connection $Webserver){
        $currentDateTime= Get-Date
        $WebInfo = Get-CimInstance -Class Win32_PerfFormattedData_W3SVC_WebService -ComputerName $WebServer | select Name, CurrentAnonymousUsers,CurrentConnections,CurrentNonAnonymousUsers | Where-Object { $_.Name -eq "ksprod-new-cd.ksnet.com" }

        while($x -lt $WebInfo.Count - $ExcludeList.Count  ){
        $HTMLReport=$HTMLReport+$htmlTable
        if($x -ige $WebInfo.Count - $ExcludeList.Count -1){$HTMLReport=$HTMLReport+"</TR>"}
        $x++
   

}
            $HTMLReport=$HTMLReport+"<TR bgcolor="+$bcolor[$i]+"><td>" + $Webserver.ToUpper() + "</td>"
            if($i -eq 1){$i=0} Else {$i=1}
        ForEach($Name in $WebInfo){

            if($Name.Name -notin $ExcludeList){

            if($Name.CurrentConnections -ige 1700 -and $Name.CurrentConnections -ile 1799) {$BackColor = "Yellow"}
            if($Name.CurrentConnections -ige 1800){$BackColor="Red"}
            $TotalUsers=$name.CurrentAnonymousUsers + $Name.CurrentNonAnonymousUsers
                           
                $HTML_Add = "<td><center>" + $Name.Name.split("_") + "</td><td><center>" + $TotalUsers + "</td><td bgcolor="+$BackColor+"><center>" + $Name.CurrentConnections +"</td></center>"
                $HTMLReport=$HTMLReport+$HTML_Add
                $BackColor=""
                $TotalUsers=0
}}
        $HTMLReport=$HTMLReport+"</TR>"
}
}
        $HTMLReport=$HTMLReport+"</table><Div id=""footer"">Script Time " + $TimeToRun.TotalSeconds + " Seconds</div></div></div>"
        $HtmlReport | Out-File $WeboutFile2
        Start-Sleep -Seconds $PollingTimer
        
        }}

        while($true)
{
            PollServers
   
}