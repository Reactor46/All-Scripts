<#
.Synopsis
        This script will Poll all the webservers listed in the webservers.txt file.  and get connection stats for each individual site.
        You can exlcude sites by adding them to the Exclude.txt file. Both files are located in C:\Scripts\Configs on the system that is running the script
        Unless changed in code.
        Once the polling is complete a HTML file is outputed with the results.  By default the Polling has a 5 second wait this can be adjusted by changing the 
        value of $PollingTimer
        You can change the path of the output file by changing the value of $WebOutFile

.Author
        Jim Adkins
        CreditOneBank Systems Administrator II
        James.Adkins@CreditOne.com
        04/07/17





.Change Log
5/11/17 Added code to change the background color of the connections field to yellow or red based on the number of connections
5/12/17 Added code to better calculate the number of users.  it origianly only pulled the CurrentAnonymouseUsers now it pulls that counter
and  CurrentNonAnonymousUsers to get an acurate count.
5/15/17 Code Cleanup.  removed old code that was commented out.  
Added code to remove the _ from the Totals website so when _Total is removed from exclude it will not show the _ in the site name.
#>

Clear-Host
Write-Host "Running Connection Monitoring Scripts" -ForegroundColor Yellow
$TimeToRun=$null
$WeboutFile2 = 'C:\Scripts\Repository\jbattista\Web\Reports\TotalWebCon_MT.html'
$bColor= New-Object System.Collections.Generic.List[System.Object]
$bColor.Add('#42C8DA')
$bColor.Add('#42DAB7')
$PollingTimer=5 #Number of seconds to wait before polling the servers again.
$currentDateTime=Get-Date
#Invoke-Expression -Command C:\Scripts\WebConTotals.ps1  #Runs the Totals Only Script in the background.


function PollServers{
$i=0
$x=0
$currentDateTime=Get-Date
$ExcludeList=Get-Content C:\Scripts\Repository\jbattista\Web\AdminScripts\Configs\Exclude.txt
$WebServers=Get-Content C:\Scripts\Repository\jbattista\Web\AdminScripts\Configs\Servers.txt
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
       "

$htmlTable="
        
                
				<TD><B>Site</B></td>
                <TD><B>Users</B></TD>
                <TD><B>Connections</B></TD>
 "
            
$TimeToRun=Measure-Command{

    foreach($Webserver in $Webservers){
        
        if(Test-Connection $Webserver){
        $currentDateTime= Get-Date
        $WebInfo = Get-WmiObject -Class Win32_PerfFormattedData_W3SVC_WebService -ComputerName $WebServer | select Name, CurrentAnonymousUsers,CurrentConnections,CurrentNonAnonymousUsers

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