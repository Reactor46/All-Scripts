<#
.Synopsis
Creates a list of all Servers from AD, checks to see if it is Online.
If it is online then it will get the number of days since the system was last rebooted.

.Author
Jim Adkins
CreditOne Bank 
4/5/2017
James.Adkins@CreditOne.com
WinSys Department.

.Information
    After the script is run it will output the data to a TXT file.  From Excel Import the Data and use SPACE as the delemiter
    inorder to Have it properly formated in excel.

.Change Log

4/7/17
Changed code to have the List written to an Object, then Sort the obhect and write the output in decending order of days since last reboot.

4/10/17 
Added code to color servers over 30 days to Yellow and servers over 45 days to red.
4/24/17 Changed all background color back to white.

#>

$RebootData = @() #Create an Object Array to store all the info in
$TotalRunTime = Measure-Command{
Clear-Host

$HostName=Get-ADComputer -Filter {operatingsystem -Like "*Server*"} | select Name | Where{$_.ipv4address -ne ""}
$HTML_Header="
<HTML>
    <TITLE> Last Reboot Report </TITLE> 
        <header>
            <meta http-equiv=""refresh"" content=""5"" >
        </header>
    <BODY background-color:white> 

        

        
        <Table border=1 cellpadding=0 cellspacing=0 width=""350"" font face=""Microsoft Tai le"" size=""6""> 
            <TR bgcolor=white align=center> 

       </Table>
       <br>&nbsp;
       
        <Table border=1 cellpadding=0 cellspacing=0 width=""600"" font face=""Microsoft Tai le"" size=""6""> 
             <TR bgcolor=white align=center> 
                <TD><B>Server Name</B></TD> 
                <TD><B>Rebooted on</B></TD>
                <TD><B>Days Since Reboot</B></TD>
            </TR>"

ForEach($hst in $HostName){
    Try{
    
        if(Test-Connection $hst.Name -Quiet -Count 1){ #a single PING to make sure the system is online before pulling the information.
        $OS=Get-WmiObject Win32_OperatingSystem -ComputerName $hst.Name | Select  CSName, LastBootupTime, LocalDateTime
        $BootTime=[System.Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootupTime)
        $Uptime=[System.Management.ManagementDateTimeConverter]::ToDateTime($os.LocalDateTime) - $BootTime
        
        ##
        $BootRecord = New-Object -TypeName PSObject
        Add-Member -InputObject $BootRecord -MemberType NoteProperty -Name ServerName -Value $os.CSName
        Add-Member -InputObject $BootRecord -MemberType NoteProperty -Name BootTime -Value $BootTime
        Add-Member -InputObject $BootRecord -MemberType NoteProperty -Name BootDays -Value $Uptime.Days
        $RebootData+=$BootRecord
        }
        }
        Catch [Exception] { 
            $ErrLog = $Hst.Name + $($_.Exception.Message) 
            $Errlog | Out-File C:\Temp\LastBootErr.log -Append
            }
}
$RebootData = $RebootData| Sort-Object -Property BootDays -Descending
foreach($obj in $RebootData){
  #all back colors set to white there were complaints that the coloring made it hard to read.      
        
        if($obj.BootDays -lt 30){$bcolor="White"}
        if($obj.BootDays -gt 29 -and  $obj.BootDays -ile 44) {$bcolor="White"}
        if($obj.BootDays -ige 45){$bcolor="White"}
        
        
        


        $HTMLADD="<TR bgcolor=""$bcolor""><TD>"+ $obj.ServerName +"</td><td>"+ $obj.BootTime + "</td><td>" + $obj.BootDays + "</td></tr>"
        $HTML_Header=$HTML_Header+$HTMLADD


}

Send-MailMessage -to "WinSysAdmin@creditone.com" -from "ServerReboots@Creditone.com" -subject "Server Last Reboot" -BodyAsHtml $HTML_Header -SmtpServer "lasexch01.fnbm.corp"

}

#$HTML_Header | Out-File C:\Temp\BootTime.html
$TotalRunTime 