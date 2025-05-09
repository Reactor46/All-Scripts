#########################################################################  Powershell Network Monitor ICMP Script #  Created by Brad Voris#  Script pings resources to determine if they are up#########################################################################Varibales$CSS = get-content "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\css\theme.css"$dated = (Get-Date -format F)$HTMLHead =  "<Style>$CSS</style>"$HTMLHead = $HTMLHead + "<CENTER><Font size=4><B>Monitored Hosts</B></font></BR>"
#Check content and ping computers$icmpresults = get-content "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\cfg\computers.cfg" | Where {-not ($_.StartsWith('#'))} | foreach {  if (test-connection $_ -quiet -count 2)  { New-Object psobject -Property @{   	Server = $_   	Status = "Online"   } }  else  {New-Object psobject -Property @{   	Server = $_   	Status = "DEAD HOST" 
#Email Notification Script#If (Status -eq "DEAD HOST") #{$smtpbody = "$machineName, $svcName, $svcState"#$smtpto = (get-content "C:\inetpub\wwwroot\cfg\sendmail.cfg")[1]#$smtpfrom = (get-content "C:\inetpub\wwwroot\cfg\sendmail.cfg")[3]#$smtpsubject = "PSNetmon Notification Host Down. $dated"#$smtpservername = (get-content "C:\inetpub\wwwroot\cfg\sendmail.cfg")[9]#Send-MailMessage -To $smtpto -From $smtpfrom -subject $smtpsubject -Body $smtpbody -SMTPServer $smtpservername
# }  
   }
 }
}
$HTMLBody = @"<meta http-equiv="refresh" content="30">
<CENTER><I>Script last run:$dated</I><BR /></CENTER>"@
# CSV for reports$countresults1=@()$dated,$Icmpresults.Server,$Icmpresults.Status | Add-Content "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\reports\icmp.csv"
#  Export to HTML$icmpresults | ConvertTo-HTML -head $HTMLHead -body $HTMLBody -Property Server,Status |
     Foreach { if ($_ -like "*<td>Online</td>*" ) 
                      {$_ -replace "<tr>","<tr bgcolor=#5CB3FF>" }
                    else {$_ -replace "<tr>","<tr bgcolor=red>"}} | 
                    out-file "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\gen\icmphosts.htm"

            
