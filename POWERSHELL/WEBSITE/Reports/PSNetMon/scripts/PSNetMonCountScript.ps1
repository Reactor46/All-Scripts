########################################################################  Powershell Network Monitor Count Script#  Created by Brad Voris
#  This script is used to generate results for the results.htm page#######################################################################
#Script variables$CSS = Get-Content "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\css\theme.css "$Dated = (Get-Date -format F)
#Script commands[array]$i = Get-Content -Path "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\cfg\computers.cfg"[array]$j = Get-Content -Path "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\cfg\porthostsa.cfg"[array]$l = Get-Content -Path "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\cfg\porthostsb.cfg"[array]$k = Get-Content -Path "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\cfg\services.cfg"$ii=$i.length$jj=$j.length$kk=$k.length$ll=$l.length$yy = $jj+$ll
#HTML Header Coding
$HTMLHead = @"
<META http-equiv="refresh" content="30">
<TITLE> 
PSNetMon - Count Module
</TITLE>
<HEAD>
<STYLE>$CSS</STYLE>
</HEAD><CENTER>"@
#HTML Body Coding
$HTMLBody = @"<CENTER>Monitored Resources<BR /><TABLE><TR bgcolor=#2554C7><TD>Hosts</TD> <TD>Ports</TD> <TD>Services</TD></TR><TR bgcolor=#5CB3FF><TD>$ii</TD><TD>$yy</TD><TD>$kk</TD></TR></TABLE><I>Script last run:$dated</I></CENTER>"@
# Export to CSV for reports#$valarray = ( $Dated,$ii,$yy,$kk)$countresults2=@"$Dated,$ii,$yy,$kk"@$countresults2 | Add-Content "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\reports\count.csv"
# Export to HTML$Script | ConvertTo-HTML -Head $HTMLHead -Body $HTMLBody | Out-file "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\gen\count.html"

