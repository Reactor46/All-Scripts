########################################################################  Powershell NetMon Port Monitor Script#  Create by Brad Voris
#  Script scans port to check status if available#######################################################################
#Variables$CSS = get-content "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\css\theme.css"$Dated = (Get-Date -format F)
#HTML Head
$HTMLHead = @"
<!DOCTYPE html>
<HEAD>
<META charset="UTF-8">
<TITLE>PSNetMon - Port Monitor Module</TITLE>
<CENTER>
<STYLE>$CSS</STYLE></HEAD>
"@
#  Get content from config files$hostida = (get-content  "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\cfg\porthostsa.cfg" | Where {-not ($_.StartsWith('#'))})$portnumbera = (get-content "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\cfg\portcfga.cfg" | Where {-not ($_.StartsWith('#'))})$hostidb = (get-content  "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\cfg\porthostsb.cfg" | Where {-not ($_.StartsWith('#'))})$portnumberb = (get-content "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\cfg\portcfgb.cfg" | Where {-not ($_.StartsWith('#'))})
#Check status of ports single port A$socketa = new-object Net.Sockets.TcpClient$socketresultsa = $socketa.Connect($hostida, $portnumbera)if ($socketa.Connected) {$statusa = “Port Open”$socketa.Close()}else {$statusa = “Port Closed”}
#HTML Body Content
$HTMLBody = @"<CENTER><meta http-equiv="refresh" content="30">
<Font size=4><B>Monitored Ports</B></font></BR><I>Script last run:$dated</I><BR /><TABLE><TR bgcolor=#2554C7><TD>Hosts</TD> <TD>Port Number</TD> <TD>Status</TD></TR><TR bgcolor=#5CB3FF><TD>$hostida</TD> <TD>$portnumbera</TD> <TD>$statusa</TD></TR></TABLE></CENTER>"@
#Check status of ports single port B$socketb = new-object Net.Sockets.TcpClient$socketresultsb = $socketb.Connect($hostidb, $portnumberb)if ($socketb.Connected) {$statusb = “Port Open”$socketb.Close()}else {$statusb = “Port Closed”}
#HTML Body Content
$HTMLBody = $HTMLBody + @"<CENTER><BR /><TABLE><TR bgcolor=#2554C7><TD>Hosts</TD> <TD>Port Number</TD> <TD>Status</TD></TR><TR bgcolor=#5CB3FF><TD>$hostidb</TD> <TD>$portnumberb</TD> <TD>$statusb</TD></TR></TABLE></CENTER>"@
# CSV for reports$countresults1=@"$Dated,$hostida,$portnumbera,$statusa,$Dated,$hostidb,$portnumberb,$statusb"@$countresults1 | Add-Content "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\reports\ports.csv"
#$countresults2=@"#$Dated,$hostidb,$portnumberb,$statusb#"@#$countresults2 | Add-Content "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\reports\ports.csv"
#Export to HTML$statusupdate | ConvertTo-HTML -head $HTMLHead -body $HTMLBody | out-file "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\gen\porthosts.htm "