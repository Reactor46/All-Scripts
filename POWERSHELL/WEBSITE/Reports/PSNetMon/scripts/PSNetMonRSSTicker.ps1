########################################################################  Powershell Network Monitor RSS Feed Ticker
#  Created by Brad Voris#  This script is used to generate results for the rssticker.html page
#  Need- run from CFG file, multiple RSS feeds, convert feeds to links
#        remove star from feed########################################################################Script variables
$Dated = (Get-Date -format F)
$RSSFeedstreams = Get-Content "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\rssfeedstreams.cfg"

#1st RSS Feed Stream
$WebClient01 = New-Object system.net.webclient
$RSSFeed01 = [xml]$WebClient01.DownloadString('http://news.google.com/?output=rss')
$RSSFeedVar01 = $RSSFeed01.rss.channel.item | Select-Object title -First 5 | ConvertTo-Html
#2nd RSS Feed Stream
$WebClient02 = New-Object system.net.webclient
$RSSFeed02 = [xml]$WebClient02.DownloadString('http://news.google.com/?output=rss')
$RSSFeedVar02 = $RSSFeed02.rss.channel.item | Select-Object title -First 5 | ConvertTo-Html

#CSS Coding
$CSS = Get-Content "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\css\rss.css "

#HTML Header Coding
$HTMLHead = @"
<!DOCTYPE html>
<HEAD>
<META charset="UTF-8">
<TITLE>PSNetMon - RSS Ticker Module</TITLE>
<CENTER>
<STYLE>$CSS</STYLE></HEAD>
"@

#HTML Body Coding
$HTMLBody = @"<CENTER><TABLE  border="0"><TR bgcolor=#5CB3FF><TD><MARQUEE Behavior="scroll" Direction="left" ScrollAmount="3">$RSSFeedVar01</MARQUEE></TD>
</TR>
<TR BGCOLOR=#5CB3FF>
<TD><MARQUEE Behavior="scroll" Direction="left" ScrollAmount="3">$RSSFeedVar02</MARQUEE></TD></TR></TABLE><I>Script last run:$Dated</I></CENTER>"@

#Export to HTML
$Script | ConvertTo-HTML -Head $HTMLHead -Body $HTMLBody | Out-file "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\gen\rssticker.html"

