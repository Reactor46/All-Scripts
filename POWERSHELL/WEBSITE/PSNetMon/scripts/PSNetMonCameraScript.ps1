###################################################################################################
# PSNetMon - Camera Module
# Written by Brad Voris 
###################################################################################################

$CSS = Get-Content "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\css\theme.css"

#HTML Header Coding
$HTMLHeader = @"
<HTML>
    <HEAD>
    <TITLE> 
    PSNetMon - Camera Module
    </TITLE>
    <STYLE>$CSS</STYLE>
    </HEAD>
"@

#Html Body Coding
$HTMLBody = @"
<BODY>
    <CENTER>
<OBJECT classid="clsid:9BE31822-FDAD-461B-AD51-BE1D1C159921"     codebase="http://downloads.videolan.org/pub/videolan/vlc/latest/win32/axvlc.cab"     width="240" height="180" id="vlc" events="True">   <param name="Src" value="rtsp://<USERNAME>:<PASSWORD>@IPADDRESS5:554//LowResolutionVideo" />   <param name="ShowDisplay" value="True" />   <param name="AutoLoop" value="False" />   <param name="AutoPlay" value="True" />   <embed id="vlcEmb"  type="application/x-google-vlc-plugin" version="VideoLAN.VLCPlugin.2" autoplay="yes" loop="no" width="240" height="180"     target="rtsp://<USERNAME>:<PASSWORD>@IPADDRESS5:554//LowResolutionVideo" ></embed></OBJECT>
<OBJECT classid="clsid:9BE31822-FDAD-461B-AD51-BE1D1C159921"     codebase="http://downloads.videolan.org/pub/videolan/vlc/latest/win32/axvlc.cab"     width="240" height="180" id="vlc" events="True">   <param name="Src" value="rtsp://<USERNAME>:<PASSWORD>@IPADDRESS0:554/ch0_0.h264" />   <param name="ShowDisplay" value="True" />   <param name="AutoLoop" value="False" />   <param name="AutoPlay" value="True" />   <embed id="vlcEmb"  type="application/x-google-vlc-plugin" version="VideoLAN.VLCPlugin.2" autoplay="yes" loop="no" width="240" height="180"     target="rtsp://<USERNAME>:<PASSWORD>@IPADDRESS0:554/ch0_0.h264" ></embed></OBJECT>
<OBJECT classid="clsid:9BE31822-FDAD-461B-AD51-BE1D1C159921"     codebase="http://downloads.videolan.org/pub/videolan/vlc/latest/win32/axvlc.cab"     width="240" height="180" id="vlc" events="True">   <param name="Src" value="rtsp://<USERNAME>:<PASSWORD>@IPADDRESS1:554/ch0_0.h264" />   <param name="ShowDisplay" value="True" />   <param name="AutoLoop" value="False" />   <param name="AutoPlay" value="True" />   <embed id="vlcEmb"  type="application/x-google-vlc-plugin" version="VideoLAN.VLCPlugin.2" autoplay="yes" loop="no" width="240" height="180"     target="rtsp://<USERNAME>:<PASSWORD>@IPADDRESS1:554/ch0_0.h264" ></embed></OBJECT>
</BR>

<OBJECT classid="clsid:9BE31822-FDAD-461B-AD51-BE1D1C159921"     codebase="http://downloads.videolan.org/pub/videolan/vlc/latest/win32/axvlc.cab"     width="240" height="180" id="vlc" events="True">   <param name="Src" value="rtsp://<USERNAME>:<PASSWORD>@IPADDRESS3:554/ch0_0.h264" />   <param name="ShowDisplay" value="True" />   <param name="AutoLoop" value="False" />   <param name="AutoPlay" value="True" />   <embed id="vlcEmb"  type="application/x-google-vlc-plugin" version="VideoLAN.VLCPlugin.2" autoplay="yes" loop="no" width="240" height="180"     target="rtsp://<USERNAME>:<PASSWORD>@IPADDRESS3:554/ch0_0.h264" ></embed></OBJECT>
<OBJECT classid="clsid:9BE31822-FDAD-461B-AD51-BE1D1C159921"     codebase="http://downloads.videolan.org/pub/videolan/vlc/latest/win32/axvlc.cab"     width="240" height="180" id="vlc" events="True">   <param name="Src" value="http://<USERNAME>:<PASSWORD>@IPADDRESS/video/mjpg.cgi" />   <param name="ShowDisplay" value="True" />   <param name="AutoLoop" value="False" />   <param name="AutoPlay" value="True" />   <embed id="vlcEmb"  type="application/x-google-vlc-plugin" version="VideoLAN.VLCPlugin.2" autoplay="yes" loop="no" width="240" height="180"     target="http://<USERNAME>:<PASSWORD>@IPADDRESS/video/mjpg.cgi" ></embed></OBJECT>
</CENTER>

</BODY>
</HTML>
"@

#HTML Export
$Script | ConvertTo-HTML -Head $HTMLHeader -Body $HTMLBody  | Out-File "C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\gen\cameras.html"
#}