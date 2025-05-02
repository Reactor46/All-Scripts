Import-Module .\Resize-Image\Resize-Image.psm1
$Source = ".\Original.jpg"
$Destination = "\\lasfs03\software\Backgrounds"


Resize-Image -InputFile $Source -Width 768 -Height 1280 -OutPutFile $Destination\background768x1280.jpg
Resize-Image -InputFile $Source -Width 900 -Height 1440 -OutPutFile $Destination\background900x1440.jpg
Resize-Image -InputFile $Source -Width 960 -Height 1280 -OutPutFile $Destination\background960x1280.jpg
Resize-Image -InputFile $Source -Width 1024 -Height 768 -OutPutFile $Destination\background1024x768.jpg
Resize-Image -InputFile $Source -Width 1024 -Height 1280 -OutPutFile $Destination\background1024x1280.jpg
Resize-Image -InputFile $Source -Width 1080 -Height 811 -OutPutFile $Destination\background1080x811.jpg
Resize-Image -InputFile $Source -Width 1280 -Height 768 -OutPutFile $Destination\background1280x768.jpg
Resize-Image -InputFile $Source -Width 1280 -Height 900 -OutPutFile $Destination\background1280x800.jpg
Resize-Image -InputFile $Source -Width 1280 -Height 960 -OutPutFile $Destination\background1280x960.jpg
Resize-Image -InputFile $Source -Width 1280 -Height 1024 -OutPutFile $Destination\background1280x1024.jpg
Resize-Image -InputFile $Source -Width 1360 -Height 768 -OutPutFile $Destination\background1360x768.jpg
Resize-Image -InputFile $Source -Width 1366 -Height 768 -OutPutFile $Destination\background1366x768.jpg
Resize-Image -InputFile $Source -Width 1368 -Height 768 -OutPutFile $Destination\background1368x768.jpg
Resize-Image -InputFile $Source -Width 1400 -Height 900 -OutPutFile $Destination\background1400x900.jpg
Resize-Image -InputFile $Source -Width 1440 -Height 900 -OutPutFile $Destination\background1440x900.jpg
Resize-Image -InputFile $Source -Width 1600 -Height 1200 -OutPutFile $Destination\background1600x1200.jpg
Resize-Image -InputFile $Source -Width 1900 -Height 1080 -OutPutFile $Destination\background1900x1080.jpg
Resize-Image -InputFile $Source -Width 1900 -Height 1200 -OutPutFile $Destination\background1900x1200.jpg
Resize-Image -InputFile $Source -Width 1920 -Height 1080 -OutPutFile $Destination\background1920x1080.jpg
Resize-Image -InputFile $Source -Width 1920 -Height 1200 -OutPutFile $Destination\background1920x1200.jpg
Resize-Image -InputFile $Source -Width 2560 -Height 1440 -OutPutFile $Destination\background2560x1440.jpg