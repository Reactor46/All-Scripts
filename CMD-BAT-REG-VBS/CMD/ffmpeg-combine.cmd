for %%a in ("*.*") do G:\NAS\XXX\ffmpeg.exe -i "%%a" -c:v libx264 -preset slow -crf 20 -c:a aac -b:a 128k "\%%~na.mp4"
pause