@ECHO OFF

FOR /F "tokens=*" %%G IN ('DIR /B /S *.mov') DO F:\ffmpeg.exe -i "%%G" -map 0 -format mp4 -vcodec copy -acodec copy %%~dG%%~pnG.mp4
FOR /F "tokens=*" %%G IN ('DIR /B /S *.AVI') DO F:\ffmpeg.exe -i "%%G" -map 0 -format mp4 -vcodec copy -acodec copy %%~dG%%~pnG.mp4
rem for %%a in ("*.mov") do ffmpeg -i ^"%%a^" -map 0 -format mp4 -vcodec copy -acodec copy  ^"mp4s/%%~na.mp4^"