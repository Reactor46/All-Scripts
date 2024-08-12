@echo off

Echo Time to Shutdown:
set /p "min=Time(Min): "
set /a sec=min*60

ping -n %sec% 127.0.0.1 > NUL 2>&1 && shutdown /h /f
