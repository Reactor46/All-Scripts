@echo off
for /f "tokens=1-7 delims=,: " %%a in ('query user ^| find /i "disc"') do if %%d GTR 2 (logoff %%b) else (if %%e GTR 2 (logoff %%b))
exit


