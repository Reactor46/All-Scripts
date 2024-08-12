IF EXIST C:\tightvncinstallsecure.txt goto end else goto install 
:install
cd c:\
del c:\tighvncinstall.txt
del c:\tighvncinstall2.txt
del c:\tighvncinstall2.txt
c:\ > tightvncinstallsecure.txt
"C:\Program Files\TightVNC\tvnserver.exe" -stop
xcopy \\usonvsvrdc\sysvol\USON.local\scripts\TightVNC\ProgramFiles\TightVNC\*.* "C:\Program Files\TightVNC\*,*" /Y
xcopy \\usonvsvrdc\sysvol\USON.local\scripts\TightVNC\StartMenu\TightVNC\*.* "C:\Documents and Settings\All Users\Start Menu\Programs\TightVNC\*.*" /Y
xcopy "\\usonvsvrdc\sysvol\USON.local\scripts\TightVNC\StartMenu\TightVNC\TightVNC Server (Application Mode)\*.*" "C:\Documents and Settings\All Users\Start Menu\Programs\TightVNC\TightVNC Server (Application Mode)\*.*" /Y 
xcopy "\\usonvsvrdc\sysvol\USON.local\scripts\TightVNC\StartMenu\TightVNC\TightVNC Server (Service Mode)\*.*" "C:\Documents and Settings\All Users\Start Menu\Programs\TightVNC\TightVNC Server (Service Mode)\*.*" /Y
"C:\Program Files\TightVNC\tvnserver.exe" -install -silent
"C:\Program Files\TightVNC\tvnserver.exe" -start
regedit /s \\usonvsvrdc\sysvol\USON.local\scripts\TightVNC\tightvncregedit.reg
ping 1.0.0.0 -n 1 -w 5000 >NUL
"C:\Program Files\TightVNC\tvnserver.exe" -stop
"C:\Program Files\TightVNC\tvnserver.exe" -start
:end




 