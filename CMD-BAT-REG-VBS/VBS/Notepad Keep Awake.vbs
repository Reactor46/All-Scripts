Set objShell = WScript.CreateObject("WScript.Shell") 
ctr = 0 
objShell.Run "Notepad.exe" 
Do Until Success = True 
Success = objShell.AppActivate("Notepad") 
Wscript.Sleep 1000 
Loop 
objShell.SendKeys ctr 
do while ctr < 100 
Wscript.sleep 420000 
objShell.AppActivate("Notepad") 
objShell.SendKeys ctr 
ctr = ctr + 1 
loop 