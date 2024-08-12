@Echo on
for /f "tokens=*" %%A in (C:\LazyWinAdmin\VDI-COMPS.txt) do (
	Echo Computer: %%A
	psexec \\%%A -s -c -f C:\LazyWinAdmin\Reset-WinUpdate.cmd >> C:\LazyWinAdmin\WSUS-Reports\WSUS-Reset\%%A.log
)