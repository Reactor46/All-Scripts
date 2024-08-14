@Echo Off

ping -n 20 localhost

if exist "C:\Program Files (x86)\Microsoft Office\Office15\Lync.exe" (
	Start "Title" "C:\Program Files (x86)\Microsoft Office\Office15\Lync.exe"
) else ( 
	Start "Title" "C:\Program Files (x86)\Microsoft Lync\communicator.exe"
)
