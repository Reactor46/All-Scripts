@ECHO OFF
md C:\Patches\Sym
copy "\\phxfs01\Shared\IT\Symantec\*.*" C:\Patches\Sym\ /Y
Start /wait C:\Patches\Sym\SylinkDrop.exe
exit