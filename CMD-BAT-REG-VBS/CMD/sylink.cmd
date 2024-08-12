@ECHO OFF
md C:\Patches\Sym
copy "\\lasfs03\Software\Current Versions\symantec\Move to LASITS02\*.*" C:\Patches\Sym\ /Y
Start /wait C:\Patches\Sym\SylinkDrop.exe
exit