set installer = createobject("WindowsInstaller.installer")
set wshell = createobject("wscript.shell")

package1 = "{A7174E8F-DC74-467C-94AC-C07CD9188E36}"

dim Packages(0)
Packages(0) = package1

for each x in Packages
ProductState = installer.ProductState(x)
'msgbox(x & " " & ProductState)

if ProductState = 5 or ProductState = 1 then
uninstallPrg x
end if
next

function uninstallPrg(pcode)
wshell.run "msiexec.exe /x " & pcode & " /qb", 1, true
end function
