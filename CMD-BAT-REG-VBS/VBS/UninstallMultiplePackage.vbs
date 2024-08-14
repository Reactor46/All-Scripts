set installer = createobject("WindowsInstaller.installer")
set wshell = createobject("wscript.shell")

package1 = "{69DB1F97-FA44-4FB0-8F49-7B7822CACD1F}"
package2 = "{7A7F7C49-1B78-49E0-BBDF-9CB1D9A8924E}"
package3 = "{6CF0B501-7A0D-4B81-82EB-A14F11E9685F}"



dim Packages(2)
Packages(0) = package1
Packages(1) = package2
Packages(2) = package3

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
