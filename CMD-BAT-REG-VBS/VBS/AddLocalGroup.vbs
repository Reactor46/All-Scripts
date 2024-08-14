Option Explicit
On Error Resume Next

'Define Variables
Dim Mydomain
Dim GlobalGroup
Dim oDomainGroup
Dim oLocalAdmGroup
Dim oNet
Dim sComputer

Set oNet = WScript.CreateObject("WScript.Network")
sComputer = oNet.ComputerName


MyDomain = "uson.local"
GlobalGroup = "domain users"

Set oDomainGroup = GetObject("WinNT://" & MyDomain & "/" & GlobalGroup & ",group")
Set oLocalAdmGroup = GetObject("WinNT://" & sComputer & "/Administrators,group")
 
oLocalAdmGroup.Add(oDomainGroup.AdsPath)


'Nullify Variables
Set Mydomain = Nothing
Set GlobalGroup = Nothing
Set oDomainGroup = Nothing
Set oLocalAdmGroup = Nothing
Set oNet = Nothing
Set sComputer = Nothing 