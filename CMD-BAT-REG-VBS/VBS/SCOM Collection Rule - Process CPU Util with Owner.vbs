Option Explicit
Dim oAPI, oBag, oArgs
Dim sProcessName
Dim oProvider
Dim sQuery
Dim cPerfItems, oPerfItem
Dim cProcess, oProcess
Dim sUser, sDomain

Set oAPI = CreateObject("MOM.ScriptAPI")
Set oArgs = WScript.Arguments
sProcessName = oArgs(0)

Err.Clear
Set oProvider = GetObject("winmgmts:{(debug)}:/root/cimv2")
sQuery = "select * from Win32_PerfFormattedData_PerfProc_Process where Name Like'" + sProcessName + "%'"
Set cPerfItems = oProvider.ExecQuery(sQuery)
For Each oPerfItem in cPerfItems
	Set oBag = oAPI.CreateTypedPropertyBag(2)
	
	Set cProcess = oProvider.ExecQuery("select * from Win32_Process where ProcessID = '" & oPerfItem.IDProcess & "'")
	For Each oProcess in cProcess
		oProcess.GetOwner sUser, sDomain
	Next
	
	Call oBag.AddValue("PerfValue", oPerfItem.PercentProcessorTime)
	Call oBag.AddValue("Owner", sUser)
	oAPI.AddItem oBag
Next

Call oAPI.ReturnItems
	