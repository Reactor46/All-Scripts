' ========================================================
'   Script:					ADReplicate.vbs
'	Description:			Replicate all Domain Controllers
'								This script replicates each domain controller in Active Directory.
'	Author:					Alan Kobb
'								Herley-CTI
'	Originally created:	7/28/2008 - 16:56:25
'  IMPORTANT:				The Domain Controller name and Domain Name must be modified below.
'                       These must be string constants, 
' ========================================================

' *************************** INSERT DOMAIN CONTROLLER NAME AND DOMAIN NAME HERE
Const DomainController = "BRANCH1DC"
Const DomainName = "ALDERGROVECU.LOCAL"
' *************************** INSERT DOMAIN CONTROLLER NAME AND DOMAIN NAME HERE

'REGION Force Script to run in CSCRIPT
'If Not WScript.FullName = WScript.Path & "\cscript.exe" Then
'CreateObject("Wscript.Shell").Run WScript.Path & "\cscript.exe //NOLOGO " & Chr(34) & WScript.scriptFullName & Chr(34),1,False
'	WScript.Quit 0
'End If
'ENDREGION

Set objIADs=CreateObject("IADsTools.DCFunctions")
'read the list of domain controllers
Result=objIADs.DsGetDCList(DomainController,DomainName,1)

If Result = -1 Then
	WScript.echo "The error returned was: " + objIADs.LastErrorText
Else
	WScript.echo "There were " & CStr(Result) & " Domain Controllers returned."
	WScript.echo "--------------------------------------------------"
	For i = 1 To Result
		'for each domain controller, get the number of Directory Partitions (non-partial) it hosts
		WScript.echo "Checking domain controller: " + objIADs.DCListEntryNetBiosName(i)
		PartitionResult=objIADs.GetNamingContexts(objIADs.DCListEntryNetBiosName(i))

		If PartitionResult=-1 Then					'if we couldn't reach the server, skip it
			WScript.echo "Could not reach the server: " + objIADs.DCListEntryNetBiosName(i)
		Else
			WScript.echo "Found " + CStr(PartitionResult) + " Directory Partitions (non-partial) on (" + objIADs.DCListEntryNetBiosName(i) + ")."
			'query the status of each directory partition
			For j = 1 To PartitionResult
				ReplResult=objIADs.GetDirectPartnersEx(objIADs.DCListEntryNetBiosName(i),objIADs.NamingContextName(j), 0)
				'See if there's a failure code other than zero for any of the replication partners
				For k = 1 To ReplResult
					If objIADs.DirectPartnerFailReason(k) > 0 Then
						WScript.echo "Failure detected replicating partition (" + objIADs.NamingContextName(j) + ") from (" + objIADs.DirectPartnerName(k) + ")."
					Else
						WScript.echo "OK --- Replicating partition (" + objIADs.NamingContextName(j) + ") from (" + objIADs.DirectPartnerName(k) + ")."

						If InStr(objIADs.NamingContextName(j),"CN=Configuration") = 0 Then
							strFrom = field(objIADs.DirectPartnerName(k),"\",2)
							strTo = objIADs.DCListEntryNetBiosName(i)
							WScript.Echo vbTab & "Forcing Replication of " & CStr(objIADs.NamingContextName(j)) & " from " & strFrom & " to " & strTo & "."
							If objIADs.ReplicaSync(CStr(strFrom),CStr(objIADs.NamingContextName(j)),CStr(strTo)) <> 0 Then
								WScript.Echo vbTab & "*** Replication failed ***"
							End If
						End If
					End If
				Next
			Next
		End If
	Next
End If

Function Field(Str,Delim,Pos)
Dim aString
aString = Split(Str,Delim)
Field = aString(Pos-1)
End Function
