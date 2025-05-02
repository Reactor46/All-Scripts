'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
' Created by Assaf Miron
' Http://assaf.miron.googlepages.com
' Date : 02/11/2009
' Sending Packet Using Winsock.vbs
' Description : Sending TCP or UDPPackets to a Computer Using WINSock.
' Using this Script You can send a customized Packet to a computer or IP Address
' In a certain Protocol (TCP or UDP) and on any Port
'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=

Function GetPackectConnectionState(intState)
	' Returns the State of the Connection
	Dim strReturn
	Select Case intState
		Case 0 		strReturn = "Connection Closed"
		Case 1 		strReturn = "Connection Open"
		Case 2 		strReturn = "Connection Listening for incoming Connections"
		Case 3 		strReturn = "Connection Pending"
		Case 4 		strReturn = "Resolving Remote Host Name"
		Case 5 		strReturn = "Remote Host Resolved"
		Case 6 		strReturn = "Connecting to Remote Host"
		Case 7 		strReturn = "Connected to Remote Host"
		Case 8 		strReturn = "Connection is Closing"
		Case 9 		strReturn = "Error Occured"
	End Select
	' Return the Connection State
	GetPackectConnectionState = strReturn
End Function

Function SendDataPacket(strRemoteHost, iProtocol, iRemotePort, strData)
	' Protocol Constants
	Const sckTCPProtocol = 0 			' Transmit Control Protocol
	Const sckUDPProtocol = 1 			' User Datagram Protocol
	' Connection Constants
	Const sckClosed = 0 				' Connection Closed
	Const sckOpen = 1 					' Connection Open
	Const sckListening = 2 				' Connection Listening for incoming Connections
	Const sckConnectionPending = 3 		' Connection Pending
	Const sckResolvingHost = 4 			' Resolving Remote Host Name
	Const sckHostResolved = 5 			' Remote Host Resolved
	Const sckConnecting = 6 			' Connecting to Remote Host
	Const sckConnected = 7 				' Connected to Remote Host
	Const sckClosing = 8 				' Connection is Closing
	Const sckError = 9 					' Error Occured
	
	Dim winSock
	Dim bReturn : bReturn = False
	' Create the WinSock Object
	Set winSock = CreateObject("MSWinSock.WinSock")
	' Set the Protocol accourding to the Protocol Enum
	Select Case UCase(iProtocol)
		Case "TCP"			iProtocol = sckTCPProtocol
		Case "UDP"			iProtocol = sckUDPProtocol
		Case Else 			iProtocol = sckTCPProtocol ' Set Default Protocol
	End Select
	' Set the Parameters of the WinSock
	With winSock
		.Protocol = iProtocol ' Set the Protocol
		.RemoteHost = strRemoteHost ' Set the Remote Host
		.RemotePort = iRemotePort ' Set the Remote Port

		If iProtocol = sckTCPProtocol Then ' Connect to Remote Host if TCP is Used
			.Connect ' Connect To Remote Host
			Do 
				WScript.Sleep 200 ' Wait until State is Updated
				If .State = 9 Then
					' Error Occoured
					Exit Function
				End If
			Loop Until .state = sckConnected
		End If
		
		.SendData strData ' Set the Data to be transmitted
	End With

	' Get the Current State of the Connection
	' 1 - Connection is Open and Data was Transmited
	' Else - Connection failed
	bReturn = winSock.State 
	' Close the Connection
	winSock.Close
	' Clean up
	Set winSock = Nothing
	' Return the Result
	SendDataPacket = bReturn
End Function

' Usage Samples
' -------------

' Send UDP Packet on remote port 1505 (FunkProxy) with Message "WAKE UP"
intReturn = SendDataPacket("MyComputer","udp",1505,"WAKE UP")
WScript.Echo GetPackectConnectionState(intReturn) ' Echo Connection Status

' Send TCP Packet on remote port 135 (FunkProxy) with Sysyem Message "<#19>Hello"
intReturn = SendDataPacket("MyComputer","tcp",139,"<119>Hello")
WScript.Echo GetPackectConnectionState(intReturn) ' Echo Connection Status
