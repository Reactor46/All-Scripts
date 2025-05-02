Const HKEY_LOCAL_MACHINE = &H80000002

Public Function sGetXPCDKey(strComputer)

On Error Resume Next

    Dim bDigitalProductID
    Dim bProductKey()
    Dim bKeyChars(24)
    Dim ilByte
    Dim nCur
    Dim sCDKey
    Dim ilKeyByte
    Dim ilBit
       
    ReDim Preserve bProductKey(14)
    
    Set objShell = CreateObject("WScript.Shell")

    Set oReg=GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
    If(Err = 0) Then
	'====================================================================================
	' Connect to computer Registry and retreve the Parent value
	'====================================================================================
		strKeyPath = "SOFTWARE\MICROSOFT\Windows NT\CurrentVersion"
		strValueName = "DigitalProductId"
		oReg.GetBinaryValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
	    bDigitalProductID = strValue
	    
	    Set objShell = Nothing
	
	    For ilByte = 52 To 66
	      bProductKey(ilByte - 52) = bDigitalProductID(ilByte)
	    Next
	  
	    'Possible characters in the CD Key:
	    bKeyChars(0) = Asc("B")
	    bKeyChars(1) = Asc("C")
	    bKeyChars(2) = Asc("D")
	    bKeyChars(3) = Asc("F")
	    bKeyChars(4) = Asc("G")
	    bKeyChars(5) = Asc("H")
	    bKeyChars(6) = Asc("J")
	    bKeyChars(7) = Asc("K")
	    bKeyChars(8) = Asc("M")
	    bKeyChars(9) = Asc("P")
	    bKeyChars(10) = Asc("Q")
	    bKeyChars(11) = Asc("R")
	    bKeyChars(12) = Asc("T")
	    bKeyChars(13) = Asc("V")
	    bKeyChars(14) = Asc("W")
	    bKeyChars(15) = Asc("X")
	    bKeyChars(16) = Asc("Y")
	    bKeyChars(17) = Asc("2")
	    bKeyChars(18) = Asc("3")
	    bKeyChars(19) = Asc("4")
	    bKeyChars(20) = Asc("6")
	    bKeyChars(21) = Asc("7")
	    bKeyChars(22) = Asc("8")
	    bKeyChars(23) = Asc("9")
	
	    For ilByte = 24 To 0 Step -1
	      
	      nCur = 0
	
	      For ilKeyByte = 14 To 0 Step -1
	        'Step through each byte in the Product Key
	        nCur = nCur * 256 Xor bProductKey(ilKeyByte)
	        bProductKey(ilKeyByte) = Int(nCur / 24)
	        nCur = nCur Mod 24
	      Next
	      
	      sCDKey = Chr(bKeyChars(nCur)) & sCDKey
	      If ilByte Mod 5 = 0 And ilByte <> 0 Then sCDKey = "-" & sCDKey
	    Next
	   
	    sGetXPCDKey = sCDKey
	Else
		If(Err.Number = 70) Then
			sGetXPCDKey = "Access is Denied"
		ElseIf(Err.Number = 462) Then
			sGetXPCDKey = "Unavialable"
		ElseIf(Err.Number = 429) Then
			sGetXPCDKey = "ActiveX component Error"
		Else
			sGetXPCDKey = "Unknown(" & Err.Number & ")"
		End If
	End If      
    
End Function

Wscript.Echo sGetXPCDKey(WScript.Arguments.Item(0))
