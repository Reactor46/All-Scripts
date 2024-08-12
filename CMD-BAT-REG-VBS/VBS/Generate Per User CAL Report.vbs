SET Args = WScript.Arguments 
 
Dim DomainNameList(100)                'List of Domains in FQDN format 
Dim DCPathList(100)                'List of Domain FQDN in LDAP format 
Dim LDAPPathList(100)                'List of LDAP Path - LDAP://DCPath 
Dim W2K3UserDetailsFile(100)            'W2K3 User Details File name for each domain 
Dim W2K8UserDetailsFile(100)            'W2K8 User Details File name for each domain 
 
Const ForReading = 1, ForWriting = 2 
 
'Checks whether the command line options are passed correctly or not 
 
IsW2K3Usage = 0 
IsW2K8Usage = 0 
 
IF Args.Length = 0 THEN 
        Help 
    WSCRIPT.QUIT(1)  
ELSE 
       SELECT CASE UCASE(Args(0)) 
    CASE "/W2K3" 
        IsW2K3Usage = 1 
    CASE "/W2K8" 
        IsW2K8Usage = 1 
    CASE "/ALL" 
        IsW2K3Usage = 1 
        IsW2K8Usage = 1 
    CASE ELSE 
        Help 
        WSCRIPT.QUIT(1) 
    END SELECT 
END IF 
 
DomArgListInd = 1 
 
'In case no LDAP Path is passed, current domain is assumed 
 
IF Args.Length = DomArgListInd THEN 
 
        ON ERROR RESUME NEXT 
        SET objRootDSE = GetObject("LDAP://RootDSE") 
        IF IsNull(objRootDSE) = TRUE OR IsEmpty(objRootDSE) = TRUE THEN 
        WSCRIPT.ECHO "Current domain is not reachable"  
                WSCRIPT.QUIT(1) 
        END IF 
     
        ON ERROR RESUME NEXT 
        strConfigurationNC = objRootDSE.Get("defaultNamingContext") 
        IF IsNull(strConfigurationNC) = TRUE THEN 
        WSCRIPT.ECHO "Current domain is not reachable" 
        WSCRIPT.QUIT(1) 
        END IF 
 
        NumOfDomain = 1 
        DCPathList(0) = strConfigurationNC 
    DomainNameList(0) = Replace(Replace(DCPathList(0),"DC=",""),",",".") 
    W2K3UserDetailsFile(0) = DomainNameList(0) & ".W2K3UserDetails.csv" 
    W2K8UserDetailsFile(0) = DomainNameList(0) & ".W2K8UserDetails.csv" 
  
ELSE 
       NumOfDomain = Args.Length - DomArgListInd 
 
       For i = 0 To (NumOfDomain - 1) 
        DomainNameList(i) = Args(i + DomArgListInd) 
        DomStr = Split(DomainNameList(i), ".") 
        DCPathList(i) = "DC=" & Join(DomStr, ",DC=") 
        W2K3UserDetailsFile(i) = DomainNameList(i) & ".W2K3UserDetails.csv" 
        W2K8UserDetailsFile(i) = DomainNameList(i) & ".W2K8UserDetails.csv" 
       Next 
END IF 
 
For i = 0 To (NumOfDomain - 1) 
 
    LDAPPathList(i) = "LDAP://" & DCPathList(i) 
 
    IsDomainAccessible = 1 
 
    On Error Resume Next 
 
    Const ADS_SCOPE_SUBTREE = 2 
 
    Set objConnection = CreateObject("ADODB.Connection") 
    Set objCommand =   CreateObject("ADODB.Command") 
    objConnection.Open "Provider=ADsDSOObject;" 
    Set objCommand.ActiveConnection = objConnection 
 
    objCommand.Properties("Page Size") = 1000 
    objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
    objCommand.Properties("Chase referrals") = &H60     'ADS_CHASE_REFERRALS_EXTERNAL 
 
 
    'Tracking W2K8 PU CAL usage 
 
    IF IsW2K8Usage = 1 THEN 
 
        WSCRIPT.ECHO "Querying Domain " & DomainNameList(i) & " for W2K8 CAL Details ..." 
     
        NumOfW2K8_Valid = 0 
        NumOfW2K8_Expired = 0 
        NumOfW2K8_Total = 0 
 
        ON ERROR RESUME NEXT 
        objCommand.CommandText = "SELECT ADsPath FROM '" & LDAPPathList(i) & "' WHERE objectCategory='person' AND objectClass='user' AND ((terminalServer='*') OR (msTSManagingLS='*' AND msTSLicenseVersion='*' AND msTSExpireDate='*'))"        
        Set objRecordSet = objCommand.Execute 
        WSCRIPT.ECHO "Total no of user objects in the domain: " & objRecordset.RecordCount 
        IF objRecordset.RecordCount = 0 THEN 
               IsDomainAccessible = 0 
        END IF 
        objRecordSet.MoveFirst 
 
        IF IsDomainAccessible = 1 THEN 
 
            'Updating W2K8 User Details into file 
     
            Set objFSO = CreateObject("Scripting.FileSystemObject") 
            If objFSO.FileExists(W2K8UserDetailsFile(i)) Then 
                    Set objTextFile = objFSO.OpenTextFile(W2K8UserDetailsFile(i), ForWriting) 
            Else 
                Set objTextFile = objFSO.CreateTextFile(W2K8UserDetailsFile(i)) 
            End If 
 
            OutputFileLine = "User Name,CAL Status" 
            objTextFile.WriteLine OutputFileLine 
             
            Do Until objRecordSet.EOF 
                IsValidPUCAL = 0 
 
                ON ERROR RESUME NEXT 
                    strPath = objRecordSet.Fields("AdsPath").Value 
                   Set objUser = GetObject(strPath) 
 
                    ON ERROR RESUME NEXT 
                    Value1 = objUser.Get("terminalServer") 
                    Value1Err = ERR.number 
 
                    ON ERROR RESUME NEXT 
                    Value2 = objUser.Get("msTSManagingLS") 
                    Value2Err = ERR.number 
 
                    ON ERROR RESUME NEXT 
                    Value3 = objUser.Get("msTSLicenseVersion") 
                    Value3Err = ERR.number 
 
                    ON ERROR RESUME NEXT 
                    Value4 = objUser.Get("msTSExpireDate") 
                    Value4Err = ERR.number 
                 
                    IF Value1Err = 0 AND Value2Err <> 0 AND Value3Err <> 0 AND Value4Err <> 0 THEN 
     
                    ' This means User Account is of Win2K3 or Older Domain Controller 
                    IsValidPUCAL = 1 
     
                    ELSEIF Value1Err <> 0 AND Value2Err = 0 AND Value3Err = 0 AND Value4Err = 0 THEN 
 
                    ' This means User Account is of w2k8 or newer Domain Controller 
                    IF Value4 < now() THEN 
                        IsValidPUCAL = 2 
                    ELSE 
                        IsValidPUCAL = 1 
                    END IF 
 
                    ELSE 
                    ' This means User does not have License" 
                    IsValidPUCAL = 0 
                    END IF 
                 
                IF IsValidPUCAL = 2 THEN 
                    OutputFileLine = objUser.sAMAccountName & ",EXPIRED" 
                    objTextFile.WriteLine OutputFileLine 
                    NumOfW2K8_Expired = NumOfW2K8_Expired + 1                 
                ELSEIF IsValidPUCAL = 1 THEN 
                    OutputFileLine = objUser.sAMAccountName & ",VALID" 
                    objTextFile.WriteLine OutputFileLine 
                    NumOfW2K8_Valid = NumOfW2K8_Valid + 1                 
                END IF 
 
                objRecordSet.MoveNext 
            Loop 
 
            objTextFile.Close 
            Set objTextFile = Nothing 
            Set objFSO = Nothing 
 
            Set objRecordSet = Nothing 
         
            NumOfW2K8_Total = NumOfW2K8_Valid + NumOfW2K8_Expired 
 
            WSCRIPT.ECHO "Done!" 
            WSCRIPT.ECHO "For Domain: " & DomainNameList(i) & "  --  Number of W2K8 CALs - Valid: " & NumOfW2K8_Valid & "  Expired: " & NumOfW2K8_Expired & "  Total: " & NumOfW2K8_Total 
            WSCRIPT.ECHO "For W2K8 User details of domain: " & DomainNameList(i) & " please refer to the file " & W2K8UserDetailsFile(i) & " saved in the current directory." 
            WSCRIPT.ECHO "" 
 
        ELSE 
            WSCRIPT.ECHO "Domain " & DomainNameList(i) & " is not reachable. W2K8 CAL details cannot be retrieved." 
            WSCRIPT.ECHO "" 
 
        END IF             
         
    END IF 
 
 
    'Tracking W2K3 PU User Connections 
    IF IsW2K3Usage = 1 THEN 
 
        WSCRIPT.ECHO "Querying Domain " & DomainNameList(i) & " for W2K3 User Details ..." 
 
        NumOfW2K3_Valid = 0 
        NumOfW2K3_Expired = 0 
        NumOfW2K3_Total = 0 
        IsMaxLimitReached = 0         
 
        'Read existing User Details from file 
 
        Dim W2K3UserName(10000) 
        Dim W2K3UserDomain(10000) 
        Dim W2K3UserLogon(10000) 
        Dim W2K3UserValidity(10000) 
        NumOfW2K3User = 0 
        rec_cnt = 0 
        NumOfLine = 0 
 
        Set objFSO = CreateObject("Scripting.FileSystemObject") 
 
        If objFSO.FileExists(W2K3UserDetailsFile(i)) Then 
 
                Set objTextFile = objFSO.OpenTextFile(W2K3UserDetailsFile(i), ForReading) 
            Do While objTextFile.AtEndOfStream <> True 
 
                  strLine = objtextFile.ReadLine 
                  If (NumOfLine > 0) AND (inStr(strLine, ",")) Then 
                        UserDetailsRecord = split(strLine, ",") 
                    W2K3UserName(rec_cnt) = UserDetailsRecord(0) 
                        W2K3UserDomain(rec_cnt) = UserDetailsRecord(1) 
                    W2K3UserLogon(rec_cnt) = UserDetailsRecord(2) 
 
                        rec_cnt = rec_cnt + 1 
                Else 
                    NumOfLine = NumOfLine + 1 
                  End If 
 
            Loop 
 
            NumOfW2K3User = rec_cnt 
            objTextFile.Close 
 
            Set objTextFile = Nothing 
            Set objFSO = Nothing 
        End If     
     
        'Finding the W2K3 TS in that domain 
 
        IsW2K3TSAvailable = 1 
        ON ERROR RESUME NEXT 
 
        objCommand.CommandText = "SELECT Name, operatingSystemVersion FROM '" & LDAPPathList(i) & "' WHERE objectClass='computer' AND operatingSystemVersion='5.2 (3790)'"        
        Set objRecordSet = objCommand.Execute 
        IF objRecordset.RecordCount = 0 THEN 
            IsW2K3TSAvailable = 0 
        END IF 
        objRecordSet.MoveFirst 
 
        IF IsW2K3TSAvailable = 1 THEN 
 
            Do Until objRecordSet.EOF 
 
                'Checking which of the TS are in PU Mode 
                    NameSpace = "\root\cimv2" 
                TSName = objRecordSet.Fields("Name") 
                IsW2K3PUTS = 0 
 
                    ON ERROR RESUME NEXT             
                    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & TSName & NameSpace ) 
                Set ObjArray = ObjWMIService.ExecQuery ("select * from Win32_TerminalServiceSetting") 
 
                FOR EACH Obj IN ObjArray 
                    IF Obj.LicensingType = 4 THEN 
                        IsW2K3PUTS = 1 
                    END IF 
                    EXIT FOR 
                NEXT 
 
                'Populating the list of all users remotely logged in to a W2K3 PU mode TS 
                IF IsW2K3PUTS = 1 THEN 
                    ON ERROR RESUME NEXT                 
                    Set colSessions = objWMIService.ExecQuery("Select * from Win32_LogonSession Where LogonType = 10") 
 
                    If colSessions.Count > 0 Then  
                           For Each objSession in colSessions 
 
                            ON ERROR RESUME NEXT 
 
                            For Each loggeduser in objWMIService.ExecQuery("select * from Win32_LoggedOnUser where dependent = """ & replace(objSession.path_.relpath, """", "\""") & """") 
                                                 
                                UserAccountString = loggeduser.antecedent 
                                UserAccountArray = Split(UserAccountString, """", -1)                                 
                                UserName = UserAccountArray(3) 
                                UserDomain = UserAccountArray(1)                             
                             
                                'If this user is already present in the user details file then edit the user details if needed 
                                IsUserAlreadyInList = 0 
                                For j = 0 To (NumOfW2K3User - 1) 
                                    If (LCase(W2K3UserName(j)) = LCase(UserName)) AND (LCase(W2K3UserDomain(j)) = LCase(UserDomain)) Then 
                                        IsUserAlreadyInList = 1                                     
                                        IF W2K3UserLogon(j) < DateAdd("d", -60, now())  THEN 
                                            W2K3UserLogon(j) = now() 
                                            W2K3UserValidity(j) = "ACTIVE"                                     
                                        END IF 
                                        EXIT FOR 
                                    End If 
                                Next 
                             
                                'If not then add user details in the file 
                                If IsUserAlreadyInList = 0 Then 
                                    If IsUserAlreadyInList >= 10000 Then 
                                        WSCRIPT.ECHO "Number of W2K3 Users reached the max limit of 10000" 
                                        IsMaxLimitReached = 1 
                                        EXIT FOR 
                                    End If 
                                    W2K3UserName(NumOfW2K3User) = UserName 
                                    W2K3UserDomain(NumOfW2K3User) = UserDomain 
                                    W2K3UserLogon(NumOfW2K3User) = now() 
                                    W2K3UserValidity(NumOfW2K3User) = "ACTIVE" 
                                    NumOfW2K3User = NumOfW2K3User + 1 
                                End If 
                                 Next                 
                            Set colList = Nothing 
                            If IsMaxLimitReached = 1 Then                                 
                                EXIT FOR 
                            End If                     
                           Next 
                    End If 
 
                    Set colSessions = Nothing 
                End If 
 
                Set objWMIService = Nothing 
                Set ObjArray = Nothing 
 
                objRecordSet.MoveNext 
 
                If IsMaxLimitReached = 1 Then                                 
                    EXIT Do 
                End If 
            Loop 
 
            ' W2K3 User Details per domain 
 
            For rec_cnt = 0 To (NumOfW2K3User - 1) 
                IF W2K3UserLogon(rec_cnt) >= DateAdd("d", -60, now())  THEN             
                    W2K3UserValidity(rec_cnt) = "ACTIVE" 
                    NumOfW2K3_Valid = NumOfW2K3_Valid + 1                 
                ELSE 
                    W2K3UserValidity(rec_cnt) = "STALE" 
                    NumOfW2K3_Expired = NumOfW2K3_Expired + 1                 
                END IF                 
            Next 
 
            'Updating W2K3 User Details into file 
     
            Set objFSO = CreateObject("Scripting.FileSystemObject") 
            If objFSO.FileExists(W2K3UserDetailsFile(i)) Then 
                    Set objTextFile = objFSO.OpenTextFile(W2K3UserDetailsFile(i), ForWriting) 
            Else 
                Set objTextFile = objFSO.CreateTextFile(W2K3UserDetailsFile(i)) 
            End If 
 
            OutputFileLine = "User Name,User Domain,Connection Time,Connection Status" 
            objTextFile.WriteLine OutputFileLine 
 
            For rec_cnt = 0 To (NumOfW2K3User - 1) 
                OutputFileLine = W2K3UserName(rec_cnt) & "," & W2K3UserDomain(rec_cnt) & "," & W2K3UserLogon(rec_cnt) & "," & W2K3UserValidity(rec_cnt) 
                objTextFile.WriteLine OutputFileLine 
            Next 
 
            objTextFile.Close 
 
            Set objTextFile = Nothing 
            Set objFSO = Nothing 
 
            Set objRecordSet = Nothing 
         
            NumOfW2K3_Total = NumOfW2K3_Valid + NumOfW2K3_Expired 
 
            WSCRIPT.ECHO "Done!" 
            WSCRIPT.ECHO "For Domain: " & DomainNameList(i) & "  --  Number of W2K3 Users - Active: " & NumOfW2K3_Valid & "  Stale: " & NumOfW2K3_Expired & "  Total: " & NumOfW2K3_Total 
            WSCRIPT.ECHO "For W2K3 User details of domain: " & DomainNameList(i) & " please refer to the file " & W2K3UserDetailsFile(i) & " saved in the current directory." 
            WSCRIPT.ECHO "" 
 
        ELSE 
            WSCRIPT.ECHO "Domain " & DomainNameList(i) & " is either not reachable or no W2K3 TS is available in this domain. W2K3 user details cannot be retrieved." 
            WSCRIPT.ECHO ""     
         
        END IF 
     
                 
    End If 
 
    IF i < NumOfDomain - 1 THEN 
        WSCRIPT.ECHO "----------------------------------------------------------------------------------------------------------------------------------------------------------" 
        WSCRIPT.ECHO "" 
    END IF 
 
Next 
 
WSCRIPT.QUIT(0) 
 
 
SUB Help() 
    HelpMesg = "Usage : cscript PerUserCALReport.vbs <option> [DomainFQDN1] [DomainFQDN2] [DomainFQDN3] ..." & vbNewLine & _ 
       "" & vbNewLine & _ 
       "where option can be either of the following - " & vbCrLf & _ 
       "    /W2K3    - to get the number of users connected to W2K3 TS in PU licensing mode for a given domain(s)." & vbNewLine & _ 
       "    /W2K8      - for counting valid & expired W2K8 PU CAL usage for a given domain(s)." & vbNewLine & _ 
       "    /All     - for combined details of both the above." & vbNewLine & _ 
    "" & vbNewLine & _ 
    "DomainFQDN needs to be in the format of contoso.corp.com. If no parameter is specified, it assumes current domain." 
    WSCRIPT.ECHO "" & HelpMesg 
END SUB