' ///////////////////////////////////////////////////////////////////////////////
' // Credit One Bank Application Services Checker
' // 
' // Adapted from ActiveXperts Network Monitor (http://www.activexperts.com)
' /
' // Orginal script source:  
' // http://www.activexperts.com/admin/vbscript/network-monitor/service/
' //
' // Adaptation by vmovsessian on 9/25/2012
' // Last modified by vmovsessian on 9/25/2012
' //
'

Option Explicit
Const  retvalUnknown = 1
Dim    SYSEXPLANATION  ' Used by Network Monitor, don't change the names

' ///////////////////////////////////////////////////////////////////////////////
' // Header and start of output
Dim bResult
WScript.Echo "*********************************************"
WScript.Echo "Credit One Bank Applications Services Checker"
WScript.Echo "*********************************************" & vbCrLF
WScript.Echo "Checking services..." & vbCrLF
' ///////////////////////////////////////////////////////////////////////////////

' ///////////////////////////////////////////////////////////////////////////////
' // Please make changes within this block only
bResult =  CheckMultipleServices( "LASAUTH01.creditoneapp.biz", "", "W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASAUTH02.creditoneapp.biz", "", "W3SVC" )
WScript.Echo SYSEXPLANATION

WScript.Echo "ContosoFinCenService should only be running on LASSVC04" & vbCrLF

bResult =  CheckMultipleServices( "LASSVC03", "", "CollectionsAgentTimeService;ContosoCheckRequestService;ContosoLPSService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASSVC04", "", "CentralizedCacheService;CreditOneBatchLetterRequestService;CreditOne.LogParser.Service;ContosoQueueProcessorService;CollectionsAgentTimeService;ContosoCheckRequestService;ContosoLPSService;fdrOutGoingFileWatcher;ValidationTriggerWatcher;ContosoFinCenService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCASMT01", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCASMT02", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCASMT03", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCASMT04", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCASMT05", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCASMT06", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCASMT07", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCASMT08", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCASMT09", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCASMT10", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCOLL01", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCOLL02", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCOLL03", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCOLL04", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCOLL05", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCOLL06", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCOLL07", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCOLL08", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCOLL09", "", "ContosoDataLayerService;W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASMT01", "", "ContosoDataLayerService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASMT02", "", "ContosoDataLayerService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASMT03", "", "ContosoDataLayerService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASMT04", "", "ContosoDataLayerService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASMT05", "", "ContosoDataLayerService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASMT06", "", "ContosoDataLayerService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASMT07", "", "ContosoDataLayerService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASMT08", "", "ContosoDataLayerService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASMT09", "", "ContosoDataLayerService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASMT10", "", "ContosoDataLayerService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASMCE01", "", "CreditEngine" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASMCE02", "", "CreditEngine" )
WScript.Echo SYSEXPLANATION

WScript.Echo " " & vbCrLF
WScript.Echo "The following services should only be running on ONE of the four LASCAPSMT servers" & vbCrLF
WScript.Echo "BoardingService " & vbCr
WScript.Echo "ContosoDebitCardHolderFileWatcher " & vbCr
WScript.Echo "FromPPSExchangeFileWatcherService " & vbCr
WScript.Echo "FromPPSExchangeFileWatcherService " & vbCr
WScript.Echo "ContosoApplicationParsingService " & vbCrLF

bResult =  CheckMultipleServices( "LASCAPSMT01", "", "CreditPullService;ContosoIPFraudCheckService;ContosoIdentityCheckService;ContosoApplicationProcessingService;ContosoApplicationImportService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCAPSMT02", "", "CreditPullService;ContosoIPFraudCheckService;ContosoIdentityCheckService;ContosoApplicationProcessingService;ContosoDebitCardHolderFileWatcher;FromPPSExchangeFileWatcherService;ContosoApplicationImportService;ContosoApplicationParsingService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCAPSMT05", "", "CreditPullService;ContosoIPFraudCheckService;ContosoIdentityCheckService;ContosoApplicationProcessingService;ContosoApplicationImportService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCAPSMT06", "", "CreditPullService;ContosoIPFraudCheckService;ContosoIdentityCheckService;ContosoApplicationProcessingService;ContosoApplicationImportService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCAPS01", "", "W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCAPS02", "", "W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCAPS05", "", "W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCAPS06", "", "W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCAS01", "", "W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCAS02", "", "W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCAS03", "", "W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCAS04", "", "W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCAS05", "", "W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCAS06", "", "W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCAS07", "", "W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCAS08", "", "W3SVC" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCHAT01", "", "WhosOnGateway;WhosOnQuery;WhosOnReports;WhosOn;WhosOnServiceMonitor" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCHAT02", "", "WhosOnGateway;WhosOnQuery;WhosOnReports;WhosOn;WhosOnServiceMonitor" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASPROCDB02", "", "MSSQLSERVER;SQLSERVERAGENT" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASPROCAPP03", "", "P360PA-DOC_LINK;P360PA-WFSERVE;P360PA-WFSERVE2;P360PA-WFSERVEARCHIVE;P360PA-WFSERVEIMPORT" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASPROCAPP04", "", "Service1" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASRFAX01", "", "RFALERT;RFDB;RFDOCTRANS;RFEDCPARENTNOTIFY1;RFEDCCONV1;RFEDCFCOPY1;RFEDCGIMG1;RFEDCGENPCL1;RFEDCGTEXT1;RFEDCLOOKUP1;RFEDCMGR;RFEDCPARENTMON1;RFEDCRFSEND1;RFEDCRFSENDSTAT1;RFMIME;RFPAGE;RFQUEUE;RFRemote;RFRPC;RFSERVER;RFWORK1;RFWORK2;RFWORK3" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASCODE02", "", "AccuRev Server;AccuRev DB Server;AccuRev RLM;JIRA050414101333" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASITS01", "", "SQLANYs_sem5;SepMasterService;semsrv;semwebsrv;SmcService" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASPRINT01", "", "Spooler" )
WScript.Echo SYSEXPLANATION

bResult =  CheckMultipleServices( "LASINFRA01", "", "Schedule" )
WScript.Echo SYSEXPLANATION



' //////////////////////////////////////////////////////////////////////////////

' //////////////////////////////////////////////////////////////////////////////

Function CheckService( strComputer, strCredentials, strService )

' Description: 
'     Checks if a service, specified by strService, is running on the machine specified by strComputer. 
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
'     3) strService As String - Name of the service
' Usage:
'     CheckService( "", "", "" )
' Sample:
'     CheckService( "localhost", "", "alerter" )

    Dim objWMIService

    CheckService      = retvalUnknown  ' Default return value
    SYSEXPLANATION    = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckService      = checkServiceWMI( objWMIService, strComputer, strService, SYSEXPLANATION )

End Function

' //////////////////////////////////////////////////////////////////////////////

' //////////////////////////////////////////////////////////////////////////////

Function CheckMultipleServices( strComputerName, strCredentials, strServices )
' Description: 
'     Checks if a list of service, specified by strServices, is running on the machine specified by strComputer. 
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
'     3) strServices As String - List of services; the services are separated by the ';' character
' Usage:
'     CheckMultipleServices( "", "", "" )
' Sample:
'     CheckServices( "localhost", "", "alerter;messenger" )

    Dim strUnknownList, strErrorList, strSuccessList
    Dim arrServices, i
	Dim numSuccess, numError, numUnknown
    Dim numResult, numTotalResult

	CheckMultipleServices  = True ' True, False or retvalUnknown
    numTotalResult         = True

	numSuccess = "0"
	numError = "0"
	numUnknown = "0"
	
    arrServices            = Split( strServices, ";" )

    For i = 0 To UBound( arrServices )
       numResult = CheckService( strComputerName, strCredentials, arrServices( i ) ) 
       If( numResult = True ) Then
          numSuccess = numSuccess + 1
		  strSuccessList   = strSuccessList + vbCrLF + ">>>>>>>>>> "
          strSuccessList   = strSuccessList + arrServices( i )
          strSuccessList   = strSuccessList + " <<<<<<<<<<"
       ElseIf( numResult = False ) Then
          numError = numError + 1
		  strErrorList     = strErrorList + vbCrLF + ">>>>>>>>>> "
          strErrorList     = strErrorList + arrServices( i )
          strErrorList     = strErrorList + " <<<<<<<<<<"
      Else
          numUnknown = numUnknown + 1
		  strUnknownList   = strUnknownList + vbCrLF + ">>>>>>>>>> "
          strUnknownList   = strUnknownList + arrServices( i )
          strUnknownList   = strUnknownList + " <<<<<<<<<<"
       End If
       
       If( numResult = True ) Then
         ' Nothing to do
       ElseIf( numResult = retvalUnknown ) Then
         If( numTotalResult <> False ) Then
           numTotalResult = numResult
         Else
           numTotalResult = False
         End If
       Else ' numResult   = False
         numTotalResult   = False
       End If
    Next
       
    SYSEXPLANATION = "There are " & numSuccess & " services RUNNING on " & strComputerName & "." & vbCrLF 
	SYSEXPLANATION = SYSEXPLANATION + "The following services are NOT RUNNING on " & strComputerName & ": " & Trim( strErrorList ) & vbCrLF
	SYSEXPLANATION = SYSEXPLANATION + "The following services have status UNKNOWN on " & strComputerName & ": " & Trim( strUnknownList ) & vbCrLF
	    
    CheckMultipleServices = numTotalResult

End Function




' //////////////////////////////////////////////////////////////////////////////
' //
' // Private Functions
' //   NOTE: Private functions are used by the above functions, and will not
' //         be called directly by the ActiveXperts Network Monitor Service.
' //         Private function names start with a lower case character and will
' //         not be listed in the Network Monitor's function browser.
' //
' //////////////////////////////////////////////////////////////////////////////

Function checkServiceWMI( objWMIService, strComputer, strService, BYREF strSysExplanation )

    Dim colServices, objService

    checkServiceWMI             = retvalUnknown  ' Default return value

On Error Resume Next

    Set colServices = objWMIService.ExecQuery( "Select * from Win32_Service" )
    If( Err.Number <> 0 ) Then
        strSysExplanation  = "Unable to query WMI on computer [" & strComputer & "]"
        Exit Function
    End If
    If( colServices.Count <= 0  ) Then
        strSysExplanation  = "Win32_Service does not exist on computer [" & strComputer & "]"
        Exit Function
    End If

On Error Goto 0


    For Each objService In colServices
        If( Err.Number <> 0 ) Then
            checkServiceWMI     = retvalUnknown
            strSysExplanation   = "Unable to list services on computer [" & strComputer & "]"
            Exit Function 
        End If

        If( UCase( objService.Name ) = UCase( strService ) OR UCase( objService.DisplayName ) = UCase( strService ) ) Then
            If( objService.State <> "Running" ) Then
                checkServiceWMI = False
                strSysExplanation = "Service [" & objService.DisplayName & "] is " & objService.State
                Exit Function
            End If

            checkServiceWMI     = True
            strSysExplanation   = "Service [" & objService.DisplayName & "] is " & objService.State
            Exit Function			    
        End If
    Next

    checkServiceWMI             = retvalUnknown
    strSysExplanation           = "Service [" & strService & "] was not found on computer [" & strComputer & "]"
End Function

' //////////////////////////////////////////////////////////////////////////////

Function getWMIObject( strComputer, strCredentials, BYREF objWMIService, BYREF strSysExplanation )	

On Error Resume Next

    Dim objNMServerCredentials, objSWbemLocator, colItems
    Dim strUsername, strPassword

    getWMIObject              = False

    Set objWMIService         = Nothing
    
    If( strCredentials = "" ) Then	
        ' Connect to remote host on same domain using same security context
        Set objWMIService     = GetObject( "winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer &"\root\cimv2" )
    Else
        Set objNMServerCredentials = CreateObject( "ActiveXperts.NMServerCredentials" )

        strUsername           = objNMServerCredentials.GetLogin( strCredentials )
        strPassword           = objNMServerCredentials.GetPassword( strCredentials )

        If( strUsername = "" ) Then
            getWMIObject      = False
            strSysExplanation = "No alternate credentials defined for [" & strCredentials & "]. In the Manager application, select 'Options' from the 'Tools' menu and select the 'Server Credentials' tab to enter alternate credentials"
            Exit Function
        End If
	
        ' Connect to remote host using different security context and/or different domain 
        Set objSWbemLocator   = CreateObject( "WbemScripting.SWbemLocator" )
        Set objWMIService     = objSWbemLocator.ConnectServer( strComputer, "root\cimv2", strUsername, strPassword )

        If( Err.Number <> 0 ) Then
            objWMIService     = Nothing
            getWMIObject      = False
            strSysExplanation = "Unable to access [" & strComputer & "]. Possible reasons: WMI not running on the remote server, Windows firewall is blocking WMI calls, insufficient rights, or remote server down"
            Exit Function
        End If

        objWMIService.Security_.ImpersonationLevel = 3

    End If
	
    If( Err.Number <> 0 ) Then
        objWMIService         = Nothing
        getWMIObject          = False
        strSysExplanation     = "Unable to access '" & strComputer & "'. Possible reasons: no WMI installed on the remote server, no rights to access remote WMI service, or remote server down"
        Exit Function
    End If    

    getWMIObject              = True 

End Function