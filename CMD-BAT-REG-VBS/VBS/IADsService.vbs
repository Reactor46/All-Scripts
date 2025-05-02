'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
' Created by Assaf Miron
' Http://assaf.miron.googlepages.com
' Date : 30/08/2010
' IADsService.vbs
' Description : This Script Controles a Service using ADSI, IADsService Control.
' This Script Connects to a Computer using COM, and Controlles the Service Status.
'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
Const ADS_SERVICE_STOPPED = 1
Const ADS_SERVICE_START_PENDING = 2
Const ADS_SERVICE_STOP_PENDING = 3
Const ADS_SERVICE_RUNNING = 4
Const ADS_SERVICE_CONTINUE_PENDING = 5
Const ADS_SERVICE_PAUSE_PENDING = 6
Const ADS_SERVICE_PAUSED = 7
Const ADS_SERVICE_ERROR = 8

Dim cp 'As IADsComputer
Dim sr 'As IADsService
Dim so 'As IADsServiceOperations
Dim strComputer, strService

strComputer = "." ' Computer Name
strService = "PolicyAgent" ' IPSec Service
' Connect to the Computer
Set cp = GetObject("WinNT://" & strComputer & ",computer")
' Connect to the Service Object
Set sr = cp.GetObject("Service", strService)
' Set the Object to be a ServiceOperation Object
Set so = sr
' Set the strSvcState to the Current Service State
strSvcState = GetServiceState(so.Status)
' Echo the Service State
WScript.Echo strSvcState
' Check if the Service is Stopped
If strSvcState = "Stopped" Then
	' Start the Service
	so.start
	' Set the strSvcState to the Current Service State
	strSvcState = GetServiceState(so.Status)
	' Echo the Service State
	WScript.Echo strSvcState
End If

Function GetServiceState(intState)
	' Return the Current Service State
	Select Case intState
	    Case ADS_SERVICE_STOPPED
	      GetServiceState = "Stopped"
	    Case ADS_SERVICE_RUNNING
	        GetServiceState = "Running"
	    Case ADS_SERVICE_PAUSED
	        GetServiceState = "Paused"
	    Case ADS_SERVICE_ERROR
	        GetServiceState = "Errors"
	End Select	
End Function