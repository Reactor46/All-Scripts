' Create Tasks to control WiFi Radio
' Disable/Enable WiFi when wired cable is plugged in/out
'##############################
'By: J.P. Klompmaker
'Email: badcluster@hotmail.com

strComputer = "." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colLAN = objWMIService.ExecQuery("Select * From Win32_NetworkAdapter Where NetConnectionID like 'Local Area Connection' and PhysicalAdapter='True'" )
Set colWiFi=objWMIService.ExecQuery ("Select * From Win32_NetworkAdapter Where Not NetConnectionID like 'Local Area Connection' and not name like 'Bluetooth%' and PhysicalAdapter='True' ")


Function disablewifi(NicService,NicIndex)
	' A constant that specifies an event trigger.
	const TriggerTypeEvent = 0

	' Create the TaskService object.
	Set service = CreateObject("Schedule.Service")
	call service.Connect()

	' Get a folder to create a task definition in. 
	Dim rootFolder
	Set rootFolder = service.GetFolder("\")

	' The taskDefinition variable is the TaskDefinition object.
	Dim taskDefinition
	' The flags parameter is 0 because it is not supported.
	Set taskDefinition = service.NewTask(0) 

	' Define information about the task.

	' Set the registration info for the task by 
	' creating the RegistrationInfo object.
	Dim regInfo
	Set regInfo = taskDefinition.RegistrationInfo
	regInfo.Description = "Turns on WIFI if wired is disconnected."
	regInfo.Author = "TMF"
	regInfo.Source = "Network Security TEAM"
	regInfo.Version = "1.0"

	' Set the task setting info for the Task Scheduler by
	' creating a TaskSettings object.
	Dim settings
	Set settings = taskDefinition.Settings
	settings.StartWhenAvailable = True
	settings.Hidden = True

	' Create an event trigger.
	Dim triggers
	Set triggers = taskDefinition.Triggers

	Dim trigger
	Set trigger = triggers.Create(TriggerTypeEvent)

	trigger.ExecutionTimeLimit = "PT0S"    'Five minutes
	trigger.Id = "EventTriggerId"

	trigger.Subscription = "<QueryList> " & _
    "<Query Id='0'> " & _
        "<Select Path='System'>*[System[Provider[@Name='" & NicService & "'] and EventID=27]]</Select>" & _
    "</Query></QueryList>"   

	Dim Action
	Set Action = taskDefinition.Actions.Create( ActionTypeExec )
	'Action.Path = "wmic path win32_networkadapter where index=" & NicIndex & " call enable"
	Action.Path = "netsh interface set interface "Wireless Network Connection" enable"
	
	WScript.Echo "Task definition created. About to submit the task..."

	call rootFolder.RegisterTaskDefinition( "Enable WIFI", taskDefinition, 6, , , 3)

	WScript.Echo "Enable WIFI Task submitted."
End Function


Function enablewifi(NicService, NicIndex)
	' A constant that specifies an event trigger.
	const TriggerTypeEvent = 0

	' Create the TaskService object.
	Set service = CreateObject("Schedule.Service")
	call service.Connect()

	' Get a folder to create a task definition in. 
	Dim rootFolder
	Set rootFolder = service.GetFolder("\")

	' The taskDefinition variable is the TaskDefinition object.
	Dim taskDefinition
	' The flags parameter is 0 because it is not supported.
	Set taskDefinition = service.NewTask(0) 

	Dim regInfo
	Set regInfo = taskDefinition.RegistrationInfo
	regInfo.Description = "Turns on WIFI if wired is disconnected."
	regInfo.Author = "TMF"
	regInfo.Source = "Network Security TEAM"
	regInfo.Version = "1.0"

	Set principals = taskDefinition.Principal
	principals.Id = "Local"
	principals.DisplayName = "Principal Discription"
	principals.UserID = "S-1-5-18"
	principals.RunLevel = "HighestAvailible"

	' Set the task setting info for the Task Scheduler by
	' creating a TaskSettings object.
	Dim settings
	Set settings = taskDefinition.Settings
	settings.StartWhenAvailable = True
	settings.hidden = True

	' Create an event trigger.
	Dim triggers
	Set triggers = taskDefinition.Triggers

	Dim trigger
	Set trigger = triggers.Create(TriggerTypeEvent)

	trigger.ExecutionTimeLimit = "PT0S"    'Five minutes
	trigger.Id = "EventTriggerId"

	trigger.Subscription = "<QueryList> " & _
    "<Query Id='0'> " & _
        "<Select Path='System'>*[System[Provider[@Name='" & NicService & "'] and EventID=32]]</Select>" & _
    "</Query></QueryList>"   

	Dim Action
	Set Action = taskDefinition.Actions.Create( ActionTypeExec )
	'Action.Path = "wmic path win32_networkadapter where index=" & NicIndex & " call disable"
	Action.Path = "netsh interface set interface "Wireless Network Connection" disable"

	WScript.Echo "Task definition created. About to submit the task..."

	call rootFolder.RegisterTaskDefinition( "Disable WIFI", taskDefinition, 6, , , 3)

	WScript.Echo "Disable WIFI Task submitted."
End Function

For Each objLAN In colLAN
	wscript.echo objLAN.name
	NicService = objLAN.servicename
Next

For Each objWifi In colWiFi
	wscript.echo objWifi.name
	NicIndex = objWifi.index 'Only used when using WMI
Next

call disablewifi(NicService,NicIndex)
call enablewifi(NicService,NicIndex)