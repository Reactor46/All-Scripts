'=*=*=*=*=*=*=*=*=*=*=*=*=
' Author : Assaf Miron 
' Http://assaf.miron.googlepages.com
' Date : 06/10/2009
' Cluster Resource Monitor.vbs
' Description : This Script Monitors a Cluster and Echos when an Event Pops up
' The Script Echos the Cluster Notifications and the Following Resource Events: 
' Cluster Up/Down Events
' Cluster Group Online/Offline
' Cluster Resource Online/Offline/Failed
' Cluster Network Resource Up/Down
'=*=*=*=*=*=*=*=*=*=*=*=*=
Option Explicit
'On Error Resume Next


Dim objClusterEvent, objClusterEvents, objClusterService
Dim objSWbemLocator
Dim objClusterNotifyDictionary
Dim strQuery, strComputer
Dim prop

' Cluster Change Resource State = 256
strQuery = "Select * from MSCluster_Event"
strComputer = "VM2K3NLBN1"
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objClusterService = objSWbemLocator.ConnectServer(strComputer,"Root\MSCluster")
Set objClusterEvents = objClusterService.ExecNotificationQuery(strQuery)

Do 
	Set objClusterEvent = objClusterEvents.NextEvent(300000)
	If Err = 0 Then
		WScript.Echo vbNewLine & "Class: " & objClusterEvent.Path_.Class
		WScript.Echo "EventObjectName: " & objClusterEvent.EventObjectName
		WScript.Echo "EventObjectPath: " & objClusterEvent.EventObjectPath
		WScript.Echo "EventObjectType: " & objClusterEvent.EventObjectType
		WScript.Echo "EventTypeMajor: " & objClusterEvent.EventTypeMajor 
		WScript.Echo "EventTypeMinor: " & objClusterEvent.EventTypeMinor
		WScript.Echo "ClusterNotification: " & GetClusterNotify(objClusterEvent.EventTypeMajor + objClusterEvent.EventTypeMinor)
		Select Case objClusterEvent.Path_.Class
			Case "MSCluster_EventPropertyChange" : 
				WScript.Echo "EventProperty: " & objClusterEvent.EventProperty
			Case "MSCluster_EventResourceStateChange" : 
				WScript.Echo "EventGroup: " & objClusterEvent.EventGroup
				WScript.Echo "EventNewState: " & objClusterEvent.EventNewState
				WScript.Echo "EventNode: " & objClusterEvent.EventNode
				WScript.Echo "Cluster State: " & ResolveState(objClusterEvent.EventObjectType,objClusterEvent.EventNewState)
			Case "MSCluster_EventGroupStateChange" :
				WScript.Echo "EventNewState: " & objClusterEvent.EventNewState
				WScript.Echo "EventNode: " & objClusterEvent.EventNode
				WScript.Echo "Cluster State: " & ResolveState(objClusterEvent.EventObjectType,objClusterEvent.EventNewState)
		End Select
	End If
Loop
		

Function ResolveState(iEventType, iEventNewState)
	Dim arrStates(6,132)
	' Cluster Node State
	arrStates(0,0) = "Cluster Node State Unknown"
	arrStates(0,1) = "Cluster Node Up"
	arrStates(0,2) = "Cluster Node Down"
	arrStates(0,3) = "Cluster Node Paused"
	arrStates(0,4) = "Cluster Node Joining"
	' Cluster Group State
	arrStates(1,0) = "Cluster Group State Unknown"
	arrStates(1,1) = "Cluster Group Online"
	arrStates(1,2) = "Cluster Group Offline"
	arrStates(1,3) = "Cluster Group Failed"
	arrStates(1,4) = "Cluster Group Partial Online"
	arrStates(1,5) = "Cluster Group Pending" ' Not Documented
	' Cluster Resource State
	arrStates(2,0) = "Cluster Resource State Unknown"
	arrStates(2,1) = "Cluster Resource Inherited"
	arrStates(2,2) = "Cluster Resource Initializing"
	arrStates(2,3) = "Cluster Resource Online"
	arrStates(2,4) = "Cluster Resource Offline"
	arrStates(2,5) = "Cluster Resource Failed"
	arrStates(2,129) = "Cluster Resource Pending"
	arrStates(2,130) = "Cluster Resource Online Pending"
	arrStates(2,131) = "Cluster Resource Offline Pending"
	' Cluster Network State
	arrStates(4,0) = "Cluster Network State Unknown"
	arrStates(4,1) = "Cluster Network Unavailable"
	arrStates(4,2) = "Cluster Network Down"
	arrStates(4,3) = "Cluster Network Partitioned"
	arrStates(4,4) = "Cluster Network Up"
	' Cluster Network Interface State
	arrStates(5,0) = "Cluster Network Interface State Unknown"
	arrStates(5,1) = "Cluster Network Interface Unavailable"
	arrStates(5,2) = "Cluster Network Interface Failed"
	arrStates(5,3) = "Cluster Network Interface Unreachable"
	arrStates(5,4) = "Cluster Network Interface Up"
	
	ResolveState = arrStates(iEventType-2,iEventNewState+1)
End Function

Function GetClusterNotify(intEventType)
' This Function will Return the Cluster Notification on the Current Event Received
	If IsEmpty(objClusterNotifyDictionary) Then initDictionary
	If objClusterNotifyDictionary.Exists(CLng(Hex(intEventType))) Then
		GetClusterNotify = objClusterNotifyDictionary.Item(CLng(Hex(intEventType)))
	Else
		GetClusterNotify = "Event does not exists in the Dictionary"
	End If
End Function

Function initDictionary
' This Function will Initialize the Cluster Notification Dictionary
	' Key Value is in Hex Format - 40000000 = 0x40000000
	Set objClusterNotifyDictionary = CreateObject("Scripting.Dictionary")
	objClusterNotifyDictionary.Add 40000000, "The Cluster's Prioritized List Of Internal Networks Changes."
	objClusterNotifyDictionary.Add 00080000, "The Connection To The Cluster Identified Is Reestablsished After A Brief Disconnect. Some Events Generated Immidiately Before Or After This Event May Have Lost. You Need To Close All Open Connections And Reconnect To Recieve Accurate State Information."
	objClusterNotifyDictionary.Add 20000000, "The Cluster Becomes Unavailable, Meaning That All Attempts To Communicate With The Cluster Fail."
	objClusterNotifyDictionary.Add 00004000, "A New Group Is Created In The Cluster."
	objClusterNotifyDictionary.Add 00002000, "An Existing Group Is Deleted."
	objClusterNotifyDictionary.Add 00008000, "The Properties Of A Group Change Or When A Resource Is Added Or Removed From A Group."
	objClusterNotifyDictionary.Add 00001000, "A Group Changes State."
	objClusterNotifyDictionary.Add 80000000, "A Handle Associated With A Cluster Object Is Closed."
	objClusterNotifyDictionary.Add 04000000, "A New Network Interface Is Added To A Cluster Node."
	objClusterNotifyDictionary.Add 02000000, "A Network Interface Is Permanently Removed From A Cluster Node."
	objClusterNotifyDictionary.Add 08000000, "The Properties Of An Existing Network Interface Changes."
	objClusterNotifyDictionary.Add 01000000, "A Network Interface Changes State."
	objClusterNotifyDictionary.Add 00400000, "A Network Is Added To The Cluster Environment."
	objClusterNotifyDictionary.Add 00200000, "A Network Is Permanently Removed From The Cluster Environment."
	objClusterNotifyDictionary.Add 00800000, "The Properties Of An Existing Network Changes."
	objClusterNotifyDictionary.Add 00100000, "A Network Changes State."
	objClusterNotifyDictionary.Add 00000004, "A New Node Is Added To The Cluster. A Node Can Be Added Only When The Cluster Service Is Initially Installed On  The Node."
	objClusterNotifyDictionary.Add 00000002, "A Node Is Permanently Removed From A Cluster."
	objClusterNotifyDictionary.Add 00000008, "Reserved for future use."
	objClusterNotifyDictionary.Add 00000001, "A Node Changes State."
	objClusterNotifyDictionary.Add 10000000, "Reserved for future use."
	objClusterNotifyDictionary.Add 00000020, "A Cluster Database Key'S Attributes Are Changed. The Only Currently Defined Cluster Database Key Attributes Is Its Security Descriptor."
	objClusterNotifyDictionary.Add 00000010, "The Name Of A Cluster Database Key Has Changed."
	objClusterNotifyDictionary.Add 00000080, "Indicates that the other CLUSTER_CHANGE_REGISTRY events apply to the entire cluster database. If this flag is not included, the events apply only to the specified key."
	objClusterNotifyDictionary.Add 00000040, "A Value Of The Specified Cluster Database Key Is Changed Or Deleted."
	objClusterNotifyDictionary.Add 00000400, "A New Resource Is Created In The Cluster."
	objClusterNotifyDictionary.Add 00000200, "A Resource Is Deleted."
	objClusterNotifyDictionary.Add 00000800, "The Properties, Dependencies, Or Possible Owner Node Of A Resource Change."
	objClusterNotifyDictionary.Add 00000100, "A Resource Changes State."
	objClusterNotifyDictionary.Add 00020000, "A New Resource Type Is Created In The Cluster."
	objClusterNotifyDictionary.Add 00010000, "An Existing Resource Type Is Deleted."
	objClusterNotifyDictionary.Add 00040000, "The Properties Of A Resource Type Change."
End Function