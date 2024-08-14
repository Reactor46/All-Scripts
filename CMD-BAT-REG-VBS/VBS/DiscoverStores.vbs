'==================================================================================
' Script: 	DiscoverStore.vbs
' Date:		4/27/09
' Author: 	Brian Wren, Microsoft Consulting Services
' Purpose:	Discovers Store class and relationship with Store Server and Client Servers for StoreApp sample application
'==================================================================================

'Constants used for event logging
SetLocale("en-us")
Const SCRIPT_NAME			= "DiscoverStore.vbs"
Const EVENT_LEVEL_ERROR	 	= 1
Const EVENT_LEVEL_WARNING 	= 2
Const EVENT_LEVEL_INFO 		= 4

Const SCRIPT_STARTED		= 801
Const CLASS_CREATED			= 802
Const RELATIONSHIP_CREATED	= 803
Const SCRIPT_ENDED			= 805

'Setup variables sent in through script arguments
SourceId = WScript.Arguments(0) 				'GUID of discovery calling the script.  Provided by the MPElement variable.
ManagedEntityId = WScript.Arguments(1)			'GUID of target object.  Provided by the Target/Id variable.
sComputerName = WScript.Arguments(2)			'Name of the computer holding the Store Server or Store Client class.
sStoreCode = WScript.Arguments(3)				'StoreCode of the Store to create.  Taken from the registry of the target computer.
sServerOrClient = LCase(WScript.Arguments(4))	'String of "server" or "client" depending on which type of class is calling script.
bDebug = CBool(WScript.Arguments(5))			'If true, information events are loggged.


'Start by setting up API object and creating a discovery data object.
'Discovery data object requires the MPElement and Target/ID variables.  The first argument in the method is always 0.
Set oAPI = CreateObject("MOM.ScriptAPI")
Set oDiscoveryData = oAPI.CreateDiscoveryData(0, SourceId, ManagedEntityId)

'Log a message that script is starting only if Debug argument is True
sMessage =	"Script started" & VbCrLf & _
"Source ID: " & SourceId & VbCrLf & _
"Managed Entity ID: " & ManagedEntityId & VbCrLf & _
"Computer Name: " & sComputerName & VbCrLf & _
"Store Code: " & sStoreCode & VbCrLf & _
"Server or Client: " & sServerOrClient
Call LogDebugEvent(SCRIPT_STARTED,sMessage)


'Create an instance of the store class and add it to the discovery data.
'The StoreCode property is required because it is the key property of the class.
Set oStoreInstance = oDiscoveryData.CreateClassInstance("$MPElement[Name='MPAuthor.Stores.Store']$")
oStoreInstance.AddProperty "$MPElement[Name='MPAuthor.Stores.Store']/StoreCode$", sStoreCode
oDiscoveryData.AddInstance(oStoreInstance)

sMessage =	"Created store class" & VbCrLf & _
"Store Code: " & sStoreCode
Call LogDebugEvent(CLASS_CREATED,sMessage)

'Create an instance of the appropriate classes depending on whether a Store Server or Store Client is calling the script.
If sServerOrClient = "server" Then
    'Create a class instance of Store Server and a Relationship Instance of Store Contains Store Server.
    Set oComputerInstance = oDiscoveryData.CreateClassInstance("$MPElement[Name='MPAuthor.Stores.ComputerRole.StoreServer']$")
    Set oRelationshipInstance = oDiscoveryData.CreateRelationshipInstance("$MPElement[Name='MPAuthor.Stores.StoreContainsStoreServer']$")
Else
    'Create a class instance of Store Client and a Relationship Instance of Store Contains Store Client.
    Set oComputerInstance = oDiscoveryData.CreateClassInstance("$MPElement[Name='MPAuthor.Stores.ComputerRole.StoreClient']$")
    Set oRelationshipInstance = oDiscoveryData.CreateRelationshipInstance("$MPElement[Name='MPAuthor.Stores.StoreContainsClients']$")
End If


'Provide the PrincipalName property for the computer instance created above and add to the discovery data.
'This is required because both the Store Server and Store Client classes are based on Windows ComputerRole.
'	Windows ComputerRole is hosted by Windows Computer.
'	When creating a new instance of a class, we need to provide the key properties of that class and any hosting classes.
oComputerInstance.AddProperty "$MPElement[Name='Windows!Microsoft.Windows.Computer']/PrincipalName$", sComputerName
oDiscoveryData.AddInstance(oComputerInstance)

sMessage =	"Created computer role class" & VbCrLf & _
"Computer Name: " & sComputerName
Call LogDebugEvent(CLASS_CREATED,sMessage)

'With the instance of Store and either Store Server or Store Client created, we can set the Source and Target of the relationship.
oRelationshipInstance.Source = oStoreInstance
oRelationshipInstance.Target = oComputerInstance
oDiscoveryData.AddInstance(oRelationshipInstance)

sMessage =	"Created relationship"
Call LogDebugEvent(RELATIONSHIP_CREATED,sMessage)

'Return the discovery data.
oAPI.Return(oDiscoveryData)

Call LogDebugEvent (SCRIPT_ENDED,"Script ended.")

'==================================================================================
' Sub:		LogDebugEvent
' Purpose:	Logs an informational event to the Operations Manager event log
'			only if Debug argument is true
'==================================================================================
Sub LogDebugEvent(EventNo,Message)

    Message = VbCrLf & Message
    If bDebug = True Then
        Call oAPI.LogScriptEvent(SCRIPT_NAME,EventNo,EVENT_LEVEL_INFO,Message)
    End If

End Sub
