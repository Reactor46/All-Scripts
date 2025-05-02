SourceId = WScript.Arguments(0)
ManagedEntityId = WScript.Arguments(1)
sComputerName = WScript.Arguments(2)
 
Set oAPI = CreateObject("MOM.ScriptAPI")
Set oDiscoveryData = oAPI.CreateDiscoveryData(0, SourceId, ManagedEntityId)
 
For i = 1 to 3
   Set oInstance = oDiscoveryData.CreateClassInstance("$MPElement[Name='MyMP.MyApplicationComponent']$")
   oInstance.AddProperty "$MPElement[Name='Windows!Microsoft.Windows.Computer']/PrincipalName$", sComputerName
   oInstance.AddProperty "$MPElement[Name='MyMP.MyApplicationComponent']/ComponentName$", "Component" & i
   oDiscoveryData.AddInstance(oInstance)
Next
oAPI.Return(oDiscoveryData)
