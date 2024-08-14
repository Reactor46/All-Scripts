' Author: Romano Jerez, NETvNext.com

' Purpose: Assign PKGs listed in a file (Package IDs without extensions) 
'          to an SCCM distribution point

' To be executed on site server

' OK to use and distribute as long as Author information is kept.


' *** SET THESE VARIABLES ***

sourceFile = "e:\scripts\packages.txt"

' these variables represent the target DP:

strSiteCode = "ST1"
strServerName = "sccmcentral"


' ********************************

Const ForReading = 1


' Connect to SCCM Provider on local machine

set objSwbemLocator = CreateObject("WbemScripting.SWbemLocator")
set objSWbemServices= objSWbemLocator.ConnectServer(".", "root\sms")
Set ProviderLoc = objSWbemServices.InstancesOf("SMS_ProviderLocation")


For Each Location In ProviderLoc
      If Location.ProviderForLocalSite = True Then
          Set objSWbemServices = objSWbemLocator.ConnectServer(Location.Machine, "root\sms\site_" + Location.SiteCode)
      End If
Next



Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile1 = objFSO.OpenTextFile(sourceFile, ForReading)


Do Until (objFile1.AtEndOfStream)

   package = objFile1.Readline

   SWDAssignPackageToDistributionPoint objSWbemservices, package, strSiteCode, strServerName

Loop


objFile1.Close



' From SDK

Sub SWDAssignPackageToDistributionPoint(connection, existingPackageID, siteCode, serverName)

    ' Create distribution point object (this is not an actual distribution point).
    Set distributionPoint = connection.Get("SMS_DistributionPoint").SpawnInstance_

    ' Associate the existing package with the new distribution point object.
    distributionPoint.PackageID = existingPackageID     

    ' This query selects a single distribution point based on the provided SiteCode and ServerName.
    query = "SELECT * FROM SMS_SystemResourceList WHERE NALPath NOT LIKE '%PXE%' AND RoleName='SMS Distribution Point' AND SiteCode='" & siteCode & "' AND ServerName='" & serverName & "'"

    Set listOfResources = connection.ExecQuery(query, , wbemFlagForwardOnly Or wbemFlagReturnImmediately)

    ' The query returns a collection that needs to be enumerated (although we should only get one instance back).
    For Each resource In ListOfResources      
        distributionPoint.ServerNALPath = Resource.NALPath
        distributionPoint.SiteCode = Resource.SiteCode  
    Next
    
    ' Save the distribution point instance for the package.
    distributionPoint.Put_ 
    
    ' Display notification text.
    Wscript.Echo "Assigned package: " & distributionPoint.PackageID 

End Sub
