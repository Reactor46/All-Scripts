
' Developed by Romano Jerez
' 

' Unassigns (and removes) packages listed in a file (package IDs),
' one per line, from
' all distribution points that belong to the specified SMS Site

' Run the script on the SMS server where the package was created


' *** SET THESE VARIABLES ***

' This variable holds the file name (with full path) that contains
' the package IDs (one per line) of the packages that will be
' removed from the distribution points that belong to the specified SMS site:

sourceFile = "E:\scripts\packagesToRemoveFromSite.txt"

' this variable represent the SMS site code where all of its DPs
' will be unassigned (and removed) from each of the packages listed
' in the sourceFile:

strSiteCode = "XXX"



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

   UnAssignPackage objSWbemServices, package

Loop


objFile1.Close


Sub UnAssignPackage(connection, pkgID)

   strQuery = "SELECT * FROM SMS_DistributionPoint WHERE " & _
              "PackageID = '" & pkgID & "' AND SiteCode =" & _
              "'" & strSiteCode & "'"  

   Set DPs = connection.ExecQuery(strQuery, , wbemFlagForwardOnly Or wbemFlagReturnImmediately)

   For Each dp in DPs
      WScript.Echo "Unassigning " & dp.PackageID & " from " & dp.ServerNALPath
      dp.Delete_

      If Err.number <> 0 Then
           WScript.Echo "Failed to delete " & pkgID
      End If
   Next
End Sub
