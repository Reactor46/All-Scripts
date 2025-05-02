'The sample scripts are not supported under any Microsoft standard support 
'program or service. The sample scripts are provided AS IS without warranty  
'of any kind. Microsoft further disclaims all implied warranties including,  
'without limitation, any implied warranties of merchantability or of fitness for 
'a particular purpose. The entire risk arising out of the use or performance of  
'the sample scripts and documentation remains with you. In no event shall 
'Microsoft, its authors, or anyone else involved in the creation, production, or 
'delivery of the scripts be liable for any damages whatsoever (including, 
'without limitation, damages for loss of business profits, business interruption, 
'loss of business information, or other pecuniary loss) arising out of the use 
'of or inability to use the sample scripts or documentation, even if Microsoft 
'has been advised of the possibility of such damages.

Option Explicit

Dim strComputer   : strComputer = "."
Dim objWMIService : Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Dim diskDrives    : Set diskDrives = objWMIService.ExecQuery("Select * from Win32_DiskDrive",,48)


'Main
WScript.Echo Check4kAligned


'Check 4k alignment issue
Function Check4kAligned()
	Dim outputMsg : outputMsg = ""
	Dim is4kAligned,name,deviceID,diskSize,partitionSize,description
	diskSize = 0
	partitionSize = 0
	
	Dim diskDrive,diskPartition,diskPartitions,logicalDisk,logicalDisks
	
	'for each disks
	For Each diskDrive In diskDrives
		'Convert disk drive to disk partition
		Set diskPartitions = objWMIService.ExecQuery _
            				("ASSOCIATORS OF {Win32_DiskDrive.DeviceID=""" & _
                			Replace(diskDrive.DeviceID,"\","\\") & """} WHERE AssocClass = " & _
                   			"Win32_DiskDriveToDiskPartition")
                   			 
        For Each diskPartition In diskPartitions
        	is4kAligned = False
        	
        	If (MMod(diskPartition.StartingOffset,4096) = 0) Then
        		is4kAligned = True
        	End If
        	
        	name = diskPartition.Name
        	'Convert Byte to GB and save as 2 decimal places: 0.00
        	partitionSize = FormatNumber(diskPartition.Size / 1073741824,,,,0)
	        description = diskPartition.Description
	        
	        'Convert disk partition to logical disk
	        Set logicalDisks = objWMIService.ExecQuery _
	        				("ASSOCIATORS OF {Win32_DiskPartition.DeviceID=""" & _
	                		diskPartition.DeviceID & """} WHERE AssocClass = " & _
	                   		"Win32_LogicalDiskToPartition")
	        
	        If logicalDisks.Count =0 Then
	        	deviceID = ""
	        	diskSize = 0
	        	outputMsg  = outputMsg + FormatResult(is4kAligned,name,deviceID,diskSize,partitionSize,description)
	        End If
	        
			For Each logicalDisk In logicalDisks
				deviceID = logicalDisk.DeviceID
				diskSize = FormatNumber(logicalDisk.Size / 1073741824,,,,0)
				outputMsg = outputMsg + FormatResult(is4kAligned,name,deviceID,diskSize,partitionSize,description)
			Next
        Next
	Next
	
	Check4kAligned = outputMsg
End Function

'Internal function: Format result
Function FormatResult(is4kAligned,name,deviceID,diskSize,partitionSize,description)
	Dim outputMsg
	outputMsg = outputMsg & "[(Is4kAligned: " & CStr(is4kAligned) & ") (Name: " & _
				name & ") (DeviceID: " & deviceID & ")" & "(DiskSize_GB: " & diskSize & ")" & _
				"(PartitionSize_GB: " & partitionSize & ")" & "(Description: " & _
				description & ")];" & vbCrLf & vbCrLf
				
	FormatResult = outputMsg
End Function

'Internal function: Big number mod
Function MMod(input,mValue)
	dim dblValue,temp
	dblValue = CDbl(input)
	temp = Fix(dblValue / mValue)
	MMod = dblValue - (temp * mValue)
End Function

