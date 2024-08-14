'########################## ShowDrives.vbs ############################
Rem                        von h.r.roesler
Option Explicit
' Hint: This script cannot be executed with WScript.exe, use
'       CScript.exe at the command line as host instead. Example:
'       CScript.exe <Path>\ShowDrives.vbs
' For the base class' description (*Win32_DiskDrive*) look at:
' <http://msdn.microsoft.com/en-us/library/aa394132(VS.85).aspx>
' For further examples look beneath:
' *WMI Tasks: Disks and File Systems*
' <http://msdn.microsoft.com/en-us/library/aa394592(VS.85).aspx>
' To the used data media sizes see: <http://xkcd.com/394/> :-)

Const ASSOC = "ASSOCIATORS OF {", ASS_CLASS = "} WHERE AssocClass = "
Const BT = " Byte", MIB = " MiB", GiB = " GiB", GB = " GB"
Const WINMGMTS = "winmgmts:{impersonationLevel=impersonate}!root/cimv2"
Dim arr, wmi, wmiDisk, i, str, y, TAB, Echo, dblCap, dblSum, intIndex

Set Echo = WScript.StdOut
TAB = Space(4)
str = vbCRLF & "In Windows the sizes of data media are counted in " & _
     "powers to the base of 2." & vbCRLF & vbCRLF & "Note: The " & _
     "based unit for 'KiB' used in this script is being equivalent" & _
     " to" & vbCRLF & "      the tenth power of 2 or " & 2^10 & _
     " Bytes. Accordingly, that applies to all" & vbCRLF & "      " & _
     "other quantity units in Byte, if not specified else."
Echo.WriteLine str
arr = Array("Unknown", "Other", "Sequential Access", "Random Access", _
            "Supports Writing", "Encryption", "Compression", _
            "Supports Removable Media", "Manual Cleaning", _
            "Automatic Cleaning", "SMART Notification", _
            "Supports Dual-Sided Media", "Ejection Prior to Drive " & _
            "Dismount Not Required")

Set wmi = GetObject(WINMGMTS)

For Each wmiDisk In wmi.InstancesOf("Win32_DiskDrive")
    With wmiDisk
        Echo.WriteLine vbCRLF & Div(72, "+", .Caption & " (" & _
                                            .DeviceID & ")") & vbCRLF
        intIndex = .Index
        Echo.WriteLine "Index:               " & intIndex
        Echo.WriteLine "Interface Type:      " & .InterfaceType

        Echo.WriteLine "Status:              " & .Status
        If IsNull(.Size) Then
            Echo.WriteLine "                     No Drives Connected."
        Else
            Echo.WriteLine "Capacity:            " & _
                         Round(CDbl(.Size) / 2^30, 3) & GiB
            dblCap = Round(CDbl(.Size) / 10^9, 3)
            Echo.WriteLine "To the decimal base: " & dblCap & GB
            Echo.WriteLine "Bytes Per Sector:    " & .BytesPerSector & BT
            Echo.WriteLine "Total Sectors:       " & CDbl(.TotalSectors)
            Echo.WriteLine "Total Cylinders:     " & CDbl(.TotalCylinders)
            Echo.WriteLine "Total Heads:         " & .TotalHeads
            Echo.WriteLine "Total Tracks:        " & CDbl(.TotalTracks)
            Echo.WriteLine "Sectors Per Track:   " & .SectorsPerTrack
            Echo.WriteLine "Tracks Per Cylinder: " & .TracksPerCylinder
            For i = 0 To Ubound(.Capabilities)
                y = 21
                If i = 0 Then Echo.Write "Capabilities:": y = 8
                Echo.WriteLine Space(y) & arr(.Capabilities(i))
            Next
            Echo.WriteLine "Manufacturer:        " & .Manufacturer
            Echo.WriteLine "Media Type:          " & .MediaType
            Echo.WriteLine "Num Logical Drives:  " & .Partitions
            Echo.WriteLine "Signature:           " & Hex(.Signature)
        End If
        If Not(IsNull(.SCSILogicalUnit)) Then
            Echo.WriteLine vbCRLF & "SCSI Bus:            " & .SCSIBus
            Echo.WriteLine "SCSI Logical Unit:   " & .SCSILogicalUnit
            Echo.WriteLine "SCSI TargetId:       " & .SCSITargetId
            Echo.WriteLine "SCSI Port:           " & .SCSIPort
        End If
        str = ASSOC & "Win32_DiskDrive.DeviceID=""" & Replace(.DeviceID, _
              "\", "\\") & """" & ASS_CLASS & "Win32_DiskDriveToDiskPartition"
        Call PartiFromDrive(str)
        'Call ExecDiskPart("ListDisk.dsk", intIndex)
    End With
    dblSum = dblSum + dblCap
Next
Echo.WriteLine "Total Capacity: " & dblSum & GB

Sub PartiFromDrive(strWQL)        ' Antecedent: Win32_DiskDrive
    Dim wmiPart, str              ' Dependent:  Win32_DiskPartition

    For Each wmiPart In wmi.ExecQuery(strWQL)
        With wmiPart
            Echo.WriteLine vbCRLF & Space(4) & Div(68,"-", .DeviceID)
            Echo.WriteLine vbCRLF & TAB & TAB & "Type:     " & .Type
            Echo.WriteLine vbCRLF & TAB & TAB & "Capacity: " & _
                           FormatNumber(CDbl(.Size) / 2^20, 2) & MiB
            Echo.WriteLine vbCRLF & TAB & TAB & "4KAligned: " & _
                           Is4KAligned(CDbl(.StartingOffset)) & vbCRLF
            Echo.WriteLine vbCRLF & TAB & TAB & "Primary Partition: " & _
                                                        .PrimaryPartition
            If IsNull(.Bootable) Then str = "No": Else str = .Bootable
            Echo.WriteLine vbCRLF & TAB & TAB & "Bootable: " & str
            Echo.WriteLine vbCRLF & TAB & TAB & "Active: " & .BootPartition
            Call PartiFromLogicDisk(.DeviceID)
        End With
    Next
End Sub

Sub PartiFromLogicDisk(strDID)      ' Antecedent: Win32_DiskPartition
    Dim wmiLogDrv, wmiLog, TAB, str ' Dependent:  Win32_LogicalDisk

    TAB = Space(6)
    str = ASSOC & "Win32_DiskPartition.DeviceID=""" & strDID & """" & _
          ASS_CLASS & "Win32_LogicalDiskToPartition"
    Set wmiLogDrv = wmi.ExecQuery(str)
    If wmiLogDrv.Count = 0 Then Exit Sub
    Echo.WriteLine vbCRLF & Space(8) & Div(64, "~", "Logical Drives")
    For Each wmiLog In wmiLogDrv
        With wmiLog
            Echo.WriteLine vbCRLF & TAB & TAB & .DeviceID & TAB & _
                                                 "    " & .VolumeName
            If .VolumeDirty Then
                Echo.WriteLine TAB & TAB & "*** Lets exec a ChkDsk-Run! ***"
            End If
            str = "Not ready."
            If Len(.FileSystem) > 0 Then str = .FileSystem
            Echo.WriteLine TAB & TAB & "Filesystem: " & str
            If Not(IsNull(.Size)) And Not(IsNull(.FreeSpace)) Then
                Echo.WriteLine TAB & TAB & "Tot. Space: " & _
                                    Round(CDbl(.Size) / 2^20, 2) & MiB
                Echo.WriteLine TAB & TAB & "Free Space: " & _
                                Round(CDbl(.FreeSpace) / 2^20, 2) & MiB
            End If
            Echo.WriteLine
        End With
    Next
End Sub

Function Is4KAligned(dblBig)
    Is4KAligned = (dblBig / 2^12 - Fix(dblBig / 2^12)) = 0
End Function

Function Div(lngLen, strChar, strIns)
    Dim str

    str = String(lngLen, strChar) & vbCRLF & strIns
    If Len(strIns) + 2 < lngLen Then
        strIns = " " & strIns & " "
        str = String((lngLen - Len(strIns)) \ 2, strChar) & strIns
        str = str & String(lngLen - Len(str), strChar)
    End If

    Div = str
End Function

Sub ExecDiskPart(strFile, intNum)
    Const TemporaryFolder = 2, COMMAND = "DiskPart.exe /s "
    Dim arr, str

    arr = Array("select disk " & intNum, "detail disk", "list partition", _
                "list volume", "exit")
    With CreateObject("Scripting.FileSystemObject")
        strFile = .BuildPath(.GetSpecialFolder(TemporaryFolder), strFile)
        With .CreateTextFile(strFile, True)
            For Each str In arr
                .WriteLine str
            Next
            .Close
        End With
        With CreateObject("WScript.Shell").Exec(COMMAND & strFile)
            .StdIn.Close
            Echo.Write .StdOut.ReadAll
        End With
        .DeleteFile strFile, True
    End With
End Sub
'########################## ShowDrives.vbs ############################
