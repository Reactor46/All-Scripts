strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colVolumes = objWMIService.ExecQuery _
    ("Select * from Win32_Volume Where Label = 'Recovery'")
For Each objVolume in colVolumes
    objVolume.DriveLetter = "R:"
    objVolume.Put_
Next	