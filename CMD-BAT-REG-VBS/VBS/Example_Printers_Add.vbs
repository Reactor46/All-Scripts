strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

set objNewPort = objWMIService.get("Win32_TCPIPPrinterPort").SpawnInstance_
Set objPrinter = objWMIService.Get("Win32_Printer").SpawnInstance_
objWMIService.Security_.Privileges.AddAsString "SeLoadDriverPrivilege", True
Set objDriver = objWMIService.Get("Win32_PrinterDriver")

'  Install printer ports
objNewPort.Name = "flanner4.flanner.nd.edu"
objNewPort.Protocol = 1
objNewPort.HostAddress = "flanner4.flanner.nd.edu"
objNewPort.PortNumber = 9100
objNewPort.SNMPEnabled = True
objNewPort.Put_

objNewPort.Name = "flanner5.flanner.nd.edu"
objNewPort.Protocol = 1
objNewPort.HostAddress = "flanner5.flanner.nd.edu"
objNewPort.PortNumber = 9100
objNewPort.SNMPEnabled = True
objNewPort.Put_

objNewPort.Name = "flanner6.flanner.nd.edu"
objNewPort.Protocol = 1
objNewPort.HostAddress = "flanner6.flanner.nd.edu"
objNewPort.PortNumber = 9100
objNewPort.SNMPEnabled = True
objNewPort.Put_

objNewPort.Name = "flanner7.flanner.nd.edu"
objNewPort.Protocol = 1
objNewPort.HostAddress = "flanner7.flanner.nd.edu"
objNewPort.PortNumber = 9100
objNewPort.SNMPEnabled = True
objNewPort.Put_

objNewPort.Name = "flanner8.flanner.nd.edu"
objNewPort.Protocol = 1
objNewPort.HostAddress = "flanner8.flanner.nd.edu"
objNewPort.PortNumber = 9100
objNewPort.SNMPEnabled = True
objNewPort.Put_

objNewPort.Name = "flanner9.flanner.nd.edu"
objNewPort.Protocol = 1
objNewPort.HostAddress = "flanner9.flanner.nd.edu"
objNewPort.PortNumber = 9100
objNewPort.SNMPEnabled = True
objNewPort.Put_

'--------------------------------------------------

'  Install printer driver HP Laserjet 4200
objDriver.Name = "HP LaserJet 4200 PCL 5e"
objDriver.SupportedPlatform = "Windows NT x86"
objDriver.Version = "3"
objDriver.FilePath = "z:\\drivers\\HP LJ4200\\"
objDriver.Infname = "z:\\drivers\\HP LJ4200\\hpc4200b.inf"
intResult = objDriver.AddPrinterDriver(objDriver)

'Laserjet 4100
'Installs Printer Driver, one already in the windows driver subset
objDriver.Name = "HP LaserJet 4100 Series PCL"
objDriver.SupportedPlatform = "Windows NT x86"
objDriver.Version = "3"
intResult = objDriver.AddPrinterDriver(objDriver)


'--------------------------------------------------

'  Install printer objects
objPrinter.DriverName = "HP LaserJet 4200 PCL 5e"
objPrinter.PortName   = "flanner4.flanner.nd.edu"
objPrinter.DeviceID   = "Flanner 4"
objPrinter.Network = False
objPrinter.Shared = False
objPrinter.Put_

objPrinter.DriverName = "HP LaserJet 4200 PCL 5e"
objPrinter.PortName   = "flanner5.flanner.nd.edu"
objPrinter.DeviceID   = "Flanner 5"
objPrinter.Network = False
objPrinter.Shared = False
objPrinter.Put_

objPrinter.DriverName = "HP LaserJet 4200 PCL 5e"
objPrinter.PortName   = "flanner6.flanner.nd.edu"
objPrinter.DeviceID   = "Flanner 6"
objPrinter.Network = False
objPrinter.Shared = False
objPrinter.Put_

objPrinter.DriverName = "HP LaserJet 4200 PCL 5e"
objPrinter.PortName   = "flanner7.flanner.nd.edu"
objPrinter.DeviceID   = "Flanner 7"
objPrinter.Network = False
objPrinter.Shared = False
objPrinter.Put_

objPrinter.DriverName = "HP LaserJet 4200 PCL 5e"
objPrinter.PortName   = "flanner8.flanner.nd.edu"
objPrinter.DeviceID   = "Flanner 8"
objPrinter.Network = False
objPrinter.Shared = False
objPrinter.Put_

objPrinter.DriverName = "HP LaserJet 4100 Series PCL"
objPrinter.PortName = "flanner9.flanner.nd.edu"
objPrinter.DeviceID = "Flanner 9"
objPrinter.Location = ""
objPrinter.Network = True
objPrinter.Put_
