########################################################################
#  Powershell NetMon Consolidation Script#  
#  Script that runs all of the PSNetMon scripts
#  Created by: Brad Voris
########################################################################

# Run PSNetMon Count Script for count of monitored resources
Invoke-Command {C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\scripts\psnetmoncountscript.ps1}

# Run PSNetMon ICMP Script to ping resources
Invoke-Command {C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\scripts\psnetmonicmpscript.ps1}

# Run PSNetMon Port Script to check if ports are open
Invoke-Command {C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\scripts\psnetmonportscript.ps1}

# Run PSNetMon Service Script checks if services are running
Invoke-Command {C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\scripts\psnetmonservicescript.ps1}

# Run PSNetMon Service Script runs RSS Feed module
Invoke-Command {C:\Scripts\Repository\jbattista\Web\Reports\PSNetMon\scripts\PSNetMonRSSTicker.ps1}