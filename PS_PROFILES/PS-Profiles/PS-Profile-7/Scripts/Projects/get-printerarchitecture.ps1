<#
.SYNOPSIS
Lists printer drivers,  whether 32- and 64-bit versions are installed,
and how many printers are using each driver.
 
.PARAMETER Printers
Includes a counter of the printers using each printer driver. On a system
with many installed printers, this can take a a little time, so this
functionality is optional.
 
.EXAMPLE
PS C:\local\scripts> .\Get-PrinterDriverArchitecture.ps1 | format-table -auto
 
Name                                       x86   x64
----                                       ---   ---
Brother HL-5250DN                         True False
Epson LQ-570+ ESC/P 2                     True False
HP Business Inkjet 2230/2280              True False
HP Business Inkjet 2250 (PCL5C)           True  True
HP Business Inkjet 2800 PCL 5             True False
 
This command lists the installed printer drivers, and whether 32-bit (x86)
or 64-bit (x64) drivers are available.
 
.EXAMPLE
PS C:\local\scripts> .\Get-PrinterDriverArchitecture.ps1 | where x64 -eq $false | ft -a
 
Name                                      x86   x64
----                                      ---   ---
Brother HL-5250DN                        True False
Epson LQ-570+ ESC/P 2                    True False
HP Business Inkjet 2230/2280             True False
HP Business Inkjet 2800 PCL 5            True False
 
This command uses the Where[-Object] cmdlet to filter out those drivers that
have a 64-bit driver installed.
 
.EXAMPLE
PS C:\local\scripts> .\Get-PrinterDriverArchitecture.ps1 -Printers | ft -a
Enumerating printers:
    WID - Kyocera TASKalfa 3050ci KX
    WC - Kyocera TASKalfa 300ci KX
    UFS - Canon iR-ADV C5035
    ...
 
Name                                       x86   x64 Printers
----                                       ---   --- --------
Brother HL-5250DN                         True False        5
Epson LQ-570+ ESC/P 2                     True False        0
HP Business Inkjet 2230/2280              True False        1
HP Business Inkjet 2250 (PCL5C)           True  True        0
HP Business Inkjet 2800 PCL 5             True False        1
 
This command includes the -Printers switch parameter to add a count of
the printers using each driver. Enumerating the printers can take a while
if there are lots of them installed, so this behavior is optional.
 
.NOTES
 - Author    : Geoffrey.Duke@uvm.edu
 - Mod. Date : May 28, 2013
#>
 
param( [switch] $printers )
 
$wmi_drivers = get-wmiobject Win32_PrinterDriver -Property Name
$drivers = @{}
foreach ($driver in $wmi_drivers) {
    # Isolate the driver name and platform
    $name,$null,$platform = $driver.Name -split ','
 
    if ( -not $drivers[$name] ) {
        switch ( $platform ) {
            'Windows NT x86' { $drivers[$name] = [ordered]@{
                                 'Name'=$name; 'x86'=$true; 'x64'=$false }; break }
            'Windows x64'    { $drivers[$name] = [ordered]@{
                                 'Name'=$name; 'x86'=$false; 'x64'=$true }; break }
             default         { write-warning "Unexpect platform $platform on driver $name"}
        }
    }
    else {
        switch ( $platform ) {
            'Windows x64'    { $drivers[$name]['x64'] = $true; break }
            'Windows NT x86' { $drivers[$name]['x86'] = $true; break }
             default         { write-warning "Unexpect platform $platform on driver $name"}
        }
    }
}
 
if ( $printers ) {
    # Initialize all printer counts
    $drivers.keys | foreach { $drivers[$_]['Printers'] = 0 }
    # Add a count of the number of printers using each driver
    # With some progress info
    write-host 'Enumerating printers:'
    $count = 0
    get-wmiobject Win32_Printer -Property Name,DriverName | foreach {
        write-host "    $($_.Name)"  -foreground darkgray
        $drivers[$_.DriverName]['Printers']++
        $count++
    }
 
    write-host "Retrieved $count printers"
}
 
# Output collection of objects
$drivers.keys | sort | foreach { New-Object PSObject -Property $drivers[$_] } 