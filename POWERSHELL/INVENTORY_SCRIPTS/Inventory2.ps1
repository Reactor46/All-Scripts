<# AD-Inventory 4.0  By: xXGh057Xx #>


$today = Get-Date
$cutoffdate = $today.AddDays(-15)

Get-ADComputer  -Properties * -Filter {LastLogonDate -gt $cutoffdate}|Select -Expand DNSHostName  | out-file G:\All-Computers.txt
$ComputerList = get-content "G:\All-Computers.txt"
$Amount = $ComputerList.count
$a=0

foreach ($hosts in $ComputerList) 
        {
                        $Ping = Test-Path "\\$hosts\C$" -ErrorAction SilentlyContinue
                        if ($Ping -eq "True") 
                           {
                                echo $hosts >> "G:\Online-Computers.txt"
                                  $a++
                                    Write-Progress -Activity "Working..." -CurrentOperation "$a complete of $Amount"  -Status "Please wait testing connections" 

                                
                           }

        }

$allhost = get-content "G:\Online-Computers.txt"
"Hostname,MAC Address,Serial Number" >> C:\Inventory.csv
$a=0
$OnlineAmount = $allhost.count

foreach ($computer in $allhost) {

$Networks = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $computer | ? {$_.IPEnabled}
   foreach ($Network in $Networks) {
    $IsDHCPEnabled = $false
    If($Network.DHCPEnabled) {
     $IsDHCPEnabled = $true
         }
          $mac = $Network.MACAddress
         }

      $enclosure = Get-WmiObject -Class win32_systemenclosure -ComputerName $computer
        $serial = $enclosure.SerialNumber
          $output = $computer + "," + $mac + "," + $serial
            $output >> G:\Inventory.csv
               $a++
                    
                         
          Write-Progress -Activity "Working..." -CurrentOperation "$a complete of $OnlineAmount"  -Status "Please wait testing connections"
    }


        Del "G:\Online-Computers.txt"
           Del "G:\All-Computers.txt"