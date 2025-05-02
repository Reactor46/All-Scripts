          $ServerList = "LASMT01","LASMT02","LASMT03","LASMT04","LASMT05" #Get-Content -Path E:\PSScripts\Conf\MTSrvs.txt

          $TimeToRun=Measure-Command {
            ForEach ($ServerName in $ServerList)
              {
              $PingStatus = Test-Connection -ComputerName $ServerName -Count 1 -Quiet

              ## If server responds, get uptime and disk info
              If ($PingStatus)
              {

                $OperatingSystem = Get-CimInstance Cim_OperatingSystem -ComputerName $ServerName
            
                $CpuUsage = Get-CimInstance Cim_Processor -Computername $ServerName |
                            Measure-Object -Property LoadPercentage -Average |
                                ForEach-Object {$_.Average; "% Used"}
                $Uptime = Get-Uptime($OperatingSystem.LastBootUpTime)
            
                $MemUsage = Get-CimInstance Cim_OperatingSystem -ComputerName $ServerName |
                            ForEach-Object {"{0:N0}" -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory) * 100)/ $_.TotalVisibleMemorySize);"% Used"}
                            
                $IPv4Address = Get-NetIPAddress -CimSession $ServerName -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notmatch 'Loopback'} | Select-Object IPAddress
            
                $Service = Get-Process -ComputerName $ServerName -Name DataLayerService -ErrorAction SilentlyContinue |
                        ForEach-Object {"{0:N0}" -f ([Math]::Round($_.WS/1MB)); "MB"}
            
                $Requests = Get-Counter -ComputerName $ServerName -Counter "\asp.net\requests queued" -ErrorAction SilentlyContinue |
                            ForEach-Object {($_.CounterSamples.CookedValue)}
            
                $TotalCon = Get-Counter -ComputerName $ServerName -Counter "\web service(_total)\current connections" -ErrorAction SilentlyContinue |
                            ForEach-Object {($_.CounterSamples.CookedValue)}
            
                $TCPCon =  Get-Counter -ComputerName $ServerName -Counter "\TCPv4\Connections Established" -ErrorAction SilentlyContinue |
                           ForEach-Object {($_.CounterSamples.CookedValue)}
                         } # End If ($PingStatus)

                         }
                         }
          $ServerName, $OperatingSystem, $CpuUsage, $Uptime, $MemUsage, $IPv4Address, $Service, $Requests, $TotalCon, $TCPCon, $TimeToRun.TotalSeconds