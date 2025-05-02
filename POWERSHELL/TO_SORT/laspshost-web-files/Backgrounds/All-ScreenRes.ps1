<#Get-ADComputer -Filter {Operatingsystem -Like 'Windows 7*' -and Enabled -eq 'true'}  -ErrorAction SilentlyContinue -Properties * | Select -ExpandProperty Name | Out-File -FilePath C:\LazyWinAdmin\Backgrounds\AllPCs.log -Append
Get-ADComputer -Filter {Operatingsystem -Like 'Windows 10*' -and Enabled -eq 'true'}  -ErrorAction SilentlyContinue -Properties * | Select -ExpandProperty Name | Out-File -FilePath C:\LazyWinAdmin\Backgrounds\AllPCs.log -Append
Get-Content 'C:\LazyWinAdmin\Backgrounds\AllPCs.log' | 
 ForEach { if (test-connection $_ -quiet) { write-output "$_" | Out-File 'C:\LazyWinAdmin\Backgrounds\Alive.log' -append
  } else { 
  write-output "$_ is Dead!!!" | Out-File 'C:\LazyWinAdmin\Backgrounds\Dead.log' -append}}
  #>
# read in the workstation names
$workstation_list=get-content "C:\LazyWinAdmin\Backgrounds\Alive.log"
foreach($pc in $workstation_list)
{
   # Ping the machine to see if it's turned on
   $query = "select * from win32_pingstatus where address = '$pc'"
   $result = Get-WmiObject -query $query
   
   if ($result.protocoladdress) {

      # Get the display details via WMI
      $displays= Get-WmiObject -class "Win32_DisplayConfiguration" -computername $pc
      #Get-WmiObject -ComputerName $comp win32_videocontroller | select pscomputername,name,videomodedescription
      
      foreach ($display in $displays) {
         $w=$display.PelsWidth
         $h=$display.PelsHeight
         "$pc Width: $w Height: $h"
         "$pc Width: $w Height: $h" | Out-File -FilePath  C:\LazyWinAdmin\Backgrounds\DisplayResolution.txt -Append
      }

   } else {
            "$pc : Not Responding"
            "$pc : Not Responding" | Out-File -FilePath  C:\LazyWinAdmin\Backgrounds\DisplayResolution.txt -Append
          }
}