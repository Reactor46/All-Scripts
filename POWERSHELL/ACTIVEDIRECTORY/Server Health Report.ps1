
$ServerListFilePath = "C:\inetpub\AdminScripts\Configs\Servers.txt"  
$ServerList = Get-Content $ServerListFilePath -ErrorAction SilentlyContinue
$ReportFilePath = "C:\Test\Report.htm" 
$Result = @()
 
ForEach($ComputerName in $ServerList) 
{
$AVGProc = Get-WmiObject -computername $ComputerName win32_processor | 
Measure-Object -property LoadPercentage -Average | Select Average
$OS = gwmi -Class win32_operatingsystem -computername $ComputerName |
Select-Object @{Name = "MemoryUsage"; Expression = {“{0:N2}” -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory)*100)/ $_.TotalVisibleMemorySize) }}
$vol = Get-WmiObject -Class win32_Volume -ComputerName $ComputerName -Filter "DriveLetter = 'C:'" |
Select-object @{Name = "C PercentUsed"; Expression = {“{0:N2}” -f  ((($_.Capacity - $_.FreeSpace ) / $_.Capacity)*100) } }
  
$Result += [PSCustomObject] @{ 
        ServerName = "$ComputerName"
        CPULoad = $AVGProc.Average
        MemLoad = $OS.MemoryUsage
        CDrive = $vol.'C PercentUsed'
    }

    $OutputReport = "<HTML><TITLE>Server Health Report</TITLE>
                     <BODY>
                     <font color =""#99000"">
                     <H2>Daily Server Health Report</H2></font>
                     <Table border=2 cellpadding=4 cellspacing=3>
                     <TR bgcolor=D1D0CE align=center>
                       <TD><B>Server</B></TD>
                       <TD><B>Avg.CPU Usage</B></TD>
                       <TD><B>Memory Usage</B></TD>
                       <TD><B>C: Drive Usage</B></TD></TR>"
                        
    Foreach($Entry in $Result) 
    { 
          #convert raw data to percentages
          $CPUAsPercent = "$($Entry.CPULoad)%"
          $MemAsPercent = "$($Entry.MemLoad)%"
          $CDriveAsPercent = "$($Entry.CDrive)%"

          $OutputReport += "<TR><TD>$($Entry.Servername)</TD>"

          # check CPU load
          if(($Entry.CPULoad) -ge 80) 
          {
              $OutputReport += "<TD bgcolor=E41B17 align=center>$($CPUAsPercent)</TD>"
          } 
          elseif((($Entry.CPULoad) -ge 70) -and (($Entry.CPULoad) -lt 80))
          {
              $OutputReport += "<TD bgcolor=yellow align=center>$($CPUAsPercent)</TD>"
          }
          else
          {
              $OutputReport += "<TD bgcolor=lightgreen align=center>$($CPUAsPercent)</TD>" 
          }

          # check RAM load
          if(($Entry.MemLoad) -ge 80)
          {
              $OutputReport += "<TD bgcolor=E41B17 align=center>$($MemAsPercent)</TD>"
          }
          elseif((($Entry.MemLoad) -ge 70) -and (($Entry.MemLoad) -lt 80))
          {
              $OutputReport += "<TD bgcolor=yellow align=center>$($MemAsPercent)</TD>"
          }
          else
          {
              $OutputReport += "<TD bgcolor=lightgreen align=center>$($MemAsPercent)</TD>"
          }

          # check C: Drive Usage
          if(($Entry.CDrive) -ge 80)
          {
              $OutputReport += "<TD bgcolor=E41B17 align=center>$($CDriveAsPercent)</TD>"
          }
          elseif((($Entry.CDrive) -ge 70) -and (($Entry.CDrive) -lt 80))
          {
              $OutputReport += "<TD bgcolor=yellow align=center>$($CDriveAsPercent)</TD>"
          }
          else
          {
              $OutputReport += "<TD bgcolor=lightgreen align=center>$($CDriveAsPercent)</TD>"
          }

          $OutputReport += "</TR>"
    }

    $OutputReport += "</Table></BODY></HTML>" 
} 
 
$OutputReport | out-file $ReportFilePath
Invoke-Expression $ReportFilePath

#Send email functionality from below line, use it if you want   
$EmailTo = "receiver@fakeemail.com"
$EmailFrom = "sender@fakeemail.com"
$EmailPW = "Pa55w0rd"
$Subject = "Server Health Report"
$Body = ""
$SMTPServer = "smtp.office365.com" 
$SMTPPort = 587
$SMTPMessage = New-Object System.Net.Mail.MailMessage($EmailFrom,$EmailTo,$Subject,$Body)
$SMTPMessage.IsBodyHTML = $true
$SMTPMessage.Body = "<head><pre>$style</pre></head>"
$SMTPMessage.Body += Get-Content $ReportFilePath
$SMTPClient = New-Object Net.Mail.SmtpClient($SMTPServer, $SMTPPort) 
$SMTPClient.EnableSsl = $true
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($EmailFrom, $EmailPW); 
$SMTPClient.Send($SMTPMessage)

Remove-Item $ReportFilePath