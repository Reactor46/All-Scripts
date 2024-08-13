###############################################################################################
# 
# Generate Disk Space Report and send email.
# If you dont want the email function, then it will just generate a html file
# Script will give you a warnings based on the threshold set.
# We can use this script to pull a Disk Space report for all servers.
# We need to give server names as an input file. This is no required for Exchange Servers.
# How to Run the script
# 1. Create Input list with server names and save the file as "Servers.txt"
# 2. If you want to send the report as email, Give $Mail = "True" else give $Mail = "False"
# 3. Move PS1 and input file into same directory and run the script
# Date: 03/26/2015
# Not added any error checking into this script yet.
# 
###############################################################################################

# Input

#$Servers = Get-ExchangeServer | sort name | %{$_.name}
$Servers = GC .\server.txt
$Mail = "False"

# Threshold

$percentWarning = 20;
$percentCritcal = 10;
$all = 0;

# Inputs for Email

$Subject = "Disk Space Report"
$FromAddress = "DiskSpaceReport@creditone.com"
$ToAddress = "john.battista@creditone.com"
$Relay = "mailgateway.Contoso.corp" 
$SMTPClient = New-Object System.Net.Mail.smtpClient
$MailMessage = New-Object System.Net.Mail.MailMessage

#############################################################################

#Add-PSSnapin Microsoft.Exchange.Management.Powershell.Admin
#$AdminSessionADSettings.ViewEntireForest = $true

$MailSend = $null
$body = $null
$body += '<style>'
$body += '<!-- '
$body += 'body { font: Calibri } '
$body += 'table { font: 11pt Calibri; border: 1px} '
$body += 'td { border: 1px ridge white; padding-left: 1px; padding-right: 1px } '
$body += 'p.header { font: 11pt Verdana; color: DarkBlue }'
$body += 'p.normal { font: 8pt Calibri; color: CadetBlue }'
$body += '-->'
$body += '</style>'
$body += '</head>'
$body += '<body>'
$body += "<p class=`"header`" align='center'><strong>Disk Space Report : " + $(get-date).toshortdatestring() + " @ " + $(get-date).ToshorttimeString()

foreach($Server in $Servers)
{
    $Disks = Get-WmiObject -ComputerName $Server -Class Win32_Volume -Filter "FileSystem = 'NTFS'" |Sort Driveletter,Capacity | Select DriveLetter,Name,Caption,Label,Capacity,FreeSpace,FileSystem,SystemName
    $Server = $Server.toupper()
    $Serverprint = "Yes"

    foreach($Disk in $Disks)
    {
        $Label = $Disk.Label
        $Name = $Disk.Name
        [float]$Size = $Disk.Capacity;
        [float]$Freespace = $Disk.FreeSpace;
        $PercentFree = [Math]::Round(($Freespace / $Size) * 100, 2);
        $SizeGB = [Math]::Round($Size / 1073741824, 2);
        $FreeSpaceGB = [Math]::Round($Freespace / 1073741824, 2);
        $UsedSpaceGB = [Math]::Round($SizeGB – $FreeSpaceGB, 2);
        $DriveLetter = $Disk.Driveletter
        If (($PercentFree -ge $all))
        {
	        while ($Serverprint -eq "Yes")
		    {
                #Table Header

		        $body += '<table width=80% align=center>'
                $body += '<td colspan=7 align=center style="background-color: #DCDCDC" border: 1 ; ><strong><font color= Black>' + $Server.toupper()  + '</font></strong></td></tr>'

                #column Header

		        $body += '<tr><td align="Center" style="background-color: #585858" width="13%" border: 1 Groove white;><p class="table" style="color:    white"><b>Drive</b></p></td>'
                $body += '<td align="Center" style="background-color: #585858" width="14%" border: 1 Groove white;><p class="table" style="color:    white"><b>Drive Label</b></p></td>'
                $body += '<td align="Center" style="background-color: #585858" width="16%" border: 1 Groove white;><p class="table" style="color:    white"><b>Total Capacity (GB)</b></p></td>'
                $body += '<td align="Center" style="background-color: #585858" width="16%" border: 1 Groove white;><p class="table" style="color:    white"><b>Used Capacity (GB)</b></p></td>'
                $body += '<td align="Center" style="background-color: #585858" width="14%" border: 1 Groove white;><p class="table" style="color:    white"><b>Free Space (GB)</b></p></td>'
		        $body += '<td align="Center" style="background-color: #585858" width="14%" border: 1 Groove white;><p class="table" style="color:    white"><b>Free Space %</b></p></td>'
		        $body += '<td align="Center" style="background-color: #585858" width="13%" border: 1 Groove white;><p class="table" style="color:    white"><b>Status</b></p></td></tr>'
		        $Serverprint = "No"
		    }		
		    If ($DriveLetter -eq $null)
			{
			    $SplitDl = $Name.split("\")
			    $DriveLetter = $SplitDl[0]
			}	
		        $body += '<tr><td align="Center" border: 1 Groove white;>' + $DriveLetter + "</td>"
		        $body += '<td align="Center" border: 1 Groove white;>' + $Label + "</td>"
		        $body += '<td align="Center" border: 1 Groove white;>' + $SizeGB + "</td>"
		        $body += '<td align="Center" border: 1 Groove white;>' + $usedSpaceGB + "</td>"
		        $body += '<td align="Center" border: 1 Groove white;>' + $freeSpaceGB + "</td>"
		        $body += '<td align="Center" border: 1 Groove white;>' + $percentFree + "</td>"
		    If ($percentFree -le $percentCritcal )
            {
		        $body += '<td align="Center" border: 1 Groove white; bgcolor="FF6347">' + "Critical" + "</td></tr>"
            }
            elseif (($percentFree -le $percentWarning ) -and ($percentFree -ge $percentCritcal ))
            {
		        $body += '<td align="Center" border: 1 Groove white; bgcolor="FFCC66">' + "Warning" + "</td></tr>"
            }
		    else 
            {
		        $body += '<td align="Center" border: 1 Groove white; bgcolor="A3E6AE">' + "Ok" + "</td></tr>" 
            }
        
		}

    }
	$body += '</table><br/>'	
}
if ($Mail -eq "True")
{
    $body += "Report Completed"
    ##############################################################################
    $MailMessage.Subject = "$Subject"
    $MailMessage.Body = $body
    $MailMessage.sender = "$FromAddress"
    $MailMessage.From = "$FromAddress"
    $MailMessage.To.add("$ToAddress")
    #$MailMessage.CC.add("$CCAddress")
    #$MailMessage.CC.add("$CCAddress2")
    $MailMessage.Priority = [System.Net.Mail.MailPriority]::High
    $MailMessage.IsBodyHTML = $True
    $smtpclient.host = "$Relay"
    $smtpclient.send($MailMessage)
    #############################################################################
}
else
{
    $Date = get-date
    $FileName = 'Disk_Space_Info'
    $FileName += $Date.year.tostring()
    $FileName += $Date.month.tostring()
    $FileName += $Date.day.tostring()
    $FileName += '.html'
    $body += "Report Completed"    
    set-content $FileName -value $body
 
}