#############################################################################
#       Author: Mahesh Sharma
#       Reviewer: Vikas SUkhija      
#       Date: 06/10/2013
#	Modified:06/19/2013 - made it to run from any path
#       Description: ExChange Health Status
#############################################################################

########################### Add Exchange Shell##############################

If ((Get-PSSnapin | where {$_.Name -match "Exchange.Management"}) -eq $null)
{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin
}

###########################Define Variables##################################

$reportpath = "\\networkshare\folder\CMSReport.htm" 
$smtphost = "smtp server" 
$from = "ExchangeStatus@labtest.com" 
$to = "Vikas.sukhija@labtest.com"

###############################HTml Report Content############################
$report = $reportpath

Clear-Content $report 
Add-Content $report "<html>" 
Add-Content $report "<head>" 
Add-Content $report "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" 
Add-Content $report '<title>Exchange Status Report</title>' 
add-content $report '<STYLE TYPE="text/css">' 
add-content $report  "<!--" 
add-content $report  "td {" 
add-content $report  "font-family: Tahoma;" 
add-content $report  "font-size: 11px;" 
add-content $report  "border-top: 1px solid #999999;" 
add-content $report  "border-right: 1px solid #999999;" 
add-content $report  "border-bottom: 1px solid #999999;" 
add-content $report  "border-left: 1px solid #999999;" 
add-content $report  "padding-top: 0px;" 
add-content $report  "padding-right: 0px;" 
add-content $report  "padding-bottom: 0px;" 
add-content $report  "padding-left: 0px;" 
add-content $report  "}" 
add-content $report  "body {" 
add-content $report  "margin-left: 5px;" 
add-content $report  "margin-top: 5px;" 
add-content $report  "margin-right: 0px;" 
add-content $report  "margin-bottom: 10px;" 
add-content $report  "" 
add-content $report  "table {" 
add-content $report  "border: thin solid #000000;" 
add-content $report  "}" 
add-content $report  "-->" 
add-content $report  "</style>" 
Add-Content $report "</head>" 
Add-Content $report "<body>" 
add-content $report  "<table width='100%'>" 
add-content $report  "<tr bgcolor='Lavender'>" 
add-content $report  "<td colspan='7' height='25' align='center'>" 
add-content $report  "<font face='tahoma' color='#003399' size='4'><strong>CMS Status Report</strong></font>" 
add-content $report  "</td>" 
add-content $report  "</tr>" 
add-content $report  "</table>" 
 
add-content $report  "<table width='100%'>" 
Add-Content $report  "<tr bgcolor='IndianRed'>" 
Add-Content $report  "<td width='10%' align='center'><B>Identity</B></td>" 
Add-Content $report  "<td width='5%' align='center'><B>State</B></td>" 
Add-Content $report  "<td width='20%' align='center'><B>OperationalMachines</B></td>" 
Add-Content $report  "<td width='15%' align='center'><B>FailedResources</B></td>" 
Add-Content $report  "<td width='15%' align='center'><B>FailedReplicationHostNames</B></td>" 
 

Add-Content $report "</tr>" 

##############################Get ALL CCR's##################################

$inputCMS = Get-ExchangeServer | where {$_.IsMemberOfCluster -eq "Yes"} | select name

######################### Create UI Button for userid & Password##############

[void][System.Reflection.Assembly]::LoadWithPartialName( 
    "System.Windows.Forms") 
[void][System.Reflection.Assembly]::LoadWithPartialName( 
    "Microsoft.VisualBasic") 
     
$Form = New-Object System.Windows.Forms.Form 
$Button = New-Object System.Windows.Forms.Button 
$TextBox1 = New-Object System.Windows.Forms.TextBox 
$TextBox2 = New-Object System.Windows.Forms.TextBox 
$Label1 = New-Object System.Windows.Forms.Label 
$Label2 = New-Object System.Windows.Forms.Label 

 
$Form.Text = "Enter Credentials" 
$Form.StartPosition =  
    [System.Windows.Forms.FormStartPosition]::CenterScreen 
 

$Label1.Text = "Domain\User Name" 
$Label1.Top = 50 
$Label1.Left = 25 


$Label2.Text = "Password" 
$Label2.Top = 90 
$Label2.Left = 25 



 
$TextBox1.Text = "" 
$TextBox1.Top = 50 
$TextBox1.Left = 150 

$Textbox2.Passwordchar = "*"
$TextBox2.Text = "" 
$TextBox2.Top = 90 
$TextBox2.Left = 150 

$Button.Text = "Run!!" 
$Button.Top = 130 
$Button.Left = 115 
$Button.Width = 70 

 

 
$Button_Click =  
{ 
        $User = $TextBox1.Text
	$Password = $TextBox2.Text
  	$form.close()  
} 
 
$Form.Controls.Add($Label2) 
$Form.Controls.Add($Label1) 
$Form.Controls.Add($Button) 
$Form.Controls.Add($TextBox1) 
$Form.Controls.Add($TextBox2) 
$Button.Add_Click($Button_Click) 
 
$Form.ShowDialog()

#####################################Active Cluster Nodes###############################################

foreach($node in $inputCMS) {

$anode = $node.name


$anode = Get-WmiObject win32_ComputerSystem -ComputerName $anode

$actnode = $anode.name

$actnode

$tempfolder = Test-Path \\$actnode\c$\temp

Write-host "$actnode tempfolder status $tempfolder"

if($tempfolder -like "False") {

Write-host "Temp folder doesnt exist on $actnode, creating Temp folder"

New-Item -ItemType directory -Path \\$actnode\c$\temp

}

else {

Write-host "temp folder exists on $actnode"

}


$scstatus = Test-Path \\$actnode\c$\temp\Get-CMS-Status.ps1

if($scstatus -like "False") {

Write-host "Get-CMS-Status script doesnt exist on $actnode, copying it to Temp folder"

Copy-Item .\CMSScripts\Get-CMS-Status.ps1 \\$actnode\c$\temp

}

else {

Write-host "Get-CMS-Status scriptexists on $actnode"

}


.\psexec.exe \\$actnode -u $user -p $password   cmd /c "echo . | powershell c:\temp\get-cms-status.ps1"

}

add-content $report  "</table>" 

################################################################################################################
################################################################################################################


$CMSList = $inputCMS
$TestCMSMailFlow =  $inputCMS

$report = $reportpath
 
Add-Content $report "<html>" 
Add-Content $report "<head>" 
Add-Content $report "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" 
Add-Content $report '<title>Exchange Status Report</title>' 
add-content $report '<STYLE TYPE="text/css">' 
add-content $report  "<!--" 
add-content $report  "td {" 
add-content $report  "font-family: Tahoma;" 
add-content $report  "font-size: 11px;" 
add-content $report  "border-top: 1px solid #999999;" 
add-content $report  "border-right: 1px solid #999999;" 
add-content $report  "border-bottom: 1px solid #999999;" 
add-content $report  "border-left: 1px solid #999999;" 
add-content $report  "padding-top: 0px;" 
add-content $report  "padding-right: 0px;" 
add-content $report  "padding-bottom: 0px;" 
add-content $report  "padding-left: 0px;" 
add-content $report  "}" 
add-content $report  "body {" 
add-content $report  "margin-left: 5px;" 
add-content $report  "margin-top: 5px;" 
add-content $report  "margin-right: 0px;" 
add-content $report  "margin-bottom: 10px;" 
add-content $report  "" 
add-content $report  "table {" 
add-content $report  "border: thin solid #000000;" 
add-content $report  "}" 
add-content $report  "-->" 
add-content $report  "</style>" 
Add-Content $report "</head>" 
Add-Content $report "<body>" 
add-content $report  "<table width='100%'>" 
add-content $report  "<tr bgcolor='Lavender'>" 
add-content $report  "<td colspan='7' height='25' align='center'>" 
add-content $report  "<font face='tahoma' color='#003399' size='4'><strong>Storage Group Status Report</strong></font>" 
add-content $report  "</td>" 
add-content $report  "</tr>" 
add-content $report  "</table>" 
 
add-content $report  "<table width='100%'>" 
Add-Content $report "<tr bgcolor='IndianRed'>" 
Add-Content $report  "<td width='25%' align='center'><B>Identity</B></td>" 
Add-Content $report "<td width='25%' align='center'><B>Copy Status</B></td>" 
Add-Content $report  "<td width='25%' align='center'><B>Copy Queue Lenght</B></td>" 
Add-Content $report  "<td width='25%' align='center'><B>Reply Queue Lenght</B></td>" 
Add-Content $report "</tr>" 


##########################################################################################################
############################################## SG Copy Status ############################################


foreach ($CMSName in $CMSlist) 
{ 
 

$CMSName = $CMSName.Name
		
		$FullStatus = Get-StorageGroupCopyStatus -Identity $CMSName\*

		Foreach ($status in $Fullstatus)
		{

			if ($status.SummaryCopyStatus -eq "Healthy")
			{
			$Identity = $status.identity
			$Copystatus =  $status.SummaryCopyStatus 
			$copylength =  $status.copyqueuelength
			$replylength = $status.ReplayQueueLength
			Add-Content $report "<tr>" 
			Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B> $Identity</B></td>" 
         		Add-Content $report "<td bgcolor= 'Aquamarine' align=center>  <B>$copystatus</B></td>" 
			Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$copylength</B></td>" 
			Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$replylength</B></td>" 
			Add-Content $report "</tr>" 
			}

			else 
			{
			$Identity = $status.identity
			$Copystatus =  $status.SummaryCopyStatus 
			$copylength =  $status.copyqueuelength
			$replylength = $status.ReplayQueueLength
			Add-Content $report "<tr>" 
			Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B> $Identity</B></td>" 
         		Add-Content $report "<td bgcolor= 'Red' align=center>  <B>$copystatus</B></td>" 
			Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$copylength</B></td>" 
			Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$replylength</B></td>" 
			Add-Content $report "</tr>" 
			}

		}
        	
		
}


##################################################################################################################
############################################## Test mail Flow For CMS ############################################


add-content $report  "<tr bgcolor='Lavender'>" 
add-content $report  "<td colspan='7' height='25' align='center'>" 
add-content $report  "<font face='tahoma' color='#003399' size='4'><strong>Mail Flow Test Report</strong></font>" 
add-content $report  "</td>" 
add-content $report  "</tr>"

add-content $report  "</tr>" 
add-content $report  "</table>" 
add-content $report  "<table width='100%'>" 
Add-Content $report "<tr bgcolor='IndianRed'>"
Add-Content $report  "<td width='25%' align='center'><B>Result</B></td>" 
Add-Content $report "<td width='25%' align='center'><B>Message Latency Time</B></td>" 
Add-Content $report  "<td width='25%' align='center'><B>IsRemoteTest</B></td>" 
Add-Content $report "</tr>" 


Foreach ($CMS in $TestCMSMailFlow)
{

$flow = Test-MailFlow -Identity $CMS.Name

$result = $flow.TestMailflowResult
$time = $Flow.MessageLatencyTime
$remote =  $Flow.IsRemoteTest
Add-Content $report "<tr>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B> $result</B></td>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$time</B></td>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$remote</B></td>" 

Add-Content $report "</tr>"

}



#####################################################################################################################
############################################## Get Queue For HUB Servers ############################################


add-content $report  "<tr bgcolor='Lavender'>" 
add-content $report  "<td colspan='7' height='25' align='center'>" 
add-content $report  "<font face='tahoma' color='#003399' size='4'><strong>Mail Queue Status</strong></font>" 
add-content $report  "</td>" 
add-content $report  "</tr>"

add-content $report  "</tr>" 
add-content $report  "</table>" 
add-content $report  "<table width='100%'>" 
Add-Content $report "<tr bgcolor='IndianRed'>"
Add-Content $report  "<td width='25%' align='center'><B>Identity</B></td>" 
Add-Content $report "<td width='25%' align='center'><B>Delivery Type</B></td>" 
Add-Content $report  "<td width='25%' align='center'><B>Status</B></td>" 
Add-Content $report "<td width='25%' align='center'><B>Message Count</B></td>" 
Add-Content $report  "<td width='25%' align='center'><B>Next Hop Domain</B></td>"
Add-Content $report "</tr>" 


Foreach ($Queue in $GetHub = Get-TransportServer | get-Queue)
{

$Identity = $Queue.Identity
$DeliveryType = $Queue.DeliveryType
$Status = $Queue.Status
$MSgCount =  $Queue.Messagecount
$NextHopDomain = $Queue.NextHopDomain


Add-Content $report "<tr>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B> $Identity</B></td>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$DeliveryType</B></td>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$Status</B></td>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$MSgCount</B></td>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$NextHopDomain</B></td>" 

Add-Content $report "</tr>"


}

###########################################################################################################################
######################################################### Send Mail #######################################################


Add-content $report  "</table>" 
Add-Content $report "</body>" 
Add-Content $report "</html>"

 
$subject = "Exchange Status Check Report" 
$body = Get-Content $reportpath
$smtp= New-Object System.Net.Mail.SmtpClient $smtphost 
$msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body 
$msg.isBodyhtml = $true 
$smtp.send($msg) 

###################################################Exchange Test Complete##################################################


