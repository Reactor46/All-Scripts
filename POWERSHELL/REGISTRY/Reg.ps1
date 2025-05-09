##########################################################################
#       Author:Aishwarya Rawat
#       Reviewer: Vikas Sukhija
#       Date: 12/26/2013
#       Description: Check for Registry Value.
##########################################################################

##################Define variables########################################

$main = "HKLM"
$Path = "\Software\Microsoft\Windows\CurrentVersion\WindowsUpdate\"
$key = "SusClientId"

####################Define email Variables#################################

$smtphost = "mailgateway.fnbm.corp" 
$from = "Registry_value@creditone.com" 
$to = "john.battista.creditone.com" 

############################################################################

$report = "C:\LazyWinAdmin\ReadRegistry\SusClientId-Report.htm" 
$servers = Get-content "C:\LazyWinAdmin\lasdmzweb.txt"

$checkrep = Test-Path $report 

If ($checkrep -like "True")

{

Remove-Item $report


}

New-Item $report -type file

################################ADD HTML Content#############################


Add-Content $report "<html>" 
Add-Content $report "<head>" 
Add-Content $report "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" 
Add-Content $report '<title>Registry Value $Key Status</title>' 
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
add-content $report  "<table width='50%'>" 
add-content $report  "<tr bgcolor='Lavender'>" 
add-content $report  "<td colspan='7' height='25' align='center'>" 

#######################Title of table####################################################

add-content $report  "<font face='tahoma' color='#003399' size='4'><strong>Registry Value $Key Status</strong></font>" 
add-content $report  "</td>" 
add-content $report  "</tr>" 
add-content $report  "</table>" 

######################Definae Columns###################################################
add-content $report  "<table width='50%'>" 
Add-Content $report "<tr bgcolor='Lavender'>" 
Add-Content $report  "<td width='10%' align='center'><B>Server Name</B></td>" 
Add-Content $report  "<td width='10%' align='center'><B>Value</B></td>" 
Add-Content $report "</tr>" 


#####Get Registry Value ####

foreach ($Server in $servers) 
{

$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($main, $Server)
$regKey= $reg.OpenSubKey($path)
$Value = $regkey.GetValue($key)
  
######################Add values inside Columns########################################  
         Add-Content $report "<tr>" 
	     Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B> $Server </B></td>" 
         Add-Content $report "<td bgcolor= 'Aquamarine' align=center><B>$Value</B></td>" 
         Add-Content $report "</tr>" 
             
                                                
   
}


#####################Close HTMl Tables###############################################


Add-content $report  "</table>" 
Add-Content $report "</body>" 
Add-Content $report "</html>" 

#####Send Email#####


#$subject = "Registry Value $Key Status" 
#$body = Get-Content ".\report.htm" 
#$smtp= New-Object System.Net.Mail.SmtpClient $smtphost 
#$msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body 
#$msg.isBodyhtml = $true 
#$smtp.send($msg) 

#############################################################################################
 
