#############################################################################
##           Script to report white space in Exchange 2010 databases                     
##           Author: Ankush Sharma
##	     Reviewer: Vikas Sukhija                  		 
##           Date: 05-08-2014                       		 
##
#############################################################################

If ((Get-PSSnapin | where {$_.Name -match "Microsoft.Exchange.Management.PowerShell.E2010"}) -eq $null)
{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
}


##############################################################################

$whitespace  = ".\whitespace.htm"

$checkrep = Test-Path ".\whitespace.htm" 

If ($checkrep -like "True")

{

Remove-Item ".\whitespace.htm"


}

New-Item ".\whitespace.htm" -type file

$smtphost = "SMTPSERVER" 
$from = "white_space_report@labtest.com" 
$to = "vikassukhija@labtest.com"
################################ADD HTML Content#############################



Add-Content $whitespace "<html>" 
Add-Content $whitespace "<head>" 
Add-Content $whitespace "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" 
Add-Content $whitespace '<title>Whitespace Report</title>' 
add-content $whitespace '<STYLE TYPE="text/css">' 
add-content $whitespace  "<!--" 
add-content $whitespace  "td {" 
add-content $whitespace  "font-family: Tahoma;" 
add-content $whitespace  "font-size: 11px;" 
add-content $whitespace  "border-top: 1px solid #999999;" 
add-content $whitespace  "border-right: 1px solid #999999;" 
add-content $whitespace  "border-bottom: 1px solid #999999;" 
add-content $whitespace  "border-left: 1px solid #999999;" 
add-content $whitespace  "padding-top: 0px;" 
add-content $whitespace  "padding-right: 0px;" 
add-content $whitespace  "padding-bottom: 0px;" 
add-content $whitespace  "padding-left: 0px;" 
add-content $whitespace  "}" 
add-content $whitespace  "body {" 
add-content $whitespace  "margin-left: 5px;" 
add-content $whitespace  "margin-top: 5px;" 
add-content $whitespace  "margin-right: 0px;" 
add-content $whitespace  "margin-bottom: 10px;" 
add-content $whitespace  "" 
add-content $whitespace  "table {" 
add-content $whitespace  "border: thin solid #000000;" 
add-content $whitespace  "}" 
add-content $whitespace  "-->" 
add-content $whitespace  "</style>" 
Add-Content $whitespace  "</head>" 
Add-Content $whitespace  "<body>" 
add-content $whitespace  "<table width='100%'>" 
add-content $whitespace  "<tr bgcolor='Aliceblue'>" 
add-content $whitespace  "<td colspan='7' height='25' align='center'>" 
add-content $whitespace  "<font face='tahoma' color='#003499' size='4'><strong>Whitespace Report</strong></font>" 
add-content $whitespace  "</td>" 
add-content $whitespace  "</tr>" 
add-content $whitespace  "</table>" 
 
add-content $whitespace  "<table width='100%'>" 
Add-Content $whitespace "<tr bgcolor='BlanchedAlmond'>" 
Add-Content $whitespace  "<td width='10%' align='center'><B>Database Name</B></td>" 
Add-Content $whitespace  "<td width='10%' align='center'><B>AvailableNewMailboxSpace</B></td>" 
Add-Content $whitespace "</tr>" 



########################################################################################################

################################## White Space #################################################

$databases = Get-MailboxDatabase -Status | sort -Descending AvailableNewMailboxSpace
foreach($Database in $databases)
{
  
      
         Write-Host $Database.Name `t $Database.Server `t $Database.AvailableNewMailboxSpace -ForegroundColor Green 
         
         $machineName = $Database.Name 
         $svcState = $Database.AvailableNewMailboxSpace

         Add-Content $whitespace "<tr>" 
         Add-Content $whitespace "<td bgcolor= 'Lavendar' align=center>  <B> $machineName</B></td>" 
         Add-Content $whitespace "<td bgcolor= 'Aquamarine' align=center><B>$svcState</B></td>" 
         Add-Content $whitespace "</tr>" 
              
}




############################################Close HTMl Tables#########################################


Add-content $whitespace  "</table>" 
Add-Content $whitespace "</body>" 
Add-Content $whitespace "</html>" 



#####################################################################################################
#############################################Send Email##############################################
$subject = "WhiteSpace Report" 
$body = Get-Content ".\whitespace.htm" 
$smtp= New-Object System.Net.Mail.SmtpClient $smtphost 
$msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body 
$msg.isBodyhtml = $true 
$smtp.send($msg) 

#####################################################################################################