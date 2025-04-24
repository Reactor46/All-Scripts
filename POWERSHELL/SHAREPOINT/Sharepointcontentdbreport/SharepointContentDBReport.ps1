#***************************************************************************
#    Script to Get Content Db report for Different Farms                    * 
#    Date   : 7th August,2104                                              *
#    Author : Abhishek Gupta                                               *
#    Reviewer: Vikas Sukhija                                               *
#    modified: Added logging, error checking & converted to function       *
#***************************************************************************
$date = get-date -format d
# replace \ by -
$time = get-date -format t
$month = get-date 
$month1 = $month.month
$year1 = $month.year

$date = $date.ToString().Replace(“/”, “-”)

$time = $time.ToString().Replace(":", "-")
$time = $time.ToString().Replace(" ", "")

$log1 = ".\Processed\Logs" + "\" + "skipcsv_" + $date + "_.log"
#$log2 = ".\Processed\Logs" + "\" + "Modified_" + $month1 +"_" + $year1 +"_.log"
#$output1 = ".\" + "G_DistributionList_" + $date + "_" + $time + "_.csv" 

$logs = ".\Processed\Logs" + "\" + "Powershell" + $date + "_" + $time + "_.txt"

Start-Transcript -Path $logs 

# ***************************************************************************
# Variable initializing to send mail
$TXTFile = ".\ContentDBReport.html"
$SMTPServer = "smtp.lab.com" 
$emailFrom = "Messaging@lab.com" 
$emailTo = "vikassukhija@lab.com" 
$subject = "Sharepoint Farms Content databases Report" 
$emailBody = "Dailyreport on Sharepoint Farms Content databases"

#****************************************************************************
# HTML code to format output
$b = "<style>"
$b = $b + "BODY{background-color:white;}"
$b = $b + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$b = $b + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}"
$b = $b + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;}"
$b = $b + "</style>"

#********************************************************************************
# Creating PSSession and Loading Snapin(make sure your account has rights to sharepoint)

$cred1 = Get-Credential

$s1 = New-PSSession -ComputerName Server1 -Authentication CredSSP -Credential $cred1
$s2 = New-PSSession -ComputerName Server2 -Authentication CredSSP -Credential $cred1
$s3 = New-PSSession -ComputerName Server3 -Authentication CredSSP -Credential $cred1
$s4 = New-PSSession -ComputerName server4 -Authentication CredSSP -Credential $cred1

Function ContentReport ($session){

Invoke-Command -Session $session -ScriptBlock {Add-PSSnapin Microsoft.SharePoint.PowerShell}

$f1 = Invoke-Command -Session $session -ScriptBlock {Get-SPWebApplication | Get-SPContentDatabase}
$g1 = $f1 | Select-Object DisplayName,WebApplication,CurrentSiteCount,disksizerequired,WarningSiteCount,MaximumSiteCount | 
      ConvertTo-Html -Fragment DisplayName,WebApplication,WarningSiteCount,MaximumSiteCount,@{L='DiskSizeRequired';E={
                                          if($_.disksizerequired -gt 100Gb){
                                               "#font"+$_.disksizerequired+"font#"
                                          }elseif($_.disksizerequired -ge 80Gb){
                                                "#blue"+$_.disksizerequired+"blue#"}
                                          else{
                                               $_.disksizerequired
                                          }
                                       }
                                 },@{L='CurrentSiteCount';E={
                                 if($_.CurrentSiteCount -ge $_.WarningSiteCount){ 
                                 "#blue"+$_.CurrentSiteCount+"blue#"
                                 }else{
                                               $_.CurrentSiteCount
                                          }
                                          }
                                          }                            
$g1 = $g1 -replace ("#font",'<span style="color:black; background-color:red" >')
$g1 = $g1 -replace "font#","</span>"
$g1 = $g1 -replace ('#blue','<span style="color:black; background-color:orange"> ')
$g1 = $g1 -replace "blue#","</span>"

return $g1
}
###############call function for diffrent farms###################################################

$h1 = ContentReport $s1
$h2 = ContentReport $s2
$h3 = ContentReport $s3
$h4 = ContentReport $s4

##############################Convert to HTML ####################################################

ConvertTo-HTML -head $b -Body "<h1>$(Get-Date) Sharepoint Farm Database Content DB Report</h1> <br /> 
<h2>SharePoint_ParentFarm $h1 SharePoint_IntranetFarm  $h2 SharePoint_GalwayFarm  $h3 SharePoint_TokyoFarm  $h4</h2>" | 
Out-File $TXTFile

# Code to Send Mail 
Send-MailMessage -SmtpServer $SMTPServer -From $emailFrom -To $emailTo -Subject $subject -Body $emailBody -Attachment $TXTFile

if ($error -ne $null)
      {
#SMTP Relay address
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)

#Mail sender
$msg.From = $emailFrom
#mail recipient
$msg.To.Add($emailTo)
#$msg.To.Add($email2)
$msg.Subject = "Sharepoint Content DB Script error"
$msg.Body = $error
$smtp.Send($msg)
$error.clear()
       }
  else

      {
    Write-host "no errors till now"
      }

Stop-Transcript
#*************************************************************************************************
