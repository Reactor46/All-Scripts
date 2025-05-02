##################################################################################
#       Author: Kaustav Samaddar
#       Reviewer : Vikas SUkhija
#       Date: 04/05/2013
#       Modified:01/28/2014
#       Modificetion: HTML, enviornment independent..
#       Description: White space report
##################################################################################
##########################Add exchange Shell######################################

If ((Get-PSSnapin | where {$_.Name -match "Exchange.Management"}) -eq $null)
{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin
}

#################################################################################

$db=get-mailboxdatabase | measure-object 

$i=0

$date=(get-date).adddays(-7)

$server_list=Get-ExchangeServer | where{($_.ServerRole -like "*mailbox*") -and ($_.AdminDisplayVersion -like "*8.3*")}
$server_list

$white_space = @()

################################## loop thru server list..

$server_list | foreach-object {

	$sname = $_
       write-host  "processing Server $sname ................MBX DB" -ForegroundColor green
        

	$data = @()

	$record = get-wmiobject -computer $sname -class Win32_NTLogEvent -Filter "logfile = 'application' AND EventIdentifier = 1074136261 AND sourcename = 'MSExchangeIS Mailbox Store'"| select -First $db.count

		$record | Foreach-object {

			$row= "" | select Server,Database,WhitespaceinMB,Timegenerated

			$row.Server=$sname

			$row.Timegenerated=Get-Date([System.Management.ManagementDateTimeconverter]::ToDateTime($_.TimeGenerated))

			$row.Database=$_.insertionstrings[1]

			$row.WhitespaceinMB=$_.insertionstrings[0]

			if($row.timegenerated -ge $date){$data+=$row}

					}

	$pf=Get-PublicFolderDatabase -server $sname

        if($pf -ne $null){

        write-host  "processing Server $_ ................PF DB" -ForegroundColor green

	$record = get-wmiobject -computer $sname -class Win32_NTLogEvent -Filter "logfile = 'application' AND EventIdentifier = 1074136261 AND sourcename = 'MSExchangeIS public Store'"| select -First $db.count| Sort-Object -Property Message

		if($record -ne $null) {

			$record | Foreach-object {

			$row= "" | select Server,Database,WhitespaceinMB,Timegenerated

			$row.Server=$sname

			$row.Timegenerated=Get-Date([System.Management.ManagementDateTimeconverter]::ToDateTime($_.TimeGenerated))

			$row.Database=$_.insertionstrings[1]

			$row.WhitespaceinMB=$_.insertionstrings[0]

			if($row.timegenerated -ge $date){$data+=$row}

						  }
				      }
                           }

	$datasorted=$data | sort database -unique



	$white_space+=$datasorted


	}

################build html style http://technet.microsoft.com/en-us/library/ff730936.aspx########
$a = "<style>"
#$a = $a + "BODY{background-color:peachpuff;}"
$a = $a + "TABLE{border-width: 2px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;}"
$a = $a + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;}"
$a = $a + "</style>"

$white_space|ConvertTo-HTML -head $a  | Out-File .\white_space.htm

$message = new-object System.Net.Mail.MailMessage(“donotreply@labtest.com“, "Vikas.sukhija@labtest.com")
$message.IsBodyHtml = $True
$message.Subject = "Exchange 2007 database White Space Report"
$smtp = new-object Net.Mail.SmtpClient(“SMTP SERVER NAME“)
$body = get-content ".\white_space.htm"
$message.body = $body
$smtp.Send($message)

#################################################################################################

