#JORGE NAVARRO MANZANO Script Daily Multiple DCs DCdiag Report
#full dcdiag checking "failed test" or "no super" (spanish and english dcdiag)
#needs .net 2.0 and powershell 2.0(check version with get-host)
#execute dcdiags remotely in every dc and sent mail to the email address you choose
#https://es.linkedin.com/in/jorgenavarromanzano
#more scripts here https://github.com/jorgenavarromanzano

#Instructions:
#create dcs.txt
#example:
#dc1,domain\userdomainadmin,password,,,,
#dc20,domain2\userdomainadmin,password,,,,
#add in each , each check you want to skip:
#example:
#dc40,domain3\userdomainadmin,password,/skip:NCSecDesc,/skip:dns,,
#you need access to port 5985 and 5986 to remote dcs, powershell 2.0 and execute winrm quickconfig in those dcs

#change this variables to your emails and smtpserver:
$destemail = "john.battista@creditone.com"
$origemail = "dc_diag_test@creditone.com"
$smtpserver = "mailgateway.Contoso.corp"

$error.clear()
Start-Transcript log.txt -append
$dcs = Get-Content -Path .\dcs.txt
$texto= @()
$textoerrores= @()
$errores = 0

$texto += "command dcdiag /c /v /skip:systemlog"

foreach($dc in $dcs)
{
	$user = ($dc -split ",")[1]
	$password = ($dc -split ",")[2]
	$argumento1 = ($dc -split ",")[3]
	$argumento2 = ($dc -split ",")[4]
	$argumento3 = ($dc -split ",")[5]
	$argumento4 = ($dc -split ",")[6]
	$dc = ($dc -split ",")[0]
	$pass = ConvertTo-SecureString -AsPlainText $Password -Force
	$cred = New-Object System.Management.Automation.PSCredential -ArgumentList $user,$pass
	
	write-host "dcdiag: " + $dc + " argumentos: " + $argumento1 + " " + $argumento2 + " " + $argumento3 + " " + $argumento4
	$textoerrores += "dcdiag: " + $dc + " argumentos: " + $argumento1 + " " + $argumento2 + " " + $argumento3 + " " + $argumento4
	$dcdiag = Invoke-Command -computername $dc -credential $cred -argumentlist $user,$password,$argumento1,$argumento2,$argumento3,$argumento4 -scriptblock {param($user,$password,$argumento1,$argumento2,$argumento3,$argumento4) cmd /c dcdiag /c /v /u:$user /p:$password /skip:systemlog $argumento1 $argumento2 $argumento3 $argumento4}
	$dcdiag | select-string "failed test" | write-host
	$texto += $dc
	$texto += $dcdiag
	$texto+= "-------------------------------"
	$textoerrores += $dcdiag | select-string "failed test","no super? la prueba"
	
	if(($dcdiag | select-string "failed test","no super? la prueba").count -gt 0)
	{
		$errores+=1
	}		
}

write-host $textoerrores

if($errores -gt 0)
{
	$asunto = "SPADM01 Mon Active Directory, dcs num errors: " + $errores
}
if($errores -eq 0)
{
	$asunto = "SPADM01 Mon Active Directory, dcs no errors"
}

if($textoerrores.count -gt 0)
{
	$texto = Out-String -Inputobject $texto
	$textoerrores = Out-String -Inputobject $textoerrores
	$textocorreo = $textoerrores + $texto
	send-mailmessage -from $origemail -to $destemail -subject $asunto -body $textocorreo -smtpServer $smtpserver
	write-host "mail sent"
}

if($error.count -gt 0)
{
	$errores = Out-String -Inputobject $error
	send-mailmessage -from $origemail -to $destemail -subject "Mon Active Directory, execution error" -body $errores -smtpServer $smtpserver
}
Stop-Transcript

if( [int]((Get-ChildItem .\log.txt).Length / 1024 / 1024) -gt 200)
{
	if(test-path .\log.txt.old)
	{
		remove-item	.\log.txt.old
	}
	Rename-Item .\log.txt .\log.txt.old
}