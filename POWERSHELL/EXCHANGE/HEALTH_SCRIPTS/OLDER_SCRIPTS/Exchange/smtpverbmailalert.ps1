$remoteHost = "host.targetdomain.com"
$domain = "targetdomain.com"
$sendingdomain = "youdomain.com"
$Global:SendAlert = 0
$Global:SentAlert = 0
$AlertFilePath="c:\temp\msMailAlert.txt"
$SmtpServerforAlertDelivery = "smsgateway.com"
$AlertFrom = "mailserverwarning@youdomain.com"
$AerttTo = "0419999999@youcelphone.com"

function readResponse($verb) {

while($stream.DataAvailable)  
   {  
      $read = $stream.Read($buffer, 0, 1024)   
      $rstring = $encoding.GetString($buffer, 0, $read)
      switch ($verb){
	"HELO" {if ($rstring.substring(0,3) -ne "220"){
		if ($rstring.substring(0,3) -ne "250"){
			SendAlert("HELO verb Result " + $rstring.substring(0,3))}} 
		}    
	"FROM" {if ($rstring.substring(0,3) -ne "250"){SendAlert("From verb Result " + $rstring.substring(0,3))} }  
	"TO" {if ($rstring.substring(0,3) -ne "250"){SendAlert("To verb ResultStatus " + $rstring.substring(0,3))} }  
      }
   } 
}

function SendAlert($Problem) {

# Create and write to a file
if (! [System.Io.File]::Exists($AlertFilePath)) 
	{ $alertFile=[System.Io.File]::Createtext($AlertFilePath)
	  $cdate = get-date
	  $alertFile.WriteLine($cdate.ToString())
	  $alertFile.Close()
	  $Global:SendAlert = 1
   } 
else{
	$alertFile = [System.Io.File]::OpenText($AlertFilePath)
	$alertTime = $alertFile.ReadLine()
	$alertFile.Close()
	$LastAlertTimeSpan =  New-TimeSpan -start ($alertTime) -end $(Get-Date)
	if ($LastAlertTimeSpan.TotalMinutes -gt 60){
		[System.Io.File]::Delete($AlertFilePath)
		if (! [System.Io.File]::Exists($AlertFilePath)) 
			{ $alertFile=[System.Io.File]::Createtext($AlertFilePath)
			 $cdate = get-date
			 $alertFile.WriteLine($cdate.ToString())
			 $alertFile.Close()
			 $Global:SendAlert = 1
			 }
   	}
	else{
		$Global:SendAlert = 0
	}
	
}

if ($Global:SendAlert -eq 1 -band $Global:SentAlert -eq 0) {

$Title = "Mail Server Failed " + $Problem
$Body = $Problem
$SmtpClient = new-object system.net.mail.smtpClient
$SmtpClient.host = $SmtpServerforAlertDelivery
$SmtpClient.Send($AlertFrom,$AerttTo,$title,$Body)
$Global:SentAlert = 1

}

}

$port = 25 
$socket = new-object System.Net.Sockets.TcpClient($remoteHost, $port) 
if($socket -eq $null) { return; } 
$stream = $socket.GetStream() 
$writer = new-object System.IO.StreamWriter($stream) 
$buffer = new-object System.Byte[] 1024 
$encoding = new-object System.Text.AsciiEncoding 
readResponse($stream)
$command = "HELO "+ $domain 
$writer.WriteLine($command) 
$writer.Flush()
start-sleep -m 500 
readResponse($stream,"HELO")
$command = "MAIL FROM: <smtpcheck@" + $sendingdomain + ">" 
$writer.WriteLine($command) 
$writer.Flush()
start-sleep -m 500 
readResponse($stream,"FROM")
$command = "RCPT TO: <postmaster@" + $domain + ">" 
$writer.WriteLine($command) 
$writer.Flush()
start-sleep -m 500 
readResponse($stream,"TO")
$command = "QUIT" 
$writer.WriteLine($command) 
$writer.Flush()
start-sleep -m 500 
readResponse($stream,"QUIT")
## Close the streams 
$writer.Close() 
$stream.Close() 

