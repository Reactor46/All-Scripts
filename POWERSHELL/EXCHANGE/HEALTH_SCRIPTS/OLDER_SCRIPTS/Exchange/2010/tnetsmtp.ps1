#param([String] $remoteHost =$(throw "Please specify the Target Server"),[String] $domain = $(throw "Please specify the #recipient Domain"),[String] $sendingdomain = $(throw "Please specify the Sending Domain"))

param([String] $remoteHost,[String] $domain, [String] $sendingdomain)


if ($remotehost -eq "" -or $domain -eq "" -or $sendingdomain -eq "") {"Please specify the Target Server, recipient domain and sending domain" 
			return; }


function readResponse {

while($stream.DataAvailable)  
   {  
      $read = $stream.Read($buffer, 0, 1024)    
      write-host -n -foregroundcolor cyan ($encoding.GetString($buffer, 0, $read))  
      ""
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
write-host -foregroundcolor DarkGreen $command
""
$writer.WriteLine($command) 
$writer.Flush()
start-sleep -m 500 
readResponse($stream)
$command = "MAIL FROM: <smtpcheck@" + $sendingdomain + ">" 
write-host -foregroundcolor DarkGreen $command
""
$writer.WriteLine($command) 
$writer.Flush()
start-sleep -m 500 
readResponse($stream)
$command = "RCPT TO: <postmaster@" + $domain + ">" 
write-host -foregroundcolor DarkGreen $command
""
$writer.WriteLine($command) 
$writer.Flush()
start-sleep -m 500 
readResponse($stream)
$command = "QUIT" 
write-host -foregroundcolor DarkGreen $command
""
$writer.WriteLine($command) 
$writer.Flush()
start-sleep -m 500 
readResponse($stream)
## Close the streams 
$writer.Close() 
$stream.Close() 
