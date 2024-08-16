$ExchangeServer = "" # Enter Client Access server's Hostname/IP

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($ExchangeServer)/PowerShell/" -Authentication Kerberos -Credential $UserCredential

$servers = Get-ExchangeServer -Domain msk.lo | ? {($_.ServerRole -like "*HubTransport*") -and ($_.Name -notlike "*IQ*") -and ($_.Name -notlike "*gltb*")}
$smtpServer = “127.0.0.1”  # Enter SMTP server's Hostname/IP
$msg = new-object Net.Mail.MailMessage  
$smtp = new-object Net.Mail.SmtpClient($smtpServer) 
$msg.From = “HUBQueue@msk.lo” # MAIL FROM 
$msg.To.Add("its.infra@msk.lo") # RCPT TO
while ($true) {
    foreach ($server in $servers){
        
        Write-Output "$(get-date)    $($server)"
        $queues = Get-Queue -Server $server.Name
        foreach ($queue in $queues){
            
            if (($queue.LastError -ne "") -and ($queue.MessageCount -gt 3)){

 
                $msg.Subject = "$($server) Transport Queue error” 
                $msg.Body = “Server reported an error on NexHop Domain named $($queue.NextHopDomain) with message $($queue.LastError). Stucked messages count is $($queue.MessageCount)” 
                $smtp.Send($msg) 

            }

        }
        write-output "Sleep 5 sec"
        start-sleep 5
    
    }

    Start-Sleep 60

}