# If the error contains
#
# Your message wasn't delivered to anyone because it's too large. The limit is XYZ MB. Your message is XYZ+ MB.
#
#Remote Server returned '550 5.2.12 RESOLVER.RST.SendSizeLimit.Org; message too large for this organization'
#
#Check the server's limits
Get-TransportConfig | Format-Table -Property MaxSendSize, MaxReceiveSize -CaseSensitive

Get-ReceiveConnector | Format-Table -Property Name, MaxMessageSize -CaseSensitive

Get-SendConnector | Format-Table -Property Name, MaxMessageSize -CaseSensitive

Get-Mailbox Administrator | Format-Table -Property Name, MaxSendSize, MaxReceiveSize -CaseSensitive
#
#Change the size limits
#
Set-TransportConfig -MaxSendSize 15MB -MaxReceiveSize 15MB

Get-Mailbox | Set-Mailbox -MaxSendSize 10MB -MaxReceiveSize 10MB

Get-TransportConfig | Set-TransportConfig -MaxSendSize 15MB -MaxReceiveSize 15MB

Get-ReceiveConnector | Set-ReceiveConnector -MaxMessageSize 10MB

Get-SendConnector | Set-SendConnector -MaxMessageSize 10MB

Get-Mailbox | Set-Mailbox -MaxSendSize 10MB -MaxReceiveSize 10MB
#
#Attachment size limit by creating a new rule
#
New-TransportRule -Name LargeAttach -AttachmentSizeOver 150MB -RejectMessageReasonText "Message attachment size over 100MB - email rejected."
