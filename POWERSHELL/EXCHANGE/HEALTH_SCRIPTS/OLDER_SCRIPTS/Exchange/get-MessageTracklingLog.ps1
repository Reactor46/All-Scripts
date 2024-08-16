Get-MessageTrackingLog -Sender "user@dmacc.edu"

Get-MessageTrackingLog -Start "9/01/2011 12:00AM" -End "09/30/2011 11:59PM" -Sender "strive@dmacc.edu" | Select-Object Timestamp,EventID,MessageSubject,Recipients,Sender > p:\scripts\powershell\exchange\strive.log