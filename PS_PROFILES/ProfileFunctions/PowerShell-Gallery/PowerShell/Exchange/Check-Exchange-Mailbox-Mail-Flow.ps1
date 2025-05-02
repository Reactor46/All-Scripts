Get-MessageTrackingLog -ResultSize Unlimited -Start "11/24/2023 8:00AM" -End "11/27/2023 5:00PM" -Sender "mailbox@example.co.uk" | Format-Table
#
# Check the mail flow using only a subject
#
Get-MessageTrackingLog -MessageSubject "SUBJECT"

