#add-pssnapin Microsoft.Exchange.Management.PowerShell.Admin
import-module activedirectory
$userinformation = Import-Csv c:\autousers\table.csv
$email = $userinformation.email
$alias = $userinformation.alias
$db = $userinformation.db
$ou = $userinformation.ou
$FullN = $userinformation.fulln
$FirstN = $userinformation.firstn
$MI = $userinformation.mi
$LastN = $userinformation.lastn
$cemail = $userinformation.cemail
$semail = $userinformation.semail
$creator = $userinformation.creator
$password = "Welcome1"
$securepassword = $Password = ConvertTo-SecureString -AsPlainText $Password -Force
$date = get-date
$logpath = "c:\autousers\logs.csv"

try{
	New-mailbox -UserPrincipalName $alias+"@acp.local" -alias $alias -database $db -Name $fulln -OrganizationalUnit $ou -password $Password -FirstName $firstn -LastName $lastn -DisplayName $fulln -ResetPasswordOnNextLogon $true -erroraction stop
    Set-CASMailbox -Identity $alias -ActiveSyncEnabled $false
if ($db -match "West-1") {
	$westmail = $alias+"@west-1.com"
	set-mailbox $alias -emailaddresses SMTP:$westmail -emailaddresspolicyenabled $false
	#Set-User -Identity $alias -ChangePasswordAtNextLogon $false

#SEND EMAIL	
$message = "The user $fulln has been created by $creator. Their login is: $alias and their Password is $password Their email address is $westmail."
write-host $message
	send-mailmessage -to $cemail -cc $semail -subject "User Creation Successful for $fulln" -body $message -from itsupport@angelsofcare.com -smtpserver server3bl.acp.local
    "$date,$creator,$FullN,$alias,$db,$ou,SUCCEEDED" | out-file $logpath -append}
    
else{

#SEND EMAIL	
$message = "The user $fulln has been created by $creator. Their login is: $alias and their First Time Password is Welcome1 Their email address is $email."
	write-host $message
	send-mailmessage -to $cemail -cc $semail -subject "User Creation Successful for $fulln" -body $message -from itsupport@angelsofcare.com -smtpserver server3bl.acp.local
    "$date,$creator,$FullN,$alias,$db,$ou,SUCCEEDED" | out-file $logpath -append}
}
catch
{
     if ($Error[0].fullyqualifiederrorid -match "1C4D69FD,Microsoft.Exchange.Management.RecipientTasks.NewMailbox")
	 {
	 $alias = $FirstN.toLower().substring(0,1)+$MI.lower()+$LastN.tolower()
	 $email = $alias+"@angelosfcare.com"
	 
$message = "User creation has failed for Employee: $fulln. An IT Ticket has already been submitted and will be reviewed by the IT staff as soon as possible. Thanks."
write-host $message
	send-mailmessage -to $cemail -subject "User Creation Failed for $name" -body $message -from itsupport@angelsofcare.com -smtpserver server3bl.acp.local
$message = "User creation has failed for Employee: $fulln Database: $db OU: $ou. Creator: $creator With Error: $error"
	send-mailmessage -to itsupport@angelsofcare.com -subject "User Creation Failed for $name" -body $message -from $cemail -smtpserver server3bl.acp.local
    "$date,$creator,$FullN,$alias,$db,$ou,FAILED,$error.FullyQualifiedErrorID" | out-file $logpath -append
}

exit
    
