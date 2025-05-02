param([string]$Mailbox,[string]$FolderName="OrgContacts");
#
# Copy-OrgContactsToUserMailboxContacts.ps1
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
# RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
# Parameters
#  Mandatory:
# -MailboxFolder : Folder to "own" for these contacts.
#
# Creates s OrgContacts folder in the Mailbox and adds the contacts into it. Does not attempt to 


$EwsUrl = ([array](Get-WebServicesVirtualDirectory))[0].InternalURL.AbsoluteURI

$ContactMapping=@{
    "FirstName" = "GivenName";
    "LastName" = "Surname";
    "Company" = "CompanyName";
    "Department" = "Department";
    "Title" = "JobTitle";
    "WindowsEmailAddress" = "Email:EmailAddress1";
    "Phone" = "Phone:BusinessPhone";
    "MobilePhone" = "Phone:MobilePhone";
}

$UserMailbox  = Get-Mailbox $Mailbox

if (!$UserMailbox)
{
	throw "Mailbox $($Mailbox) not found";
	exit;
}

$EmailAddress = $UserMailbox.PrimarySMTPAddress

# Load EWS Managed API
[void][Reflection.Assembly]::LoadFile("C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll");

$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)
$service.UseDefaultCredentials = $true;
$service.URL = New-Object Uri($EwsUrl);

# Search for an existing copy of the Folder to store Org contacts 
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress);
$RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
$RootFolder.Load()

$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress);
$FolderView = new-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$ContactsFolderSearch = $RootFolder.FindFolders($FolderView) | Where {$_.DisplayName -eq $FolderName}
if ($ContactsFolderSearch)
{
	# Empty if found
	$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress);
	$ContactsFolder = [Microsoft.Exchange.WebServices.Data.ContactsFolder]::Bind($service,$ContactsFolderSearch.Id);
	$ContactsFolder.Empty([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete, $true)
} else {

	# Create new contacts folder
	$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress);
	$ContactsFolder = New-Object Microsoft.Exchange.WebServices.Data.ContactsFolder($service);
	$ContactsFolder.DisplayName = $FolderName
	$ContactsFolder.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
	# Search for the new folder instance
	$RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
	$RootFolder.Load()
	$FolderView = new-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
	$ContactsFolderSearch = $RootFolder.FindFolders($FolderView) | Where {$_.DisplayName -eq $FolderName}
	$ContactsFolder = [Microsoft.Exchange.WebServices.Data.ContactsFolder]::Bind($service,$ContactsFolderSearch.Id);
}

# Add contacts
$Users = get-user -Filter {WindowsEmailAddress -ne $null -and (MobilePhone -ne $null -or Phone -ne $null) -and WindowsEmailAddress -ne $EmailAddress} 
$Users = $Users | select DisplayName,FirstName,LastName,Title,Company,Department,WindowsEmailAddress,Phone,MobilePhone

foreach ($ContactItem in $Users)
{
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress);
    
    $ExchangeContact = New-Object Microsoft.Exchange.WebServices.Data.Contact($service);
    if ($ContactItem.FirstName -and $ContactItem.LastName)
    {
        $ExchangeContact.NickName = $ContactItem.FirstName + " " + $ContactItem.LastName;
    }
    elseif ($ContactItem.FirstName -and !$ContactItem.LastName)
    {
        $ExchangeContact.NickName = $ContactItem.FirstName;
    }
    elseif (!$ContactItem.FirstName -and $ContactItem.LastName)
    {
        $ExchangeContact.NickName = $ContactItem.LastName;
    }
	elseif (!$ContactItem.FirstName -and !$ContactItem.LastName)
    {
        $ExchangeContact.NickName = $ContactItem.DisplayName;
		$ContactItem.FirstName = $ContactItem.DisplayName;
    }
	
    $ExchangeContact.DisplayName = $ExchangeContact.NickName;
    $ExchangeContact.FileAs = $ExchangeContact.NickName;
    
    # This uses the Contact Mapping above to save coding each and every field, one by one. Instead we look for a mapping and perform an action on
    # what maps across. As some methods need more "code" a fake multi-dimensional array (seperated by :'s) is used where needed.
    foreach ($Key in $ContactMapping.Keys)
    {
        # Only do something if the key exists
        if ($ContactItem.$Key)
        {
            # Will this call a more complicated mapping?
            if ($ContactMapping[$Key] -like "*:*")
            {
                # Make an array using the : to split items.
                $MappingArray = $ContactMapping[$Key].Split(":")
                # Do action
                switch ($MappingArray[0])
                {
                    "Email"
                    {
						$ExchangeContact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::($MappingArray[1])] = $ContactItem.$Key.ToString();
                    }
                    "Phone"
                    {
                        $ExchangeContact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::($MappingArray[1])] = $ContactItem.$Key;
                    }
                }                
            } else {
                $ExchangeContact.($ContactMapping[$Key]) = $ContactItem.$Key;            
            }
            
        }    
    }
    # Save the contact    
    $ExchangeContact.Save($ContactsFolder.Id);
    
    # Provide output that can be used on the pipeline
    $Output_Object = New-Object Object;
    $Output_Object | Add-Member NoteProperty FileAs $ExchangeContact.FileAs;
    $Output_Object | Add-Member NoteProperty GivenName $ExchangeContact.GivenName;
    $Output_Object | Add-Member NoteProperty Surname $ExchangeContact.Surname;
    $Output_Object | Add-Member NoteProperty EmailAddress1 $ExchangeContact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1]
    $Output_Object;
}