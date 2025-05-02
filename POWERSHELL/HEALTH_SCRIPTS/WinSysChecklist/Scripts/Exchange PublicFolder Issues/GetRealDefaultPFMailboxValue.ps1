# GetRealDefaultPFMailboxValue.ps1
# 
# In Exchange Management Shell, DefaultPublicFolderMailbox shows a random
# value if the value is not set on the mailbox. This makes it difficult
# to determine whether this value has been set on a mailbox or not.
# 
# The purpose of this script is to take a set of mailboxes and retrieve
# the actual DefaultPublicFolderMailbox value.
# 
# To use this script, retrieve the mailboxes you want to process and pipe
# them to the script.
# 
# Syntax:
# 
# $mailboxes = Get-Mailbox | .\GetRealDefaultPFMailboxValue.ps1
# 
# You can use the usual parameters to filter the mailbox result, such
# as only returning mailboxes for a certain server:
# 
# $mailboxes = GetMailbox -Server EXCH1 | .\GetRealDefaultPFMailboxValue.ps1
# 
# The mailboxes in $mailboxes will have a new property,
# msExchPublicFolderMailbox. This value shows you the real value of the
# LDAP property. For example:
# 
# [PS] C:\>$mailboxes = Get-Mailbox | .\GetRealDefaultPFMailboxValue.ps1
# [PS] C:\>$mailboxes | ft Name,msExch*
# 
# Name                                                        msExchPublicFolderMailbox
# ----                                                        -------------------------
# Administrator
# User 1                                                      CN=PFMB1,CN=Users,DC=child,DC=root,DC=test
# User 2
# 
# We can now see that the value is actually set on User 1, but not on the other two mailboxes.
# Since these objects still have the full mailbox properties, you can do
# any additional filtering you need based on whatever criteria you like.

foreach ($mailbox in $input)
{
    $dn = $mailbox.DistinguishedName
    $adObject = [ADSI]("GC://" + $dn)
    if ($adObject.Properties["msExchPublicFolderMailbox"].Count -gt 0)
    {
        $mailbox | Add-Member NoteProperty msExchPublicFolderMailbox $adObject.Properties["msExchPublicFolderMailbox"][0].ToString()
    }
    else
    {
        $mailbox | Add-Member NoteProperty msExchPublicFolderMailbox ""
    }

    $mailbox
}