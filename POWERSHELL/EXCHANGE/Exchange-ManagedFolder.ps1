$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection 
Import-PSSession $Session 

Get-Mailbox -Filter {ArchiveStatus -Eq None -AND RecipientTypeDetails -eq UserMailbox} |

Enable-Mailbox -Archive $AllMailboxes = Get-Mailbox -Filter {(ArchiveStatus -Eq Active -AND RecipientTypeDetails -eq UserMailbox)} $AllMailboxes |
    ForEach {Start-ManagedFolderAssistant $_.Identity} Get-Mailbox -Filter {RecipientTypeDetails -eq UserMailbox} |
        Select PrimarySmtpAddress,ArchiveStatus |
            Out-GridView


$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session 
$UserMailbox = Get-Mailbox |
    Out-GridView -PassThru Get-Mailbox -Identity $UserMailbox.Alias |
        Enable-Mailbox -Archive Start-ManagedFolderAssistant -Identity
        Get-Mailbox -Identity $UserMailbox.Alias |
            Select PrimarySmtpAddress,ArchiveStatus | Out-GridView