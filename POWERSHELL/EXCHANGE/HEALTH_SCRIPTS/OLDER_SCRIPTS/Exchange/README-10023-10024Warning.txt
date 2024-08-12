1. Run Checkfor10024Warning in PowerShell ISE
2. Run Checkfor10024Warning in PowerShell ISE
3. Make note of any GUID returned (if any go to step 4). If no GUID is returned, then there are no mailboxes over 20GB or 30GB.

4. Copy GUID from step 3 (such as d94e51a8-8cb5-47a6-a519-62c9d6a92561)

5. Open Exchange PowerShell
Get-Mailbox (paste in guid, such as Get-Mailbox d94e51a8-8cb5-47a6-a519-62c9d6a92561)
Results will covert GUID into a mailbox name.
Make note of the ALIAS

6. Run the following in Exchange Powershell (replace ALIAS with actual alias), this will give the size of recoverable items.
Get-MailboxFolderStatistics -Identity 'ALIAS' -FolderScope RecoverableItems | Format-Table Name,FolderPath,ItemsInFolder,FolderAndSubfolderSize

7. Check mailbox retention, make note of the number (mailbox 
Get-Mailbox ALIAS | fl *retain* 

8. If results in step 6 are => 14, then set mailbox for 10 day retention
Set-Mailbox ALIAS -RetainDeletedItems 10

9. Turn off litigation hold
Set-Mailbox ALIAS -LitigationHoldEnabled $false

10. Wait ~5mins

11. Run command
Search-Mailbox ALIAS -SearchDumpsterOnly -DeleteContent -Confirm:$false

12. To monitor results, open second Exchange Shell, and run the command in step 5.

13. When command completes set mailbox back to defaults
Set-Mailbox ALIAS -RetainDeletedItems (number from results in step 7)
Set-Mailbox ALIAS -LitigationHoldEnabled $true


