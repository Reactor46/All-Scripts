

Get-User -ResultSize Unlimited | ? { $_.RecipientType -eq 'UserMailbox' -and $_.useraccountcontrol -like '*accountdisabled*'}| set-mailbox -HiddenFromAddressListsEnabled $true
## if users are placed in an OU
get-mailbox -OrganizationalUnit 'disabled users' | set-mailbox -HiddenFromAddressListsEnabled $true

get-mailbox  |? { $_.ExchangeUserAccountControl  -like '*accountdisabled*' -and $_.HiddenFromAddressListsEnabled -like 'False'} | set-mailbox -HiddenFromAddressListsEnabled $true

# Active directory

Get-ADUser  -Filter {(enabled -eq "false") -and (msExchHideFromAddressLists -notlike "*")} -Properties msExchHideFromAddressLists <# -SearchBase "OU=<OU>,DC=<Domain>,DC=<TLD>"`  #> | Set-ADUser -Add @{msExchHideFromAddressLists="TRUE"}