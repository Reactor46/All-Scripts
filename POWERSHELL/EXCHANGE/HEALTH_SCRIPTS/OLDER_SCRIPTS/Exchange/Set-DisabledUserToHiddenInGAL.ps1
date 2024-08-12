
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

Get-Mailbox -Filter{(HiddenFromAddressListsEnabled -eq $false) -AND (UserAccountControl -eq "AccountDisabled, NormalAccount")} #| 
    #Set-mailbox -HiddenFromAddressListsEnabled $True | Select Name, HiddenFromAddressListsEnabled , WhenChangedUTC |
        # Export-CSV -Path "C:\Scripts\Hide Users From GAL Reports\HiddenUsersGAL_$((Get-Date).ToString('MM-dd-yyyy_hh-mm-ss')).csv" -NoTypeInformation