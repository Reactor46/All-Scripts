# api: multitool
# version: 0.2
# title: Unlock user
# description: Unlock-ADAccount
# category: user
# type: inline
# shortcut: 5
# icon: unlock
#
# Unlock user account


Param(
    $user=(Read-Host "User")
)

#-- check, unlock
if ((Get-ADUser $user -Prop Lockedout).LockedOut) {
    Write-Host -f Green "The User ID '$user' has been unlocked. ✔"
    Unlock-ADAccount $user -Verbose
}
else {
    Write-Host -f Red "☒ The User ID '$user' wasn't locked."
}
Get-ADUser $user -Prop * | FL -Prop samaccountname,employeeType,Enabled,LockedOut,SID,PasswordExpired,BadLogonCount
