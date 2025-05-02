# api: multitool
# version: 1.4
# title: Find user 
# description: Performs an AD search for phone numbers or employee numbers
# type: inline
# category: user
# hidden: 0
# icon: finduser
# key: i4|fu|f|find|find-?user
# keycode: F6
# shortcut: 4
# config: -
# 
# Scans AD user list for all phone numbers, or for employee IDs.
#
#  → The graphical WPF-MultiTool has a simpler version built into
#    the `User` field/dropdown already (- only scans for the main
#    telephone number though).
#
#  → Should use Out-GridView or Format-Table per config option.
#    (Not reimplemented yet for GUI or CLI version).


Param($find = (Read-Host "User"));

# adapt for LIKE
if ($find -match "%") {        # SQL placeholders
    $find = $find -replace "%","*"
}
if ($find -notmatch "\*") {
    if ($find -match "[A-Z]{3,}") {   # usernames
        $find = "*$find*"
    }
    else {                         # numbers only
        $find = "*$find"
    }
}

# search
$ls = (Get-ADUser -Filter {
  (TelephoneNumber -like $find) -or (MobilePhone -like $find) -or
  (employeeNumber -like $find) -or (homePhone -like $find) -or (displayname -like $find)
} -Properties samaccountname,displayname,telephoneNumber,mobilePhone,homePhone,employeeNumber |
Select-Object samaccountname,displayname,telephonenumber,mobilephone,homePhone,employeeNumber)
                      
# output
if ($cfg.gridview -match "GridView") {
    $ls | Out-GridView
}
else {
    $ls | Format-Table -Auto -Wrap | Out-String -Width 120
}
