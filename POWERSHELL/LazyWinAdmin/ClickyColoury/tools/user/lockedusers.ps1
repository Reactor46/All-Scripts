# api: multitool
# version: 1.0
# title: Locked out users
# description: Find locked out user names
# type: inline
# category: user
# icon: user
# hidden: 0
# key: i5|lock|locked|lockedout|ll
# config: {}
# 
# Scans AD for locked-out accounts


Search-ADAccount -LockedOut | Select samaccountname, name

