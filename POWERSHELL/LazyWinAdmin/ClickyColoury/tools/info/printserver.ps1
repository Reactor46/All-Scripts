# api: multitool
# version: 1.1
# title: Printservers
# description: scans AD for print servers
# type: inline
# category: info
# hidden: 0
# key: i7|print|print-?se?rve?r?
# config: {}
# 
# scans AD for print servers


$ls = Get-ADObject -LDAPFilter "(&(&(&(uncName=*)(objectCategory=printQueue))))" -Prop * |
        Sort-Object -Unique -Property servername |
        Select servername

$ls | Format-Table -Auto -Wrap

