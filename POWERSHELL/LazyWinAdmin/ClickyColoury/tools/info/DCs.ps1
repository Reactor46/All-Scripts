# api: multitool
# version: 1.3
# title: Domain controllers
# description: List DCs
# type: inline
# category: info
# hidden: 0
# key: i1|DCs|domain|controllers|pdc
# config: {}
# 
# Shows list of active domain controlles


Get-ADDomainController -Filter * | FT -Auto -Wrap Name,Enabled,Site,IPV4Address,SslPort,LdapPort,IsReadOnly | Out-String -Width 100

