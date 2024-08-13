# api: multitool
# version: 0.1
# title: NET USER
# description: Get NET USER details
# type: inline
# category: user
# hidden: 0
# key: i6|netuser
# config: {}
# 
# NET USER


Param($username = (Read-Host "USer"))

NET USER $username /DOMAIN