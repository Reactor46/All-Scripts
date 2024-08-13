# api: multitool
# version: 0.1
# title: EventVwr
# description: start EventViewer on remote computer
# type: inline
# category: extras
# img: tools.png
# hidden: 1
# noheader: 1
# key: t9|eventvwr
# config: -
#
# End

Start-Process eventvwr.exe -ArgumentList $machine

