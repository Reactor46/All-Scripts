# api: multitool
# title: inline command
# description: PS1 script without actual code
# category: test
# version: 0.1
# type: command
# command: Write-Host '--inline--'
#
# This file will not be run. Instead the `command:` gets executed
# per Invoke-Command. (Usually Run-GuiTask would invoke the $meta.fn
# script. But the older approach with command: still works.)
#
# So basically this is just a comment script. Originally the command:
# feature was meant for predefined/ad-hoc $menu[] entries.
