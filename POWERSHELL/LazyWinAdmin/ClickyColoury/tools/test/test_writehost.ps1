# api: multitool
# type: inline
# category: test
# title: Write-Host
# description: test color output
# version: 0.1

"Red,Green,Yellow,Black,Orange".split(",") | % { Write-Host $_ -f $_ }


