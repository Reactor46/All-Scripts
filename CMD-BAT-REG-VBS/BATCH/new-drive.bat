@ECHO OFF
:Map Network Share Drive
net use X: \\sdrive\Company /persistent:yes
net use Y: \\msoit03\Test /persistent:yes