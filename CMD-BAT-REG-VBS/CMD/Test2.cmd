$1 = gci C:\Scripts\Repository\jbattista\Web\Sydi\TST\*.xml

Clear-host
$ErrorActionPreference = 'SilentlyContinue'

foreach ($file in $1){

$3 = $file.Name

$split = $3.split('.')[0].split(' ')

cscript "C:\Scripts\Repository\jbattista\SYDI\sydi-server\tools\sydi-transform.vbs" -x"$file" -sC:\Scripts\Repository\jbattista\Web\Sydi\TST\serverhtml.xsl -oC:\Scripts\Repository\jbattista\Web\Sydi\TST\$split.html
}