##
## Author   : Roman Kuzmin
## Synopsis : Formats output as a table with a chart column
## Modified : 2006.11.06
##
## -Property: properties where the last one is numeric for a chart.
## -Width: chart column width, default is 1/2 of screen buffer.
## -ForeChar: character for chart bars.
## -BackChar: character for appending bars.
## -InputObject: the objects to be formatted.
#
# Chart of process working sets
#ps | Format-Chart Id, Name, WS
#
# Chart of file sizes in descending order
#dir | sort Length -desc | Format-Chart Name, Length
#
##

param
(
    [object[]]$Property = $(throw 'Supply properties'),
    [int]$Width = ($Host.UI.RawUI.BufferSize.Width/2),
    [char]$ForeChar = 9600,
    [char]$BackChar = 9617,
    [object[]]$InputObject
)

$set = $(if ($InputObject) {$InputObject} else {@($Input)}) |
Select-Object $Property

$max = ($set | Measure-Object ($Property[-1]) -Maximum).Maximum
if ($max -eq 0) {$max = 1}

$set | .{process{
    $_ | Add-Member -PassThru NoteProperty Chart (("$ForeChar"*(
    $_.$($Property[-1])/$max*$Width)).PadRight($Width, $BackChar))
}} |
Format-Table ($Property + 'Chart') -AutoSize

