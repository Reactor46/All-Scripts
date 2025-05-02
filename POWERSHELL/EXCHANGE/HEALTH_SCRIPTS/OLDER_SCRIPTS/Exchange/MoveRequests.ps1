Get-Mailbox -Database NewStudents -ResultSize 25 | New-MoveRequest -TargetDatabase Students_01 -BatchName "StudsMove"

Get-MoveRequest -BatcheName StudsMove

Get-MoveRequest -MoveStatus Completed | Remove-MoveRequest