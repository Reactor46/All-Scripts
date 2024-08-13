Get-ChildItem -path C:\ -recurse |dir | ? {$_.lastwritetime -gt '11/7/15' -AND $_.lastwritetime -lt '11/9/15'}|
    ft -AutoSize |
        out-file -filepath "c:\temp\modified.txt"-append