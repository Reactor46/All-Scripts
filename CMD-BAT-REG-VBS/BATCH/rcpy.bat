robocopy %1 %2 /MIR /copyall /r:4 /w:1 /zb /fp  /V /ts | find /i /v "same" >> C:\scripts\robocopy\logs\%3_Share_1.txt

timeout /T 5

exit