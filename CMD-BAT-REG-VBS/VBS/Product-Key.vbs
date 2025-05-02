'Convert binary to chars 
Function ConvertToKey(Key) 
    Const KeyOffset = 52 
    Dim isWin8, Maps, i, j, Current, KeyOutput, Last, keypart1, insert 
    'Check if OS is Windows 8 
    isWin8 = (Key(66) \ 6) And 1 
    Key(66) = (Key(66) And &HF7) Or ((isWin8 And 2) * 4) 
    i = 24 
    Maps = "BCDFGHJKMPQRTVWXY2346789" 
    Do 
           Current= 0 
        j = 14 
        Do 
           Current = Current* 256 
           Current = Key(j + KeyOffset) + Current 
           Key(j + KeyOffset) = (Current \ 24) 
           Current=Current Mod 24 
            j = j -1 
        Loop While j >= 0 
        i = i -1 
        KeyOutput = Mid(Maps,Current+ 1, 1) & KeyOutput 
        Last = Current 
    Loop While i >= 0  
    keypart1 = Mid(KeyOutput, 2, Last) 
    insert = "N" 
    KeyOutput = Replace(KeyOutput, keypart1, keypart1 & insert, 2, 1, 0) 
    If Last = 0 Then KeyOutput = insert & KeyOutput 
    ConvertToKey = Mid(KeyOutput, 1, 5) & "-" & Mid(KeyOutput, 6, 5) & "-" & Mid(KeyOutput, 11, 5) & "-" & Mid(KeyOutput, 16, 5) & "-" & Mid(KeyOutput, 21, 5)