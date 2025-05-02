On Error Resume Next 

Const HKEY_CLASSES_ROOT = &H80000000

strComputer = "."
strKeyPath1 = "Installer\Products\3101A8DA64E420E4582C13863C234823" 
strKeyPath2 = "Installer\Products\EA5CB419B592F4D46A65381157D2988B" 

Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv") 

DeleteSubkeys1 HKEY_CLASSES_ROOT, strKeypath1
DeleteSubkeys2 HKEY_CLASSES_ROOT, strKeypath2 

Sub DeleteSubkeys1(HKEY_CLASSES_ROOT, strKeyPath1) 
    objRegistry.EnumKey HKEY_CLASSES_ROOT, strKeyPath1, arrSubkeys 

    If IsArray(arrSubkeys1) Then 
        For Each strSubkey1 In arrSubkeys1 
            DeleteSubkeys HKEY_CLASSES_ROOT, strKeyPath1 & "\" & strSubkey1 
        Next 
    End If
    objRegistry.DeleteKey HKEY_CLASSES_ROOT, strKeyPath1 
End Sub

Sub DeleteSubkeys2(HKEY_CLASSES_ROOT, strKeyPath2) 
    objRegistry.EnumKey HKEY_CLASSES_ROOT, strKeyPath2, arrSubkeys 

    If IsArray(arrSubkeys2) Then 
        For Each strSubkey2 In arrSubkeys2 
            DeleteSubkeys HKEY_CLASSES_ROOT, strKeyPath2 & "\" & strSubkey2
        Next 
    End If
    objRegistry.DeleteKey HKEY_CLASSES_ROOT, strKeyPath2 
End Sub