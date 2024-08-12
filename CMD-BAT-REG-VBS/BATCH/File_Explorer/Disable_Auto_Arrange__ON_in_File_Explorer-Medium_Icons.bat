@echo off

:: Created by: Shawn Brink
:: http://www.tenforums.com
:: Tutorial: http://www.tenforums.com/tutorials/14856-auto-arrange-folders-enable-disable-windows-10-a.html


:: To reset all folder views to default
Reg Delete "HKCU\SOFTWARE\Microsoft\Windows\Shell\BagMRU" /F
Reg Delete "HKCU\SOFTWARE\Microsoft\Windows\Shell\Bags" /F

Reg Delete "HKCU\SOFTWARE\Microsoft\Windows\ShellNoRoam\Bags" /F
Reg Delete "HKCU\SOFTWARE\Microsoft\Windows\ShellNoRoam\BagMRU" /F

Reg Delete "HKCU\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\Shell\BagMRU" /F
Reg Delete "HKCU\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags" /F

Reg Delete "HKCU\SOFTWARE\Classes\Wow6432Node\Local Settings\Software\Microsoft\Windows\Shell\Bags" /F
Reg Delete "HKCU\SOFTWARE\Classes\Wow6432Node\Local Settings\Software\Microsoft\Windows\Shell\BagMRU" /F


:: To turn off full row select setting
REG ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /V FullRowSelect /T REG_SZ /D 0 /F


:: To disable sniff folder feature and sets folder type to not specific
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags\AllFolders\Shell" /V FolderType /T REG_SZ /D NotSpecified /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags\AllFolders\Shell" /V KnownFolderDerivedFolderType /T REG_SZ /D {57807898-8C4F-4462-BB63-71042380B109} /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags\AllFolders\Shell" /V SniffedFolderType /T REG_SZ /D Generic /F



:: Disable Auto Arrange for General Items (not specific) folder template
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}" /V FFLAGS /T REG_DWORD /D 1124073477 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}" /V GroupByDirection /T REG_DWORD /D 1 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}" /V GroupByKey:FMTID /T REG_SZ /D {00000000-0000-0000-0000-000000000000} /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}" /V GroupByKey:PID /T REG_DWORD /D 0 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}" /V GroupView /T REG_DWORD /D 0 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}" /V IconSize /T REG_DWORD /D 48 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}" /V LogicalViewMode /T REG_DWORD /D 3 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}" /V Mode /T REG_DWORD /D 1 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}" /V Rev /T REG_DWORD /D 0 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}" /V Vid /T REG_SZ /D {0057D0E0-3573-11CF-AE69-08002B2E1262} /F



:: Disable Auto Arrange for Documents folder template
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{7d49d726-3c21-4f05-99aa-fdc2c9474656}" /V FFLAGS /T REG_DWORD /D 1124073477 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{7d49d726-3c21-4f05-99aa-fdc2c9474656}" /V GroupByDirection /T REG_DWORD /D 1 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{7d49d726-3c21-4f05-99aa-fdc2c9474656}" /V GroupByKey:FMTID /T REG_SZ /D {00000000-0000-0000-0000-000000000000} /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{7d49d726-3c21-4f05-99aa-fdc2c9474656}" /V GroupByKey:PID /T REG_DWORD /D 0 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{7d49d726-3c21-4f05-99aa-fdc2c9474656}" /V GroupView /T REG_DWORD /D 0 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{7d49d726-3c21-4f05-99aa-fdc2c9474656}" /V IconSize /T REG_DWORD /D 48 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{7d49d726-3c21-4f05-99aa-fdc2c9474656}" /V LogicalViewMode /T REG_dWORD /D 2 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{7d49d726-3c21-4f05-99aa-fdc2c9474656}" /V Mode /T REG_DWORD /D 6 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{7d49d726-3c21-4f05-99aa-fdc2c9474656}" /V Rev /T REG_DWORD /D 0 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{7d49d726-3c21-4f05-99aa-fdc2c9474656}" /V Vid /T REG_SZ /D {137E7700-3573-11CF-AE69-08002B2E1262} /F



:: Disable Auto Arrange for Music folder template
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{94d6ddcc-4a68-4175-a374-bd584a510b78}" /V FFLAGS /T REG_DWORD /D 1124073477 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{94d6ddcc-4a68-4175-a374-bd584a510b78}" /V GroupByDirection /T REG_DWORD /D 1 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{94d6ddcc-4a68-4175-a374-bd584a510b78}" /V GroupByKey:FMTID /T REG_SZ /D {00000000-0000-0000-0000-000000000000} /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{94d6ddcc-4a68-4175-a374-bd584a510b78}" /V GroupByKey:PID /T REG_DWORD /D 0 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{94d6ddcc-4a68-4175-a374-bd584a510b78}" /V GroupView /T REG_DWORD /D 0 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{94d6ddcc-4a68-4175-a374-bd584a510b78}" /V IconSize /T REG_DWORD /D 48 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{94d6ddcc-4a68-4175-a374-bd584a510b78}" /V LogicalViewMode /T REG_dWORD /D 3 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{94d6ddcc-4a68-4175-a374-bd584a510b78}" /V Mode /T REG_DWORD /D 1 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{94d6ddcc-4a68-4175-a374-bd584a510b78}" /V Rev /T REG_DWORD /D 0 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{94d6ddcc-4a68-4175-a374-bd584a510b78}" /V Vid /T REG_SZ /D {0057D0E0-3573-11CF-AE69-08002B2E1262} /F



:: Disable Auto Arrange for Pictures folder template
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{b3690e58-e961-423b-b687-386ebfd83239}" /V FFLAGS /T REG_DWORD /D 1124073477 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{b3690e58-e961-423b-b687-386ebfd83239}" /V GroupByDirection /T REG_DWORD /D 1 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{b3690e58-e961-423b-b687-386ebfd83239}" /V GroupByKey:FMTID /T REG_SZ /D {00000000-0000-0000-0000-000000000000} /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{b3690e58-e961-423b-b687-386ebfd83239}" /V GroupByKey:PID /T REG_DWORD /D 0 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{b3690e58-e961-423b-b687-386ebfd83239}" /V GroupView /T REG_DWORD /D 0 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{b3690e58-e961-423b-b687-386ebfd83239}" /V IconSize /T REG_DWORD /D 48 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{b3690e58-e961-423b-b687-386ebfd83239}" /V LogicalViewMode /T REG_dWORD /D 2 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{b3690e58-e961-423b-b687-386ebfd83239}" /V Mode /T REG_DWORD /D 6 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{b3690e58-e961-423b-b687-386ebfd83239}" /V Rev /T REG_DWORD /D 0 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{b3690e58-e961-423b-b687-386ebfd83239}" /V Vid /T REG_SZ /D {0057D0E0-3573-11CF-AE69-08002B2E1262} /F



:: Disable Auto Arrange for Videos folder template
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5fa96407-7e77-483c-ac93-691d05850de8}" /V FFLAGS /T REG_DWORD /D 1124073477 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5fa96407-7e77-483c-ac93-691d05850de8}" /V GroupByDirection /T REG_DWORD /D 1 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5fa96407-7e77-483c-ac93-691d05850de8}" /V GroupByKey:FMTID /T REG_SZ /D {00000000-0000-0000-0000-000000000000} /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5fa96407-7e77-483c-ac93-691d05850de8}" /V GroupByKey:PID /T REG_DWORD /D 0 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5fa96407-7e77-483c-ac93-691d05850de8}" /V GroupView /T REG_DWORD /D 0 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5fa96407-7e77-483c-ac93-691d05850de8}" /V IconSize /T REG_DWORD /D 48 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5fa96407-7e77-483c-ac93-691d05850de8}" /V LogicalViewMode /T REG_dWORD /D 2 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5fa96407-7e77-483c-ac93-691d05850de8}" /V Mode /T REG_DWORD /D 6 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5fa96407-7e77-483c-ac93-691d05850de8}" /V Rev /T REG_DWORD /D 0 /F
REG ADD "HKCU\SOFTWARE\Classes\Local Settings\SOFTWARE\Microsoft\Windows\Shell\Bags\AllFolders\Shell\{5fa96407-7e77-483c-ac93-691d05850de8}" /V Vid /T REG_SZ /D {0057D0E0-3573-11CF-AE69-08002B2E1262} /F


:: To kill and restart explorer
taskkill /f /im explorer.exe
start explorer.exe
