Set WshShell = WScript.CreateObject("WScript.Shell")
Return = WshShell.Run("cmd /c C:\Windows\system32\certutil -Template ""CertificateTamplateName"" > \\netapp\CertLogs\%username%.txt", 0, true)