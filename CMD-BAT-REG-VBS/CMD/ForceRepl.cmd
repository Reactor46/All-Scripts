@ECHO ON
ntfrsutl.exe forcerepl LASDC05 /r "Domain System Volume (SYSVOL share)" /p LASDC01
ntfrsutl.exe forcerepl LASDC02 /r "Domain System Volume (SYSVOL share)" /p LASDC01