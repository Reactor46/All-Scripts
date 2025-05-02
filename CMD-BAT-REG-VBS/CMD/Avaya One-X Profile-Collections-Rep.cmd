@ECHO OFF
RMDIR /S/Q "%APPDATA%\Avaya\one-X Agent\2.5\Profiles"
MD "%APPDATA%\Avaya\one-X Agent\2.5\Profiles"
XCOPY /E /C /H /I /Y "\\Contoso.corp\netlogon\Avaya One-X Profiles\CreditOneCASH" "%APPDATA%\Avaya\one-X Agent\2.5\Profiles\CreditOneCASH"