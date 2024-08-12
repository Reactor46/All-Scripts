@ECHO OFF
RMDIR /S/Q "%APPDATA%\Avaya\one-X Agent\2.5\Profiles"
MD "%APPDATA%\Avaya\one-X Agent\2.5\Profiles"
XCOPY /E /C /H /I /Y "\\fnbm.corp\netlogon\Avaya One-X Profiles\CreditOneCAS" "%APPDATA%\Avaya\one-X Agent\2.5\Profiles\CreditOneCAS"
XCOPY /E /C /H /I /Y "\\fnbm.corp\netlogon\Avaya One-X Profiles\CreditOneCASH" "%APPDATA%\Avaya\one-X Agent\2.5\Profiles\CreditOneCASH"
XCOPY /E /C /H /I /Y "\\fnbm.corp\netlogon\Avaya One-X Profiles\CreditOneTest6416" "%APPDATA%\Avaya\one-X Agent\2.5\Profiles\CreditOneTest6416"
XCOPY /E /C /H /I /Y "\\fnbm.corp\netlogon\Avaya One-X Profiles\CreditOneTest2410" "%APPDATA%\Avaya\one-X Agent\2.5\Profiles\CreditOneTest2410"
XCOPY /E /C /H /I /Y "\\fnbm.corp\netlogon\Avaya One-X Profiles\CreditOneTestSup" "%APPDATA%\Avaya\one-X Agent\2.5\Profiles\CreditOneTestSup"
XCOPY /E /C /H /I /Y "\\fnbm.corp\netlogon\Avaya One-X Profiles\default" "%APPDATA%\Avaya\one-X Agent\2.5\Profiles\default"