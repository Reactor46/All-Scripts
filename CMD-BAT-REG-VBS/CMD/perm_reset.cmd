@ECHO ON
icacls E:\Home\*.* /inheritance:r /T
takeown E:\Home\*.* /r /d y
REM User DOMAIN\abos
SET _user=abos
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\aclark
SET _user=aclark
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\ajensen
SET _user=ajensen
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\amarshall
SET _user=amarshall
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\amourad
SET _user=amourad
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\bcolnar
SET _user=bcolnar
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\bdhillon
SET _user=bdhillon
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\beilers
SET _user=beilers
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM**** Replace DOMAIN with your Domain Name (IE: CONTOSO\_user%)
REM**** This will reset the NTFS permissions on the home drive for users
REM**** on the domain. This is typically used for Personal Drives for users.


REM User DOMAIN\bjanzen
SET _user=bjanzen
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\cjacobsen
SET _user=cjacobsen
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\cwilson
SET _user=cwilson
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\dholt
SET _user=dholt
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\dmenzies
SET _user=dmenzies
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\ghartl
SET _user=ghartl
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\hmullan
SET _user=hmullan
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\jbattista
SET _user=jbattista
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\jhutton
SET _user=jhutton
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\johnfriesen
SET _user=johnfriesen
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\jwhite
SET _user=jwhite
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\kbrassart
SET _user=kbrassart
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\kbreukelman
SET _user=kbreukelman
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\keckstein
SET _user=keckstein
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\krobertson
SET _user=krobertson
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\lclayton
SET _user=lclayton
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\lendersby
SET _user=lendersby
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\lsnider
SET _user=lsnider
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\mbaker
SET _user=mbaker
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\mbothwell
SET _user=mbothwell
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\melodyb
SET _user=melodyb
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\mfiliatrault
SET _user=mfiliatrault
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\nredline
SET _user=nredline
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\nbains
SET _user=nbains
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\plamboo
SET _user=plamboo
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\pstrandt
SET _user=pstrandt
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\reception
SET _user=reception
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\sbrown
SET _user=sbrown
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\stipton
SET _user=stipton
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\tmeyer
SET _user=tmeyer
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\tyolland
SET _user=tyolland
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
REM User DOMAIN\vmackey
SET _user=vmackey
icacls e:\home\%_user% /inheritance:r /remove:g * /T /C
icacls e:\home\%_user% /grant DOMAIN\%_user%:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /grant BUILTIN\Administrators:(OI)(CI)(F) /T /C
icacls e:\home\%_user% /setowner BUILTIN\Administrators /T /C
