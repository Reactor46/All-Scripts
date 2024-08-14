move *IvrAURecord*.csv IvrAURecordings.csv
xcopy ivrau*.csv \\fnbmcorp\share\shared\IT\SupportServices\NOC\Monitoring\AU_Recording_Monitoring\lists\ /y
timeout 2
del ivraurecordings.csv
timeout 360
del \\fnbmcorp\share\shared\IT\SupportServices\NOC\Monitoring\AU_Recording_Monitoring\lists\IvrAURecordings.csv