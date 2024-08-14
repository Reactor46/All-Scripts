move *fulfillment*.csv IvrAppFulfillmentRecordings.csv
xcopy *fulfillment*.csv \\fnbmcorp\share\shared\IT\SupportServices\NOC\Monitoring\APPFulfillment_Recording_Monitoring\lists\ /y
timeout 2
del *fulfillment*.csv
timeout 360
del \\fnbmcorp\share\shared\IT\SupportServices\NOC\Monitoring\APPFulfillment_Recording_Monitoring\lists\ivrapp*.csv