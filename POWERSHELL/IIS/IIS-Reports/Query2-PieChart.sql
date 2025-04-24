SELECT TOP 5 cs-uri-stem, time-taken INTO MyChart.GIF FROM 'C:\inetpub\logs\LogFiles\W3SVC1\*.log' 
ORDER BY time-taken DESC