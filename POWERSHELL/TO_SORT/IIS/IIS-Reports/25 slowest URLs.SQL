SELECT TOP 25 cs-uri-stem,
 COUNT(*) AS TotalRequest,
 MAX(time-taken) AS MaxTime,
 AVG(time-taken) AS AvgTime 
INTO 'C:\Users\a24696\AppData\Roaming\ExLPT\Log Parser Studio\Output\BeforeavgProb_6-6-2023_12.42.176_PM.csv'
FROM 'C:\inetpub\logs\LogFiles\W3SVC1\*.log'
WHERE TO_TIME(time)
 BETWEEN TIMESTAMP('00:00:00','hh:mm:ss') AND TIMESTAMP('14:00:00','hh:mm:ss')
 GROUP BY cs-uri-stem ORDER BY TotalRequest DESC