SELECT TOP 20 cs-uri-stem,
COUNT(*) AS TotalRequest,
MAX(time-taken) AS MaxTime,
AVG(time-taken) AS AvgTime
INTO 'C:\Users\a24696\AppData\Roaming\ExLPT\Log Parser Studio\Output\avg_6-6-2023_12.41.775_PM.csv'
FROM 'C:\inetpub\logs\LogFiles\W3SVC1\*.log'
GROUP BY cs-uri-stem ORDER BY TotalRequest DESC