SELECT to_lowercase(cs-uri-stem) AS URI,
 SUM([_200]) AS [200s], SUM([_304]) AS [304s],
 SUM([_302]) AS [302s], SUM([_404]) AS [404s],
 SUM([_301]) AS [301s], SUM([_500]) AS [500s],
 SUM([_501]) AS [501s],SUM([_403]) AS [403s], 
 SUM([_206]) AS [206s],SUM([_406]) AS [406s], 
 SUM([_400]) AS [400s], 
sub(count(*),add([200s],[206s])) as Failures USING CASE TO_STRING(sc-status)
 WHEN '200' THEN 1 ELSE 0 END AS [_200], CASE TO_STRING(sc-status)
 WHEN '304' THEN 1 ELSE 0 END AS [_304], CASE TO_STRING(sc-status)
 WHEN '302' THEN 1 ELSE 0 END AS [_302], CASE TO_STRING(sc-status) 
 WHEN '404' THEN 1 ELSE 0 END AS [_404], CASE TO_STRING(sc-status)
 WHEN '301' THEN 1 ELSE 0 END AS [_301], CASE TO_STRING(sc-status)
 WHEN '500' THEN 1 ELSE 0 END AS [_500], CASE TO_STRING(sc-status)
 WHEN '501' THEN 1 ELSE 0 END AS [_501], CASE TO_STRING(sc-status)
 WHEN '403' THEN 1 ELSE 0 END AS [_403], CASE TO_STRING(sc-status)
 WHEN '206' THEN 1 ELSE 0 END AS [_206], CASE TO_STRING(sc-status)
 WHEN '406' THEN 1 ELSE 0 END AS [_406], CASE TO_STRING(sc-status)
 WHEN '400' THEN 1 ELSE 0 END AS [_400] 
FROM 'C:\inetpub\logs\LogFiles\W3SVC1\*.log'
GROUP BY URI ORDER BY Failures DESC