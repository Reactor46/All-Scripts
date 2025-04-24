SELECT 
    TO_LOWERCASE(EXTRACT_TOKEN(EXTRACT_EXTENSION(cs-uri-stem), -1, '/')) AS Site,
    COUNT(*) AS TotalRequests,
    SUM(CASE WHEN sc-status = 200 THEN 1 ELSE 0 END) AS SuccessfulRequests,
    SUM(CASE WHEN sc-status <> 200 THEN 1 ELSE 0 END) AS FailedRequests
INTO HTML 'D:\Reports\IISLogAnalysis.html'
FROM 'C:\inetpub\logs\LogFiles\W3SVC*\*.log'
GROUP BY Site
