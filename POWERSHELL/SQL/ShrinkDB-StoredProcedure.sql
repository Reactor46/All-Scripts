-- ================================================================================= 
-- Author:         Eli Leiba 
-- Create date:    2020-03
-- Procedure Name: dbo.usp_ShrinkAllLogsExcludeSysDBS 
-- Description:    This procedure shrinks all user databases log files or a specific user database log 
-- ================================================================================== 
CREATE PROCEDURE dbo.usp_ShrinkAllLogsExcludeSysDBS (@dbname SYSNAME = '%')
AS
BEGIN
 
DECLARE @TSQLExec VARCHAR (MAX) = '';
SET NOCOUNT ON;
 
IF OBJECT_ID('tempdb..#Temp') IS NOT NULL
       DROP TABLE #temp

CREATE TABLE #temp (dbname sysname, dbid int, logFileSizeBeforeMB decimal(15,2), logFileSizeAfterMB decimal(15,2));

WITH fs
AS
(
    SELECT database_id, TYPE, SIZE * 8.0 / 1024 SIZE
    FROM sys.master_files
)
INSERT INTO #temp (dbname, dbid, logFileSizeBeforeMB)
SELECT 
    name, database_id,
    (SELECT SUM(SIZE) FROM fs WHERE TYPE = 1 AND fs.database_id = db.database_id) LogFileSizeMB
FROM sys.databases db
WHERE database_id > 4
AND NAME LIKE @dbname;
 
SELECT @TSQLExec = CONCAT (
  @TSQLExec,
  'USE [',
  d.NAME,
  ']; CHECKPOINT; DBCC SHRINKFILE ([',
  f.NAME + ']) with no_infomsgs;' ,
  Char (13),Char (10))
FROM sys.databases d,
     sys.master_files f
WHERE d.database_id = f.database_id
      AND d.database_id > 4
      AND f.type = 1
      AND d.NAME LIKE @dbname;
PRINT @TSQLExec;
EXEC (@TSQLExec);
 
WITH fs
AS
(
    SELECT database_id, TYPE, SIZE * 8.0 / 1024 SIZE
    FROM sys.master_files
)
UPDATE a
set a.logFileSizeAfterMB = (SELECT SUM(SIZE) FROM fs WHERE TYPE = 1 AND fs.database_id = db.database_id)
FROM #temp a
inner join sys.databases db on a.dbid = db.database_id
WHERE database_id > 4
AND NAME LIKE @dbname
 
SELECT * FROM #temp ORDER BY dbname
SET NOCOUNT OFF;
END;
GO