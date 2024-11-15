/****** Script Clear Data for Inventory  ******/
/** USE ServerInventory; **/
/** DELETE FROM dbo.tbAdminShare **/
/** WHERE InventoryDate < DATEADD(dd,-30,GETDATE());  Delete data older than 30days **/ 
/** WHERE InventoryDate < DATEADD(hh,-12,GETDATE());  Delete data older than 12hrs  **/ 
/** WHERE InventoryDate < DATEADD(mi,-15,GETDATE());  Delete data older than 15mins **/ 
/** DBCC SHRINKDATABASE(N'ServerInventory' ) **/
/**		**/
/** GO **/
/** Delete All data from All Tables **/

USE [ServerInventory]

GO

TRUNCATE TABLE ServerInventory.dbo.tbAdminShare;
TRUNCATE TABLE ServerInventory.dbo.tbAntiVirus;
TRUNCATE TABLE ServerInventory.dbo.tbDrives;
TRUNCATE TABLE ServerInventory.dbo.tbGeneral;
TRUNCATE TABLE ServerInventory.dbo.tbGroups;
TRUNCATE TABLE ServerInventory.dbo.tbMemory;
TRUNCATE TABLE ServerInventory.dbo.tbNetwork;
TRUNCATE TABLE ServerInventory.dbo.tbOperatingSystem;
TRUNCATE TABLE ServerInventory.dbo.tbProcessor;
TRUNCATE TABLE ServerInventory.dbo.tbScheduledTasks;
TRUNCATE TABLE ServerInventory.dbo.tbServerRoles;
TRUNCATE TABLE ServerInventory.dbo.tbServices;
TRUNCATE TABLE ServerInventory.dbo.tbSoftware;
TRUNCATE TABLE ServerInventory.dbo.tbUpdates;
TRUNCATE TABLE ServerInventory.dbo.tbUsers;
TRUNCATE TABLE ServerInventory.dbo.tbUserShare;

DBCC SHRINKDATABASE(N'ServerInventory' )

GO
