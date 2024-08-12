@ECHO OFF
SET SCRIPT=C:\LazyWinAdmin\ServerHealthDashboard\ServerHealth-TESTING.ps1
SET OUTPUT=C:\LazyWinAdmin\ServerHealthDashboard\Output\

SET LISTCAS=C:\LazyWinAdmin\ServerHealthDashboard\Configs\CASSrvs.txt
SET OFNCAS=ServerStatusReportCASSrvs.htm

SET LISTCASMT=C:\LazyWinAdmin\ServerHealthDashboard\Configs\CASMTSrvs.txt
SET OFNCASMT=ServerStatusReportCASMTSrvs.htm

SET LISTCOLL=C:\LazyWinAdmin\ServerHealthDashboard\Configs\COLLSrvs.txt
SET OFNCOLL=ServerStatusReportCOLLSrvs.htm

SET LISTMT=C:\LazyWinAdmin\ServerHealthDashboard\Configs\MTSrvs.txt
SET OFNMT=ServerStatusReportMTSrvs.htm

SET LISTTEST=C:\LazyWinAdmin\ServerHealthDashboard\Configs\Test.txt
SET OFNTEST=ServerStatusReportTESTSrvs.htm

SET THRESHOLD=-CpuAlertThreshold 80 -MemAlertThreshold 80 -RequestAlertThreshold 30 -ServiceMemThreshold 1800 -TotalWebThreshold 250 -TCPConThreshold 6000

powershell.exe "%SCRIPT% -List %LISTTEST% -O %OUTPUT% -OFN %OFNTEST% %THRESHOLD%"
REM powershell.exe "%SCRIPT% -List %LISTCAS% -O %OUTPUT%  -OFN %OFNCAS%  %THRESHOLD%"
REM powershell.exe "%SCRIPT% -List %LISTCASMT% -O %OUTPUT% -OFN %OFNCASMT% %THRESHOLD%"
REM powershell.exe "%SCRIPT% -List %LISTCOLL% -O %OUTPUT% -OFN %OFNCOLL% %THRESHOLD%"
REM powershell.exe "%SCRIPT% -List %LISTMT% -O %OUTPUT% -OFN %OFNMT% %THRESHOLD%"