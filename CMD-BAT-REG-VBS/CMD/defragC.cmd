::
:: Polimorphic Defrag
:: Defragmentation based on the last character of the file name
:: EX.: change the name to defragK.cmd 
::      and K: will be defragmented
:: 
::
:: Gas - sett 2007 v.02
:: windows 2003 / windows 2008 / windows 8.1
@ECHO off
setlocal

set filename=%~n0
set logfile="%temp%\_%filename%.txt"
:: *** get the last char
set disk=%filename:~-1%
echo ---------START-%time%--%date%----- >>%logfile%

defrag %disk%: >>%logfile% 2>&1

echo -----------END-%time%--%date%----- >>%logfile%
