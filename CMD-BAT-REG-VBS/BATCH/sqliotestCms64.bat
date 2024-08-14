@echo off
rem SQLIO test C
rem Basic tests running 
rem FIXED
rem Duration = 10s 
rem Block Size = 64K
rem Threads 4 aus -Fparam.txt
rem Files 2
rem File Size each 10 GB
rem VARIABLE
rem Outstanding IOs = 0,1,2,4,8,16

echo Outstanding IOs, 0, 1,2,4,8,16,32,64,128,256 - 2x10 GB file run for 10 seconds each.


echo Test for RANDOM WRITE
call sqlio -kW -s30 -frandom -o0 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -frandom -o1 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -frandom -o2 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -frandom -o4 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -frandom -o8 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -frandom -o16 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -frandom -o32 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -frandom -o64 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -frandom -o128 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -frandom -o256 -b64 -LS -Fparam.txt

echo repeat for RANDOM READ
call sqlio -kR -s30 -frandom -o0 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -frandom -o1 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -frandom -o2 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -frandom -o4 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -frandom -o8 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -frandom -o16 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -frandom -o32 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -frandom -o64 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -frandom -o128 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -frandom -o256 -b64 -LS -Fparam.txt

echo repeat for SEQUENTIAL READ
call sqlio -kR -s30 -fsequential -o0 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -fsequential -o1 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -fsequential -o2 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -fsequential -o4 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -fsequential -o8 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -fsequential -o16 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -fsequential -o32 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -fsequential -o64 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -fsequential -o128 -b64 -LS -Fparam.txt
call sqlio -kR -s30 -fsequential -o256 -b64 -LS -Fparam.txt

echo repeat for SEQUENTIAL WRITE
call sqlio -kW -s30 -fsequential -o0 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -fsequential -o1 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -fsequential -o2 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -fsequential -o4 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -fsequential -o8 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -fsequential -o16 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -fsequential -o32 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -fsequential -o64 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -fsequential -o128 -b64 -LS -Fparam.txt
call sqlio -kW -s30 -fsequential -o256 -b64 -LS -Fparam.txt

goto Done

:Done
echo Test C Complete

rem Test C COMPLETE