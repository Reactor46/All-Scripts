while (1) {Get-Process | sort -desc cpu | select -first 100 ; 
sleep -seconds 3; cls; 

write-host "Handles  NPM(K)    PM(K)      WS(K) VM(M)   CPU(s)     Id ProcessName"; 
write-host "-------  ------    -----      ----- -----   ------     -- -----------"
}
