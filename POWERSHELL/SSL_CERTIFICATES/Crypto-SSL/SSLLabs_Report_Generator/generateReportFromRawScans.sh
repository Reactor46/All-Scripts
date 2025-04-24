#!/bin/bash
echo "Domain,TLS Version,Cipher Suite,Code,Forward Secrecy,Weak,SSLLabs Grade" > report.txt
for file in *scanNumber*.txt; do 

domain=$(awk 'NR==1' "$file" | strings)
grade=$(awk 'NR==2' "$file")

echo "Domain: ${domain}"
echo "Grade: ${grade}"
filename=$(echo $filename)


sed -n '/Cipher Suites/,/Handshake Simulation/p' $file | awk '!p;/Handshake Simulation/{p=1}' | strings | awk 'BEGIN{RS="TLS" ;FS="\n"} ; {print "TLS"$1,$2}' | awk '{print $1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11}' | awk '/FS/ {print "FS",$0;next}{print "NO-FS",$0}' | awk '/WEAK/ {print "WEAK",$0;next}{print "NOT-WEAK",$0}' | awk '/TLS 1\.3/ {TLS_VER="TLS1\.3";next } /TLS 1\.2/ {TLS_VER="TLS1\.2";next } /TLS 1\.1/ {TLS_VER="TLS1\.1";next } /TLS 1\.0/ {TLS_VER="TLS1\.0";next} {print TLS_VER,$0}' > readyForCSV-${filename}

awk '{OFS=","}{print "'$domain'",$1,$4,$5,$3,$2,"'$grade'"}' readyForCSV-${filename} | strings | awk '{if (NR!=1) {print}}'  >> report.txt

rm readyForCSV-${filename}

done
