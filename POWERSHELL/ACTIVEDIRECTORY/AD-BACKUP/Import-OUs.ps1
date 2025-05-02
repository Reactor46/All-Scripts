#################################################################################### 
# Bulk Import of Organizational Units from CSV 
#################################################################################### 
# Created by Brad Voris 
# Description: Bulk import of Active Directory OUs via Powershell with a CSV file 
#################################################################################### 
# Notes for CSV File 
# First line of CSV file should contain the following in each cell 
# NAME,DistinguishedName 
# 
#################################################################################### 
#################################################################################### 
#Import AD Module RSAT must be installed 
#################################################################################### 
Import-Module ActiveDirectory  
#################################################################################### 
#Varibale location for CSV file 
#################################################################################### 
$ous = Import-Csv -Path "C:\LazyWinAdmin\ActiveDirectory\AD-BACKUP\Las-Vegas - Testing OU.csv"   
#################################################################################### 
# For each function to create accounts 
#################################################################################### 
foreach ($ou in $ous)   
{   
#################################################################################### 
# Function Variables 
#################################################################################### 
    $ouname = $ou.name   
    $oudn = $ou.DistinguishedName  
     
#################################################################################### 
# Function 
#################################################################################### 
    New-ADOrganizationalUnit -Name $ouname -Path $oudn  -ManagedBy 'domain admins' 
}  