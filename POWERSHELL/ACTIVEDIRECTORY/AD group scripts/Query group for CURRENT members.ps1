#change 'uson staff' and faxmaker-out accordingly
import-module activedirectory
get-adgroupmember -identity 'uson staff' -recursive | select SamAccountName | out-file “c:\scripts\outfile\uson-staff-out.csv”