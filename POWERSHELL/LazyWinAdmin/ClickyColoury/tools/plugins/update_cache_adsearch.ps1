# api: multitool
# version: 0.1
# title: update ADSearch.txt
# description: Update ADSearch cache file
# type: inline
# category: update
# hidden: 1
# img: log
# status: obsolete
#
# Update data\adsearch.txt
#
#  - used by WPF multiTool user search
#  - is a plain line-wise text file of format:
#    "ADUser | Name, USer | +1-234-567890"
#

$cache_fn = ".\data\adsearch.txt"

echo "Get-ADUser ... > $cache_fn"

Get-ADUser -Filter * -Properties SAMAccountName,DisplayName,TelephoneNumber | % {
   "{0} | {1} | {2}" -f @($_.SAMAccountName, $_.DisplayName, $_.TelephoneNumber)
} | Out-File $cache_fn -Encoding UTF8

