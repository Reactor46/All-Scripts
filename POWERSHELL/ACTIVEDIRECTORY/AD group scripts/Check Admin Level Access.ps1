import-module activedirectory

$memberlist = new-item -type file -force “c:\scripts\outfile\ITadminMembers.csv”

import-csv “C:\scripts\AD group scripts\infile\AdminGroups.csv” | foreach-object {
	$gname = $_.samaccountname
	$group = get-adgroup $gname
	$group.name | out-file $memberlist -encoding unicode -append
		foreach ($member in get-adgroupmember -recursive $group) {$member.samaccountname | out-file $MemberList -encoding unicode -append}
$nl = [environment]::newline | out-file $memberlist -encoding ascii -append
}