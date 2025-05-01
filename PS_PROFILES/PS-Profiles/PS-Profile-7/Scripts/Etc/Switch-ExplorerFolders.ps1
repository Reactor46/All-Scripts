param(
    [ValidateSet('Show', 'Hide')]
	[string]
    $Switch
)
$list = "{f42ee2d3-909f-4907-8871-4c22fc0bf756}", 
        "{0ddd015d-b06c-45d5-8c4c-f59713854639}", 
        "{35286a68-3c57-41a1-bbb1-0eae73d76c95}", 
        "{7d83ee9b-2244-4e70-b1f5-5393042af1e4}", 
        "{a0c69a99-21c8-4671-8703-7934162fcf1d}", 
        "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"

foreach ( $item in $list )
    {
    $path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\$item\PropertyBag"
    Get-ItemProperty $Path | Set-ItemProperty -Name ThisPCPolicy -Value $Switch
    }