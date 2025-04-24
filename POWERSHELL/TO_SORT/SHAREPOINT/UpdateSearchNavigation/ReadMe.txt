Name
.\UpdateSearchNavigation.ps1


Description

This script copys the search navigation nodes of the root site of the site collection url you provide to all of its subsites.
If the -RemoveExisting parameter is defined as $true, the existing search navigation nodes will be removed and repleaced.


Syntax
.\UpdateSearchNavigation.ps1 -Scope <"site" or "webapplication"> -URL <the root site collection/webapplication url> [-RemoveExisting <$true or $false>]

-RemoveExisting is optional and defined as $false by default


Examples

.\UpdateSearchNavigation.ps1 -Scope site -RemoveExisting $true -URL "http://www.mysharepoint.net"
This command will remove the existing nodes and updates all the subsites in the site collections

Author
Kampan - www.sharepointloupe.net