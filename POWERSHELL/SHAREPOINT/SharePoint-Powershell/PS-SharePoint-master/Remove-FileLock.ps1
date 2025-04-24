$web = get-spweb "http://srv-test-224.ponyex.local"
$list = $web.Lists["MyListName"]
$item = $list.Items.GetItemById(6)

$userId = $item.File.LockedByUser.ID
$user = $web.AllUsers.GetById($userId)

$impSite = New-Object Microsoft.SharePoint.SPSite($web.Url, $user.UserToken);
$impWeb = $impSite.OpenWeb();
$impList = $impWeb.Lists[$list.Title];
$impItem = $impList.GetItemById($item.ID);
$impItem.File.ReleaseLock($impItem.File.LockId)

$impWeb.Dispose()
$impSite.Dispose()
$web.Dispose()