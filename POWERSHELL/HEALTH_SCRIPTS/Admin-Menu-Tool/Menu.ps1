$script:Invocation = (Get-Variable MyInvocation -Scope 0).Value
$script:Path = $Invocation.MyCommand.Path
$script:Folder = Split-Path -Parent $Invocation.MyCommand.Definition
$script:FileName = $Invocation.MyCommand.Name
$script:LogFile = $("$Folder\logs\" + [system.io.path]::getfilenamewithoutextension($Invocation.MyCommand.Name) + "-" + $(get-date -format yyyyMMdd) + ".txt")
$script:UseInfo = $($(get-date -format HH:mm:ss) + "`t" + $env:username + "`t")
#Set-Location $Folder
$script:TopSeperator = @"
╔$("═" * 44)╗
"@ 
$script:BottomSeperator = @"
╚$("═" * 44)╝
"@ 
$script:FileHeader = @"
$TopSeperator 
       ***** Log File Information *****
  Filename:   $FileName
  Created by: Stewart Basterash
              ATLED Consulting & Engineering
              stewart.basterash@hotmail.com
  Modified:   $(Get-Date -Date (get-item $Path).LastWriteTime -f MM-dd-yyyy)
$BottomSeperator
`r 
---------------- BEGIN LOG -----------------
`r 
"@ 

[xml]$xmldata = get-content ".\menu.xml"

#Title
$Script:Title = "ATLED PowerMenu"

#General Info
if ($xmldata.menu.GetAttribute("GeneralInfo") -ne "") {
  $Script:GeneralInfo = $xmldata.menu.GetAttribute("GeneralInfo") }
else { $Script:GeneralInfo = "CTRL+C will exit this menu... .\menu at the powershell promt to reexecute!" }

#User
[string]$Script:UserDisplay = [string]::Empty
[ADSI]$Script:User = $null

if ($xmldata.menu.GetAttribute("title") -ne "") {
  $Script:Title = $xmldata.menu.GetAttribute("title") }

# Black, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, DarkGray
# White, Blue,     Green,     Cyan,     Red,     Magenta,     Yellow,     Gray

#Default Colors
$Foreground = "Green"
$Background = "Black"
$FolderColor = "White"
$ExitColor = "Gray"
$MenuitemColor = "Yellow"
$DisableditemColor = "DarkYellow"
$InfoColor = "Magenta"
$PromptColor = "Yellow"

if ($xmldata.menu.GetAttribute("foreground") -ne "") {
  $Foreground = $xmldata.menu.GetAttribute("foreground") }

if ($xmldata.menu.GetAttribute("background") -ne "") {
  $Background = $xmldata.menu.GetAttribute("background") }

if ($xmldata.menu.GetAttribute("foldercolor") -ne "") {
  $FolderColor = $xmldata.menu.GetAttribute("foldercolor") }

if ($xmldata.menu.GetAttribute("exitcolor") -ne "") {
  $ExitColor = $xmldata.menu.GetAttribute("exitcolor") }

if ($xmldata.menu.GetAttribute("menuitemcolor") -ne "") {
  $MenuitemColor = $xmldata.menu.GetAttribute("menuitemcolor") }

if ($xmldata.menu.GetAttribute("disableditemcolor") -ne "") {
  $DisableditemColor = $xmldata.menu.GetAttribute("disableditemcolor") }

if ($xmldata.menu.GetAttribute("infocolor") -ne "") {
  $InfoColor = $xmldata.menu.GetAttribute("infocolor") }

if ($xmldata.menu.GetAttribute("promptcolor") -ne "") {
  $PromptColor = $xmldata.menu.GetAttribute("promptcolor") }
  
#Constants for Icons
$critical = 16
$question = 32
$exclamation = 48
$information = 64  

#Constants for Button Sets
$OKOnly = 0
$OKCancel = 1 
$AbortRetryIgnore = 2
$YesNoCancel = 3
$YesNo = 4
$RetryCancel = 5

#Constants for Default Button Locations
$DefaultButton1 = 0   # Left
$DefaultButton2 = 256 # Middle
$DefaultButton3 = 512 # Right
 
function Script:write-log([string]$entry)
{
  if (!(test-path "$Folder\logs" -pathtype container)) {
    new-item logs -type directory }
  if (!(test-path $LogFile -pathtype leaf)) {
    $FileHeader >> $LogFile 
    $entry >> $LogFile }
  else {
    $entry >> $LogFile }
}
 
Trap [Exception] {
  write-log $("$UseInfo`t$_. - Line:(" + $($_.InvocationInfo.ScriptLineNUmber)+":"+$($_.InvocationInfo.OffsetInLine)+ ") " + $($_.InvocationInfo.Line))
  #write-host $("$UseInfo`t$_. - Line:(" + $($_.InvocationInfo.ScriptLineNUmber)+":"+$($_.InvocationInfo.OffsetInLine)+ ") " + $($_.InvocationInfo.Line))
  continue
}

function Script:Show-Popup([string]$title,[string]$message,[int]$options)
{
  $a = new-object -comobject wscript.shell
  return ($a.popup($message,0,$title,$options))
}

function Script:Get-LoggedInUserInfo()
{
  $result = [string]::Empty
  if (IsDomainMember) {
    $dse = New-Object DirectoryServices.DirectoryEntry
    $searchRoot = $dse.adsPath
    $srch = New-Object DirectoryServices.DirectorySearcher( New-Object DirectoryServices.DirectoryEntry($searchRoot))
    $srch.Filter = "(&(objectClass=user)(objectCategory=Person)(samAccountName=$env:username))"
    $srch.PropertiesToLoad.Add('displayname') | Out-Null
    $srch.PageSize = 10000
    $srch.SearchScope = "Subtree"
    $usr = $srch.FindOne()
    $Script:DisplayName = ($usr).properties['displayname']
    $Script:User = ( New-Object DirectoryServices.DirectoryEntry($usr.Path) ) }
  else {
    $Script:User = [ADSI]"WinNT://$env:COMPUTERNAME/$env:USERNAME,user"
    $Script:DisplayName = $usr.PSBase.properties['fullname']
    if ($Script:DisplayName.length -eq 0) {
      $Script:DisplayName = $($env:UserName.ToUpper()) } }
}

function Script:IsDomainMember()
{
  return ((gwmi win32_computersystem).partofdomain)
}

function Script:Get-User([string]$username)
{
  $dse = New-Object DirectoryServices.DirectoryEntry
  $searchRoot = $dse.adsPath
  $srch = New-Object DirectoryServices.DirectorySearcher( New-Object DirectoryServices.DirectoryEntry($searchRoot))
  $srch.Filter = "(&(objectClass=user)(objectCategory=Person)(samAccountName=$username))"
  $srch.PageSize = 10000
  $srch.SearchScope = "Subtree"
  return ( New-Object DirectoryServices.DirectoryEntry($srch.FindOne().Path) )
}

function Script:Group-Exists([string]$group)
{
  $dse = New-Object DirectoryServices.DirectoryEntry
  $searchRoot = $dse.adsPath
  $srch = New-Object DirectoryServices.DirectorySearcher( New-Object DirectoryServices.DirectoryEntry($searchRoot))
  $srch.Filter = "(&(objectClass=group)(samAccountName=$group))"
  $srch.PageSize = 10000
  $srch.SearchScope = "Subtree"
  return ($srch.FindOne() -ne $null)
}

function Script:Get-Group([string]$group)
{
  $dse = New-Object DirectoryServices.DirectoryEntry
  $searchRoot = $dse.adsPath
  $srch = New-Object DirectoryServices.DirectorySearcher( New-Object DirectoryServices.DirectoryEntry($searchRoot))
  $srch.Filter = "(&(objectClass=group)(samAccountName=$group))"
  $srch.PageSize = 10000
  $srch.SearchScope = "Subtree"
  return ( New-Object DirectoryServices.DirectoryEntry($srch.FindOne().Path) )
}

function Script:Object-IsMemberOf($object, $group) # [DirectoryServices.DirectoryEntry]
{
  foreach($dn in $object.Properties["memberOf"]) 
  {
    if ($group.distinguishedName -eq $dn) {
      $result = $true 
      return $result
      break }
  }
  return $result
}

function Script:UpdateUI()
{
  MODE con:cols=110 lines=60
  $ui = (Get-Host).UI.RawUI
  $ui.ForegroundColor = $Foreground
  $ui.BackgroundColor = $Background
  $ui.WindowTitle     = $Title
}
 
function Script:Add-Exit($node)
{
  $e = $node.SelectSingleNode("./menuitem[@exititem='true']") 
  if (-not $e) {
    $e = $xmldata.CreateElement("menuitem") 
    $e.SetAttribute("id", $c)
    if (($node.ParentNode) -and ($node.ParentNode.GetType().Name -eq 'XmlDocument')) {
      $e.SetAttribute("display", "Exit Menu") }
    else {
      $e.SetAttribute("display", "Previous Menu") }  
    $e.SetAttribute("exititem", 'true')
    $node.AppendChild($e) | Out-Null 
  }
}

function Script:Remove-Exit($node)
{
  $e = $node.SelectSingleNode("./menuitem[@exititem='true']") 
  if ($e) {
    $node.RemoveChild($e) | out-null }
}

function Script:Menu-Header([string]$title)
{
  $w = (Get-Host).UI.RawUI.WindowSize.Width - 20
  $top = "`t╔$("═" * $w)╗"
  $bot = "`t╚$("═" * $w)╝"
  $lft = [int](($w -$title.Length) / 2)
  $rgt = $w - $title.Length - $lft
  $welcome = "";
  $welcome = "Welcome $DisplayName" # (Get-UserDisplay $([Environment]::UserName))
  $wlt = [int](($w -$welcome.Length) / 2)
  $wrt = $w - $welcome.Length - $wlt
  clear-host
  write-host $top -Foreground $Foreground
  write-host $("`t║$(" " * $wlt)$welcome$(" " * $wrt)║") -Foreground $Foreground
  write-host $("`t║$(" " * $lft)$title$(" " * $rgt)║") -Foreground $Foreground
  write-host $bot -Foreground $Foreground
  write-host "`t"
  write-host $("`tItem".PadRight(8) + "Option".PadRight(36) + "Description")
  write-host "`t"
}

function Script:Display-Menu($xmlnode)
{
  UpdateUI
  Menu-Header $Title
  Remove-Exit $xmlnode
  
  $c=0
  foreach ($node in $xmlnode.ChildNodes | Sort-Object name, display )
  { # ($User.Properties["memberOf"].value -match 'grp_private').Length
    if ((($node.Permissions -eq $null) -or ($node.Permissions -eq "")) -or (($node.Permissions.Length -gt 0) -and ($User.Properties["memberOf"].value -match $node.Permissions).Length -ne 0)) #((Group-Exists $node.Permissions) -and (Object-IsMemberOf ($Script:User) (Get-Group $node.Permissions) ))
    {
      $node.SetAttribute("id", $c)
      $a = ($c.ToString().PadLeft(2))
      switch ($node.Name)
      {
        "folder"   { write-host $("`t[" + $a + "]".PadRight(4) + $node.Display.ToString().PadRight(36) + $node.Description) -ForeGround $FolderColor }
        "menuitem" { 
          if ($node.GetAttribute("execute").Length -gt 0) {
            $node.SetAttribute("enabled", $true)
            write-host $("`t[" + $a + "]".PadRight(4) + $node.Display.ToString().PadRight(36) + $node.Description) -ForeGround $MenuitemColor }
          else {
            $node.SetAttribute("enabled", $false)
            write-host $("`t[" + $a + "]".PadRight(4) + $node.Display.ToString().PadRight(36) + $node.Description) -ForeGround $DisableditemColor }
        }    
      }
      $c++
    }  
  }

  Add-Exit $xmlnode
  write-host $("`t[" + $c.ToString().PadLeft(2) + "]".PadRight(4) + $xmlnode.LastChild.Display.ToString().PadRight(36) + $xmlnode.LastChild.Description) -ForeGround $ExitColor
  
  write-host "    "
  write-host "`t$Script:GeneralInfo" -Foreground $InfoColor
  write-host "    "
}

function Script:Execute-Script([string]$scriptname)
{
  if (($scriptname.Length -gt 0) -and (Test-Path $scriptname)) {
    if ($scriptname.EndsWith('.ps1')) {
      write-Log $("$useInfo`t" + $scriptname)
      Invoke-Expression $(".{.\"+$scriptname+"}") }
    else {
      write-Log $("$useInfo`t" + $scriptname)
      Invoke-Item $($scriptname) } 
  }
}

Get-LoggedInUserInfo
$current = $xmldata.menu

write-log $UseInfo
do {
  $select = $null
  do {
    Clear-Host
    Display-Menu $current
    [console]::ForegroundColor = $PromptColor
    $select = $null
    $select = Read-Host "`tSelect"
  }
  until (($select -match "\d") -and ($select -ge 0) -and ($select -lt $current.ChildNodes.Count))
  $selected = $current.SelectSingleNode("./*[@id='$select']")
  write-log $($UseInfo + " selected " + $selected.Name + " - " + $selected.Display)

  switch ($selected.Name)
  {
    "folder"   {  
      $current = $selected
      if (($current.GeneralInfo -ne $null) -and ($current.GeneralInfo -ne "")) {
        $Script.GeneralInfo = $current.GeneralInfo }
    }
    "menuitem" {  
      if ($selected.exititem) {
        $current = $selected.SelectSingleNode("../..") 
        if ($current.parentnode -eq $null) {
         clear-host
         write-host "`tThank you for using the $Title..." -Foreground $Foreground
         write-host "    "
         write-host "`tTo re-execute the menu simply type:  Powershell " -Foreground $Foreground
         write-host "    "
         exit-pssession
         exit }
      }
      else {
        if ($selected.Enabled -eq "True") {  
          write-host ""
          write-host $("`t< Currently Executing: {0} >" -f $selected.Display)
          write-host ""
          Execute-Script $selected.execute }
        else {
          Show-Popup "Item Disabled" "This item is not currently Active. Contact your Administrator!" ($information+$OKOnly+$DefaultButton1)
        }
      }    
    }
  }
}
until ( ($current.ParentNode -eq $null) -and ([int]$current.LastChild.id -eq $select) )
Exit-PSSession
Exit