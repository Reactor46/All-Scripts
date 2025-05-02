# api: ps
# type: functions
# title: clipboard
# description: Reads/writes to clipboard
# version: 0.8
# status: stable
# category: win32
# config: -
#
# Provides clipboard read and write shortcuts, with an HTML formatter.
#

# syslib
Add-Type -AN System.Windows.Forms


#-- Send HTML formatted content to clipboard
function Set-ClipboardHtml($html) {
  #-- escape non-ASCII characters first
  $html = [regex]::replace($html, "([^\x01-\x7F])", {
      "&#" + ([int][char][string]$args[0]) + ";"
  })
  #-- wrap with HTML Format header
  $len = $html.length;  # section sizes
  $pfx = 190; $start = 165; $end = 35;
  $data = @"
Version:1.0
StartHTML:$(($pfx).toString().PadLeft(9, '0'))
EndHTML:$(($pfx+$start+$len+$end).toString().PadLeft(9, '0'))
StartFragment:$(($pfx+$start).toString().PadLeft(9, '0'))
EndFragment:$(($pfx+$start+$len).toString().PadLeft(9, '0'))
StartSelection:$(($pfx+$start).toString().PadLeft(9, '0'))
EndSelection:$(($pfx+$start+$len).toString().PadLeft(9, '0'))
SourceURL:http://localhost/#cmd-table
   <!DOCTYPE html><HTML> 
   <HEAD><TITLE>Clibboard</TITLE><META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=UTF-8'></HEAD>
   <BODY><!--Start-->      
$html
   <!--End--></BODY> </HTML>      
"@
  #-- send to system clipboard
  [System.Windows.Forms.Clipboard]::setData([System.Windows.Forms.Dataformats]::Html, $data)
}

#-- Write clipboard (text/plain)
function Set-Clipboard($text) {
    if ($text.length) {
        [void] [System.Windows.Forms.Clipboard]::setText($text)
    }
}

#-- Return current content (text/plain)
function Get-Clipboard() {
    return [System.Windows.Forms.Clipboard]::getText().trim()
}

<#
 #@tests
 return [System.Windows.Forms.Clipboard]::GetDataObject().getFormats()
#>
