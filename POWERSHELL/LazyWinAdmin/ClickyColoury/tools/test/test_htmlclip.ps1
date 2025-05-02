# api: multitool
# title: test html clip
# description: Set-clipboardHtml
# version: 0.2
# depends: clipboard
# category: test
# type: inline
#
# testy test

$email ="test@test"
$user = "user123"
Set-ClipboardHtml "<p><font style='color:green'>&#x2714;➟</font> The mailbox &lt;<font style='color:blue'>$email</font>&gt; for the account <b style='color:#553311'>$user</b> has been created.</p>"
