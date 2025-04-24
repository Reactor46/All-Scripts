function Test-Youtube(){
[CmdletBinding()]
param(
[Parameter(Mandatory)]
[Int32]$PingCount

)

    Test-Connection youtube.com -Count $PingCount
}

Test-Youtube -PingCount 3