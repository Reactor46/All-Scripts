function Test-Youtube(){
param($PingCount)

    Test-Connection youtube.com -Count $PingCount
}

Test-Youtube -PingCount 3