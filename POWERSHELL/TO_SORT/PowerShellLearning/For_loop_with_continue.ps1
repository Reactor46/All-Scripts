for($a=2; $a -le 12; $a++){
    if($a % 2 -eq 0){
    continue
    }
    Write-Output $a
}