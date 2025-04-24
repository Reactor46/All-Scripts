#Defining a variable
$beast_wars = 'Megatron is all powerful'

#If statement to see if the value 'Megatron' matches in any way with what is in the variable
If($beast_wars -match 'Megatron'){
    Write-Host 'There is a match'
} else {
    Write-Host 'No Match'
}