# Save this as TestSubject.ps1
function Test-StringLength {
    param(
        [string]$inputString,
        [int]$minLength,
        [int]$maxLength
    )

    $length = $inputString.Length

    if ($length -ge $minLength -and $length -le $maxLength) {
        Write-Output "String length is within the expected range."
    } else {
        Write-Output "String length is outside the expected range."
    }
}
