# Save this as TestScript.ps1
. .\TestSubject.ps1

# Test Case 1: String length within the expected range
Test-StringLength -inputString "Hello, PowerShell!" -minLength 5 -maxLength 20

# Test Case 2: String length outside the expected range
Test-StringLength -inputString "Short" -minLength 10 -maxLength 20

