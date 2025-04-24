#Storing the text in this txt file in a variable:
$lines = Get-Content -Path 'C:\Users\silad\Documents\text.txt'

#The 2 lines of text in the file can be treated as array object with [0] representing line, so to store the 2nd line as a var, I need to refer to it as [1].
$2nd_line = $lines[1]

#Displaying the 2nd line
$2nd_line