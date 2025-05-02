a = "This is for testing"

for i = 1 to len(a)
 b = mid(a, i,1)
  if b <> " " then
   c = c & b
 end if
 
next
msgbox c
