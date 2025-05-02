#Creating the array
$GilmoreGirls = @(
'Lorelai',
'Rory',
'Emily'
)

#Displaying the third array item on its own
$GilmoreGirls[2]

#Redefining element number 3:
$GilmoreGirls[2] = 'Jess'

#Displaying the third array item on its own after I changed its value to Jess
$GilmoreGirls[2]

#Adding a fourth element
$GilmoreGirls += 'Luke'

#Display full array
$GilmoreGirls