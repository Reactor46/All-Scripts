#Creating an array begins with naming it and then populating it with values

$Beast_Wars_Transformers = @(
"Waspinator",
"Optimus Primal",
"Megatron",
"Dinobot",
"Rattrap",
"Rhinox",
"Scorponok",
"Tarantulis"
)

#The next line of code searches for values in the array that contain the letter 'r' and prints them to the console.

$Beast_Wars_Transformers | Where-Object { $PSItem -like "*r*" }