$farm = get-spfarm
$farm.properties.remove("NintexWorkflowServer2013License")
$farm.properties.remove("NW2007ConfigurationDatabase ")
$farm.update()