strAction = lcase(wscript.arguments(0))
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
strresdep1 = ListServices(objWMIService,"IIS Admin")
strresdep2 = ListServices(objWMIService,"Microsoft Exchange")

function ListServices(objWMIService,Strservicename)

Set colListOfServices = objWMIService.ExecQuery _
        ("Select * from Win32_Service")

For Each objService in colListOfServices
	if instr(objService.Displayname,Strservicename) then
		if objService.startmode = "Auto" then strdepres = listdepservice(objWMIService,objService.Name)
		wscript.echo
		wscript.echo "Service Name: " & objService.displayname & " Current State: " &  objService.state
		wscript.echo "Statup:" & objService.startmode
		Wscript.echo
		select case strAction
			case "stop"
				if objService.state = "Running" then strsrest = StartStopService(strAction,objService)
			case "start" 
				if objService.startmode = "Auto" then strsrest = StartStopService(strAction,objService)
		end select
	end if
next

end function

function listdepservice(objWMIService,Strservicename)

Set colListOfDepServices = objWMIService.ExecQuery _
        ("Associators of " _
     & "{Win32_Service.Name='" & Strservicename & "'} Where " _
     & "AssocClass=Win32_DependentService Role=Antecedent" )

For Each objService in colListOfDepServices
	strdepres1 = listdepservice(objWMIService,objService.Name)
	wscript.echo
	wscript.echo "Dependant Service"
	wscript.echo "Service Name: " & objService.displayname & " Current State: " &  objService.state
	wscript.echo "Statup:" & objService.startmode
	Wscript.echo
	select case strAction
		case "stop"
			if objService.state = "Running" then strsrest = StartStopService(strAction,objService)
		case "start" 
			if objService.startmode = "Auto" then strsrest = StartStopService(strAction,objService)
	end select
next

end function

function StartStopService(strAction,objService)

select case strAction
	case "stop" 
		intres = objService.stopservice
	 	select case intres
			case 0 wscript.echo "Stopping Service The request was accepted"
	 		case 5 wscript.echo "Service is Stopped"
			case else wscript.echo "Error Stopping Service Error Code " & intres
       		end select
		if objService.InterrogateService = 4 then
			intex = 0
			intlcount = 0
			while intex = 0
				if objService.InterrogateService = 6 then
					intex = 1
				else
					Wscript.echo "Waiting for Service to stop"
					wscript.sleep 1000
					intlcount = intlcount + 1
				end if
				if intlcount = 120 then
					Wscript.echo "Service didn't stop within 2 Minutes"
					intex = 1
				end if
			wend
			if objService.InterrogateService = 6 then wscript.echo  "Service is Stopped"
		end if
         case "start" objService.startservice
		intres = objService.startservice
		select case intres
			case 0 wscript.echo "Starting Service The request was accepted"
	 		case 10 wscript.echo "Service is Started"
			case else wscript.echo "Error Starting Service Error Code " & intres
       		end select
		if objService.InterrogateService = 4 then
			intex = 0
			intlcount = 0
			while intex = 0
				if objService.InterrogateService = 0 then
					intex = 1
				else
					Wscript.echo "Waiting for Service to start"
					wscript.sleep 1000
					intlcount = intlcount + 1
				end if
				if intlcount = 120 then
					Wscript.echo "Service didn't start within 2 Minutes"
					intex = 1
				end if
			wend
			if objService.InterrogateService = 0 then wscript.echo  "Service is Started"	
		end if
end select

end function