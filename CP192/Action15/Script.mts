Call Loguearse
Call SeleccionarScore

Sub Loguearse()
			SystemUtil.CloseProcessByName "Chrome.exe"
			
			SystemUtil.Run "Chrome.exe", DataTable("e_URL", dtLocalSheet)
			While Browser("Equifax - Portal").Page("Equifax - Portal").WebEdit("txtLogin").Exist = false
				wait 1
			Wend
			
			Browser("Equifax - Portal").Page("Equifax - Portal").WebEdit("txtLogin").Set "TDP10" 
			Browser("Equifax - Portal").Page("Equifax - Portal").WebEdit("txtContrasenia").Set "N0Bl0qu3arP0rF@v0rX3"

			Browser("Equifax - Portal").Page("Equifax - Portal").WebElement("Ingresar   >>").Click

End Sub


Sub SeleccionarScore()
	While Browser("Equifax - Portal").Page("Equifax - Portal_2").Link("Telefonica Moviles-Residencial").Exist = False
		wait 1
	Wend
	Browser("Equifax - Portal").Page("Equifax - Portal_2").Link("Telefonica Moviles-Residencial").Click
	Set shell = CreateObject("Wscript.Shell")
shell.SendKeys "{ENTER}"
	While Browser("Interconnect").Page("Interconnect").WebButton("Registro BD Excepciones").Exist = False
		wait 1
	Wend
	Browser("Interconnect").Page("Interconnect").WebButton("Registro BD Excepciones").Click
	While Browser("Interconnect").Page("Interconnect").WebList("select").Exist = false
		wait 1
	Wend
	Browser("Interconnect").Page("Interconnect").WebList("select").Select "Carnét de Extranjería"
	wait 3
	
Browser("Interconnect").Page("Interconnect").WebEdit("Score").Set "9992"
Browser("Interconnect").Page("Interconnect").WebEdit("Nro de Líneas Disponibles").Set "90"
Set shell = CreateObject("Wscript.Shell")
shell.SendKeys "{TAB}"
wait 1
shell.SendKeys "{TAB}"
wait 1
shell.SendKeys "{TAB}"
wait 1
shell.SendKeys "{TAB}"
wait 1
shell.SendKeys "{TAB}"
wait 1
shell.SendKeys "{TAB}"
wait 1
shell.SendKeys "970706530"
wait 1
Browser("Interconnect").Page("Interconnect").WebElement("WebElement").Click
wait 1
Browser("Interconnect").Page("Interconnect").WebButton("Registrar").Click

Browser("Interconnect").Page("Interconnect").WebButton("Limpiar").Click
Browser("Interconnect").Page("Interconnect").WebElement("×").Click
Browser("Interconnect").Close



End Sub









