If Browser("Entrar").Exist Then
	Browser("Entrar").CloseAllTabs
End If


Dim str_usuario, str_password

str_usuario = DataTable("e_Usuario", dtLocalsheet)
str_password = DataTable("e_Password", dtLocalsheet)
str_nroporta = DataTable("e_NroPorta", dtLocalsheet)
str_fechaActivacion = DataTable("e_FechaActivación", dtLocalsheet) 
str_fechaTermino = DataTable("e_FechaTerminoContrato", dtLocalsheet) 

SystemUtil.Run "chrome.exe", "http://10.117.1.46/Portaflow/faces/login.jsf"

While(Browser("Entrar").Page("Entrar").WebEdit("formLogin:user").Exist) = False
	wait 1
Wend

Browser("Entrar").Page("Entrar").WebEdit("formLogin:user").Set str_usuario
wait 1
Browser("Entrar").Page("Entrar").WebEdit("formLogin:password").Set str_password
wait 1
Browser("Entrar").CaptureBitmap RutaEvidencias() & "LoginPortaflow"&Num_Iter&".png", True
imagenToWord "Login Portaflow", RutaEvidencias() &"LoginPortaflow"&Num_Iter&".png"
Browser("Entrar").Page("Entrar").WebButton("formLogin:loginButton").Click
wait 1

While(Browser("Entrar").Page("Portaflow").WebElement("Portaflow").Exist) = False
	wait 1
Wend

Browser("Entrar").Page("Portaflow").WebElement("menuform:j_idt16").HoverTap
wait 1
Browser("Entrar").Page("Portaflow").Link("Envío de mensajes").Click
wait 1
Browser("Entrar").CaptureBitmap RutaEvidencias() & "RolCedente"&Num_Iter&".png", True
imagenToWord "Rol Cedente", RutaEvidencias() &"RolCedente"&Num_Iter&".png"
Browser("Entrar").Page("Portaflow").Link("Rol Cedente").Click
While(Browser("Rol Cedente").Page("Rol Cedente").WebElement("Buscar").Exist) = False
	wait 1
Wend
wait 2
Browser("Rol Cedente").Page("Rol Cedente").WebElement("WebElement").Click
Set shell = CreateObject("Wscript.Shell") 
shell.SendKeys "{DOWN}"
wait 1
shell.SendKeys "{ENTER}"
wait 1
While(Browser("Entrar").Page("Rol Cedente").WebEdit("frmQuery:number").Exist) = False
	wait 1
Wend
wait 1
Browser("Entrar").Page("Rol Cedente").WebEdit("frmQuery:number").Set str_nroporta
wait 1
Browser("Rol Cedente").Page("Rol Cedente").WebElement("Buscar").Click
wait 1
Browser("Entrar").CaptureBitmap RutaEvidencias() & "ConsultaPrevia"&Num_Iter&".png", True
imagenToWord "Consulta Previa", RutaEvidencias() &"ConsultaPrevia"&Num_Iter&".png"
Browser("Rol Cedente").Page("Rol Cedente").WebButton("frmConsolaCedente:consolaCeden").Click
wait 1

While(Browser("Rol Cedente").Page("Rol Cedente").WebButton("Show Calendar").Exist) = False
	wait 1
Wend

Browser("Rol Cedente").Page("Rol Cedente").WebButton("Show Calendar").Click
wait 1
Select Case str_fechaActivacion
	Case "1"
		Browser("Rol Cedente").Page("Rol Cedente").Link("1").Click
	Case "2"
		Browser("Rol Cedente").Page("Rol Cedente").Link("2").Click
	Case "3"
		Browser("Rol Cedente").Page("Rol Cedente").Link("3").Click
	Case "4"
		Browser("Rol Cedente").Page("Rol Cedente").Link("4").Click
	Case "5"
		Browser("Rol Cedente").Page("Rol Cedente").Link("5").Click
	Case "6"
		Browser("Rol Cedente").Page("Rol Cedente").Link("6").Click
	Case "7"
		Browser("Rol Cedente").Page("Rol Cedente").Link("7").Click
	Case "8"
		Browser("Rol Cedente").Page("Rol Cedente").Link("8").Click
	Case "9"
		Browser("Rol Cedente").Page("Rol Cedente").Link("9").Click
	Case "10"
		Browser("Rol Cedente").Page("Rol Cedente").Link("10").Click
	Case "11"
		Browser("Rol Cedente").Page("Rol Cedente").Link("11").Click
	Case "12"
		Browser("Rol Cedente").Page("Rol Cedente").Link("12").Click
	Case "13"
		Browser("Rol Cedente").Page("Rol Cedente").Link("13").Click
	Case "14"
		Browser("Rol Cedente").Page("Rol Cedente").Link("14").Click
	Case "15"
		Browser("Rol Cedente").Page("Rol Cedente").Link("15").Click
	Case "16"
		Browser("Rol Cedente").Page("Rol Cedente").Link("16").Click
	Case "17"
		Browser("Rol Cedente").Page("Rol Cedente").Link("17").Click
	Case "18"
		Browser("Rol Cedente").Page("Rol Cedente").Link("18").Click
	Case "19"
		Browser("Rol Cedente").Page("Rol Cedente").Link("19").Click
	Case "20"
		Browser("Rol Cedente").Page("Rol Cedente").Link("20").Click
	Case "21"
		Browser("Rol Cedente").Page("Rol Cedente").Link("21").Click
	Case "22"
		Browser("Rol Cedente").Page("Rol Cedente").Link("22").Click
	Case "23"
		Browser("Rol Cedente").Page("Rol Cedente").Link("23").Click
	Case "24"
		Browser("Rol Cedente").Page("Rol Cedente").Link("24").Click
	Case "25"
		Browser("Rol Cedente").Page("Rol Cedente").Link("25").Click
	Case "26"
		Browser("Rol Cedente").Page("Rol Cedente").Link("26").Click
	Case "27"
		Browser("Rol Cedente").Page("Rol Cedente").Link("27").Click
	Case "28"
		Browser("Rol Cedente").Page("Rol Cedente").Link("28").Click
	Case "29"
		Browser("Rol Cedente").Page("Rol Cedente").Link("29").Click
	Case "30"
		Browser("Rol Cedente").Page("Rol Cedente").Link("30").Click
	Case "31"
		Browser("Rol Cedente").Page("Rol Cedente").Link("31").Click
End Select
wait 1
Browser("Rol Cedente").Page("Rol Cedente").WebButton("Show Calendar").Click
wait 1
Browser("Rol Cedente").Page("Rol Cedente").WebButton("Show Calendar_2").Click
wait 1
Select Case str_fechaTermino
	Case "1"
		Browser("Rol Cedente").Page("Rol Cedente").Link("1").Click
	Case "2"
		Browser("Rol Cedente").Page("Rol Cedente").Link("2").Click
	Case "3"
		Browser("Rol Cedente").Page("Rol Cedente").Link("3").Click
	Case "4"
		Browser("Rol Cedente").Page("Rol Cedente").Link("4").Click
	Case "5"
		Browser("Rol Cedente").Page("Rol Cedente").Link("5").Click
	Case "6"
		Browser("Rol Cedente").Page("Rol Cedente").Link("6").Click
	Case "7"
		Browser("Rol Cedente").Page("Rol Cedente").Link("7").Click
	Case "8"
		Browser("Rol Cedente").Page("Rol Cedente").Link("8").Click
	Case "9"
		Browser("Rol Cedente").Page("Rol Cedente").Link("9").Click
	Case "10"
		Browser("Rol Cedente").Page("Rol Cedente").Link("10").Click
	Case "11"
		Browser("Rol Cedente").Page("Rol Cedente").Link("11").Click
	Case "12"
		Browser("Rol Cedente").Page("Rol Cedente").Link("12").Click
	Case "13"
		Browser("Rol Cedente").Page("Rol Cedente").Link("13").Click
	Case "14"
		Browser("Rol Cedente").Page("Rol Cedente").Link("14").Click
	Case "15"
		Browser("Rol Cedente").Page("Rol Cedente").Link("15").Click
	Case "16"
		Browser("Rol Cedente").Page("Rol Cedente").Link("16").Click
	Case "17"
		Browser("Rol Cedente").Page("Rol Cedente").Link("17").Click
	Case "18"
		Browser("Rol Cedente").Page("Rol Cedente").Link("18").Click
	Case "19"
		Browser("Rol Cedente").Page("Rol Cedente").Link("19").Click
	Case "20"
		Browser("Rol Cedente").Page("Rol Cedente").Link("20").Click
	Case "21"
		Browser("Rol Cedente").Page("Rol Cedente").Link("21").Click
	Case "22"
		Browser("Rol Cedente").Page("Rol Cedente").Link("22").Click
	Case "23"
		Browser("Rol Cedente").Page("Rol Cedente").Link("23").Click
	Case "24"
		Browser("Rol Cedente").Page("Rol Cedente").Link("24").Click
	Case "25"
		Browser("Rol Cedente").Page("Rol Cedente").Link("25").Click
	Case "26"
		Browser("Rol Cedente").Page("Rol Cedente").Link("26").Click
	Case "27"
		Browser("Rol Cedente").Page("Rol Cedente").Link("27").Click
	Case "28"
		Browser("Rol Cedente").Page("Rol Cedente").Link("28").Click
	Case "29"
		Browser("Rol Cedente").Page("Rol Cedente").Link("29").Click
	Case "30"
		Browser("Rol Cedente").Page("Rol Cedente").Link("30").Click
	Case "31"
		Browser("Rol Cedente").Page("Rol Cedente").Link("31").Click
End Select
wait 1
Browser("Rol Cedente").Page("Rol Cedente").WebEdit("formRespuesta1:observaciones").Click
wait 1
Browser("Rol Cedente").Page("Rol Cedente").WebEdit("formRespuesta1:observaciones").Set "QA"
Browser("Rol Cedente").CaptureBitmap RutaEvidencias() & "FechaRegistro"&Num_Iter&".png", True
imagenToWord "Fecha de Registro", RutaEvidencias() &"FechaRegistro"&Num_Iter&".png"
wait 1
Browser("Rol Cedente").Page("Rol Cedente").WebElement("Aceptar").Click
wait 1

While(Browser("Rol Cedente").Page("Rol Cedente").WebElement("Buscar").Exist) = False
	wait 1
Wend
wait 1
Browser("Rol Cedente").CaptureBitmap RutaEvidencias() & "ConsultaExitosaECPC"&Num_Iter&".png", True
imagenToWord "Consulta Exitosa ECPC", RutaEvidencias() &"ConsultaExitosaECPC"&Num_Iter&".png"
Browser("Entrar").Page("Rol Cedente").WebButton("frmConsolaCedente:consolaCeden").Click
wait 1
While(Browser("Entrar").Page("Rol Cedente").WebTable("Tipo de mensaje").Exist) = False
	wait 1
Wend
Browser("Rol Cedente").CaptureBitmap RutaEvidencias() & "DetalleECPC"&Num_Iter&".png", True
imagenToWord "Detalle del mensaje - ECPC", RutaEvidencias() &"DetalleECPC"&Num_Iter&".png"
Browser("Entrar").Page("Rol Cedente").WebButton("Close").Click
wait 1
Browser("Entrar").Page("Rol Cedente").Link("Jorge Luis Ramos Velarde").Click @@ script infofile_;_ZIP::ssf7.xml_;_
wait 1
Browser("Entrar").Page("Rol Cedente").Image("avatar.png").Click @@ script infofile_;_ZIP::ssf8.xml_;_
wait 1
Browser("Entrar").Page("Rol Cedente").Link("Salir").Click @@ script infofile_;_ZIP::ssf9.xml_;_
Browser("Entrar").CloseAllTabs






