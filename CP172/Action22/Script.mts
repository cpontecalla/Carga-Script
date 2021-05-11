Call Loguearse
Call SeleccionarScore

Sub Loguearse()
			SystemUtil.CloseProcessByName "Chrome.exe"
			
			SystemUtil.Run "Chrome.exe", DataTable("e_URL", dtLocalSheet)
'			While Browser("Equifax - Portal").InsightObject("InsightObject").Exist = false
'				wait 1
'			Wend
'			Browser("Equifax - Portal").InsightObject("InsightObject").Click

			
			wait 2
			Set shell = CreateObject("Wscript.Shell") 
			shell.SendKeys DataTable("e_Usuario",dtLocalSheet)
			wait 2
			Browser("Equifax - Portal").InsightObject("InsightObject_2").Click
			wait 2
			Set shell = CreateObject("Wscript.Shell") 
			shell.SendKeys DataTable("e_Contraseña",dtLocalSheet)
			wait 2
			Browser("Equifax - Portal").CaptureBitmap RutaEvidencias() & "Login.png", True
			imagenToWord "Login Equifax:", RutaEvidencias() & "Login.png"
			wait 2
			Browser("Equifax - Portal").InsightObject("InsightObject_3").Click

	 
End Sub


Sub SeleccionarScore()
	While 	Browser("Equifax - Portal").InsightObject("InsightObject_4").Exist = False
		wait 1
	Wend

	Browser("Equifax - Portal").CaptureBitmap RutaEvidencias() & "Aplicaciones.png", True
	imagenToWord "Aplicaciones:", RutaEvidencias() & "Aplicaciones.png"
	Browser("Equifax - Portal").InsightObject("InsightObject_4").Click
	wait 10
	If Browser("Interconnect").InsightObject("InsightObject_15").Exist=False Then
		wait 3
		DataTable("s_Resultado","") ="Fallido"
		DataTable("s_Detalle","loginEquipax") ="Equifax vacio"
		Reporter.ReportEvent micFail, DataTable("s_Resultado","loginEquipax"), DataTable("s_Detalle","loginEquipax")
		Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorServicio.png", True
		imagenToWord "Error al Consultar Servicio Web", RutaEvidencias() & "ErrorServicio.png"
		ExitTestIteration
		Else 
		Set shell = CreateObject("Wscript.Shell")
		shell.SendKeys "{PGDN}"
		wait 1
		shell.SendKeys "{PGDN}"
		wait 1
		shell.SendKeys "{PGDN}"
		wait 1
		shell.SendKeys "{PGDN}"
		wait 1
		While Browser("Interconnect").InsightObject("InsightObject_16").Exist= False
			wait 1
		Wend
		Browser("Interconnect").CaptureBitmap RutaEvidencias() & "RegistroBDExcepciones.png", True
		imagenToWord "Registro BD Excepciones:", RutaEvidencias() & "RegistroBDExcepciones.png"
		wait 1
		Browser("Interconnect").InsightObject("InsightObject_16").Click
		wait 2
		Dim cfecha, c
		cfecha = Date	
		DataTable ("e_fecha",dtLocalSheet) = cfecha

		wait 2
		While Browser("Interconnect").InsightObject("InsightObject_3").Exist=False
			wait 1
		Wend 
		While Browser("Interconnect").InsightObject("InsightObject_4").Exist = false
			wait 1
		Wend
		Browser("Interconnect").InsightObject("InsightObject_4").Click
		wait 2
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys DataTable("e_TipoDocumento",dtLocalSheet)
		wait 2
		Browser("Interconnect").InsightObject("InsightObject_5").Click
		wait 2
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys DataTable("e_NumDoc",dtLocalSheet)

		wait 1
		Browser("Interconnect").InsightObject("InsightObject_2").Click
		wait 2

		While Browser("Interconnect").InsightObject("InsightObject_11").Exist =False
			wait 1
		Wend
		wait 2
		Browser("Interconnect").InsightObject("InsightObject_7").Click
		wait 2
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys DataTable("e_Score",dtLocalSheet)
		wait 1
		Browser("Interconnect").InsightObject("InsightObject_10").Click
		wait 2
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys DataTable("e_CantLineas",dtLocalSheet)
		wait 2
		Browser("Interconnect").InsightObject("InsightObject_14").Click
		wait 2
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys DataTable("e_fecha",dtLocalSheet)
		wait 2
		Browser("Interconnect").InsightObject("InsightObject_8").Click
		wait 2
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys DataTable("e_cantFinan",dtLocalSheet)
		wait 2
		Browser("Interconnect").InsightObject("InsightObject_9").Click
		wait 3
		If Browser("Interconnect").InsightObject("InsightObject_13").Exist =False Then
			wait 4
			DataTable("s_Resultado","") ="Fallido"
			DataTable("s_Detalle","loginEquipax") ="Error al registrar score en Equipax"
			Reporter.ReportEvent micFail, DataTable("s_Resultado","loginEquipax"), DataTable("s_Detalle","loginEquipax")
			Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorServicio.png", True
			imagenToWord "Error al Consultar Servicio Web", RutaEvidencias() & "ErrorServicio.png"
			ExitTestIteration
		Else 
			wait 4
			Browser("Interconnect").CaptureBitmap RutaEvidencias() & "UsuarioRegistrado.png", True
			imagenToWord "Usuario Registrado:", RutaEvidencias() & "UsuarioRegistrado.png"
	
			
		End If	

	End If	
			
	wait 2
	Browser("Interconnect").Close
	wait 2
	Browser("Equifax - Portal").Close
End Sub









