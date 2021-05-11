


While Window("Ejecutivo de interacción").InsightObject("InsightObject").Exist = false
	wait 1
Wend
If (Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Exist or Window("Ejecutivo de interacción").InsightObject("InsightObject_27").Exist)= True Then
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Incidencia.png", True
	imagenToWord "Aviso: Actualmente mos encontramos en incidencia", RutaEvidencias() & "Incidencia.png"
	wait 1
	If Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Exist = True Then
		Window("Ejecutivo de interacción").InsightObject("InsightObject_7").click
	End If
	If Window("Ejecutivo de interacción").InsightObject("InsightObject_27").Exist = True Then
		Window("Ejecutivo de interacción").InsightObject("InsightObject_27").Click
	End If
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN 2}"
	wait 1
	Window("Ejecutivo de interacción").InsightObject("InsightObject_21").Click
	wait 1
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN 2}"
	wait 1
	Window("Ejecutivo de interacción").InsightObject("InsightObject_22").Click
	wait 2
	MsgBox "Realizar la parte de validacion de datos de familiares de forma manual y LUEGO ACEPTAR EL MENSAJE"
	wait 1
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Consulta.png", True
	imagenToWord "Cuestionario aprobado", RutaEvidencias() & "Consulta.png"
	wait 1
	Window("Ejecutivo de interacción").InsightObject("InsightObject_23").Click
	wait 2
	
	Window("Ejecutivo de interacción").InsightObject("InsightObject_6").Click
Else
	'Window("Ejecutivo de interacción").InsightObject("InsightObject").Click
	Window("Ejecutivo de interacción").InsightObject("InsightObject_21").Click
	wait 1

	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ContratacionValidacion.png", True
	imagenToWord "Contratación y Validación de Identidad", RutaEvidencias() & "ContratacionValidacion.png"
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN 1}"
	If Window("Ejecutivo de interacción").InsightObject("InsightObject_9").Exist = True Then
			Window("Ejecutivo de interacción").InsightObject("InsightObject_14").Click

			Window("Ejecutivo de interacción").InsightObject("InsightObject_19").Click

			Set shell = CreateObject("Wscript.Shell") 
			shell.SendKeys "permanencia en el"
			shell.SendKeys "{ENTER}"
			wait 1
			Window("Ejecutivo de interacción").InsightObject("InsightObject_20").Click
			wait 1
			Window("Ejecutivo de interacción").InsightObject("InsightObject_20").Click
			Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Permanencia.png", True
			imagenToWord "Plazo de permanencia", RutaEvidencias() & "Permanencia.png"

	End If
	If Window("Ejecutivo de interacción").InsightObject("InsightObject_11").Exist = True Then
		Window("Ejecutivo de interacción").InsightObject("InsightObject_11").click
		wait 2
		Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Contratos.png", True
		imagenToWord "Contrato", RutaEvidencias() & "Contratos.png"

	End If
	wait 1
	Window("Ejecutivo de interacción").InsightObject("InsightObject_22").Click
	wait 2
	MsgBox "Realizar la parte de validacion de datos de familiares de forma manual y LUEGO ACEPTAR EL MENSAJE"
	wait 1
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Consulta.png", True
	imagenToWord "Cuestionario aprobado", RutaEvidencias() & "Consulta.png"
	wait 1
	Window("Ejecutivo de interacción").InsightObject("InsightObject_23").Click
	wait 2
	
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_2").Exist = false
		wait 1
	Wend
	wait 2
	Window("Ejecutivo de interacción").InsightObject("InsightObject_2").Click
	wait 2
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Exist = False
		wait 1
	Wend
	wait 2
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "DescargarContrato.png", True
	imagenToWord "Descargar Contrato", RutaEvidencias() & "DescargarContrato.png"
	wait 2
	Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Click
	
	Set shell = CreateObject("Wscript.Shell") 	
		shell.SendKeys "{PGDN 2}"
	wait 2
	Window("Ejecutivo de interacción").InsightObject("InsightObject_24").Click
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_25").Exist = False
		wait 1
	Wend
	wait 1
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "DescargaTerminado.png", True
	imagenToWord "Descarga de documentos", RutaEvidencias() & "DescargaTerminado.png"
	Window("Ejecutivo de interacción").InsightObject("InsightObject_25").Click

	While Window("Ejecutivo de interacción").InsightObject("InsightObject_26").Exist = False
		wait 1
	Wend
	wait 2
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "DescargaCompleta.png", True
	imagenToWord "Descarga Completa", RutaEvidencias() & "DescargaCompleta.png"
	Window("Ejecutivo de interacción").InsightObject("InsightObject_6").Click
End If

 @@ hightlight id_;_26_;_script infofile_;_ZIP::ssf12.xml_;_
