


While Window("Ejecutivo de interacción").InsightObject("InsightObject").Exist = false
	wait 1
Wend
If Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Exist = True Then
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Incidencia.png", True
	imagenToWord "Aviso: Actualmente mos encontramos en incidencia", RutaEvidencias() & "Incidencia.png"
	wait 1
	Window("Ejecutivo de interacción").InsightObject("InsightObject_7").click
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN 2}"
	wait 1
	Window("Ejecutivo de interacción").InsightObject("InsightObject_6").Click
Else
	Window("Ejecutivo de interacción").InsightObject("InsightObject").Click
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ContratacionValidacion.png", True
	imagenToWord "Contratación y Validación de Identidad", RutaEvidencias() & "ContratacionValidacion.png"
'	Set shell = CreateObject("Wscript.Shell") 
'	shell.SendKeys "{DOWN 5}"
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN}"
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
	wait 2
	If Window("Ejecutivo de interacción").InsightObject("InsightObject_11").Exist = True Then
		Window("Ejecutivo de interacción").InsightObject("InsightObject_11").click
		wait 2
		Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Contratos.png", True
		imagenToWord "Contrato", RutaEvidencias() & "Contratos.png"

	End If
	shell.SendKeys "{PGDN 2}"
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
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_5").Exist = False
		wait 1
	Wend
	wait 2
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "DescargaCompleta.png", True
	imagenToWord "Descarga Completa", RutaEvidencias() & "DescargaCompleta.png"
	Window("Ejecutivo de interacción").InsightObject("InsightObject_6").Click
End If

 @@ hightlight id_;_26_;_script infofile_;_ZIP::ssf12.xml_;_
