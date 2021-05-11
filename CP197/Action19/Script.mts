
Dim myDeviceReplay, filas
Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
While JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").JavaTable("SearchJTable").Exist = False
	wait 1
Wend

filas = JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").JavaTable("SearchJTable").GetROProperty("rows")
filas = filas-1
For Iterator = 0 To filas Step 1
	JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").JavaTable("SearchJTable").DoubleClickCell "#"&Iterator, "#0", "LEFT"
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{ENTER}"
	wait 10
	If Iterator = 0 Then
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "^f"
		wait 2
		Window("Ejecutivo de interacción").Window("Resumen de la orden (Orden_2").Window("Resumen de la orden (Orden").WinEdit("Edit").Set "plazo de permanencia en el contrato"

'		Set shell = CreateObject("Wscript.Shell") 
'		shell.SendKeys "plazo de permanencia en el contrato"
'		wait 1
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{ENTER}"
		wait 2
	End If
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "DescargaContrato"&Iterator&".png", True
	imagenToWord "Contrato", RutaEvidencias() & "DescargaContrato"&Iterator&".png"
'	While Window("Ejecutivo de interacción").InsightObject("InsightObject").Exist = false
'		wait 0.2
'	Wend
'	myDeviceReplay.MouseMove 600, 90
'	Window("Ejecutivo de interacción").InsightObject("InsightObject").Click
'
'	While Window("Ejecutivo de interacción").Window("Resumen de la orden (Orden").Window("Resumen de la orden (Orden").Dialog("Guardar como").WinEdit("WinEdit").Exist = False
'		wait 1
'	Wend
'	
'	 Window("Ejecutivo de interacción").Window("Resumen de la orden (Orden").Window("Resumen de la orden (Orden").Dialog("Guardar como").WinEdit("WinEdit").Set "C:\CONTRATOS\"&Iterator&".pdf"
'	 Window("Ejecutivo de interacción").Window("Resumen de la orden (Orden").Window("Resumen de la orden (Orden").Dialog("Guardar como").WinButton("Guardar").Click
'	 
'	While Window("Ejecutivo de interacción").InsightObject("InsightObject").Exist = false
'		wait 1
'	Wend
	
	JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden_2").JavaDialog("Resumen de la orden (Orden").Close

Next
wait 1
 @@ hightlight id_;_33_;_script infofile_;_ZIP::ssf7.xml_;_
