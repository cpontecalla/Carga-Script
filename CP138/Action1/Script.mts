Option Explicit

Dim objKey

Set objKey = CreateObject("WScript.Shell")

Call ConsultaSaldo()
Sub ConsultaSaldo()
	
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaStaticText("Número de documento(st)").Exist) = False 
			wait 1	
		Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Panel.png", True
	imagenToWord "Panel de Interacción", RutaEvidencias() & "Panel.png"

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTab("Acciones rápidas").Select "Suscripciones" @@ hightlight id_;_10313193_;_script infofile_;_ZIP::ssf1.xml_;_
	wait 1
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaEdit("TextFieldNative$1").Exist=false
			wait 1
		Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaEdit("TextFieldNative$1").Set DataTable("e_ID_Servicio", dtLocalSheet)
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Buscar ahora").Click
	wait 3
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Susc.png", True
	imagenToWord "Busqueda Suscripción", RutaEvidencias() & "Susc.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Acciones rápidas").DoubleClickCell "#0", "Número de Suscripción" @@ hightlight id_;_22352119_;_script infofile_;_ZIP::ssf4.xml_;_
	wait 2
	objKey.SendKeys ("{ENTER}") @@ hightlight id_;_13546060_;_script infofile_;_ZIP::ssf3.xml_;_
	wait 2
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles de la Suscripción:").JavaCheckBox("Bolsas").Exist) = False
			wait 1
		Wend @@ hightlight id_;_8910786_;_script infofile_;_ZIP::ssf6.xml_;_
	wait 5
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Saldos.png", True
	imagenToWord "Verificacion de Saldos", RutaEvidencias() & "Saldos.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles de la Suscripción:").JavaButton("<html>Herramientas de").Click @@ hightlight id_;_6661801_;_script infofile_;_ZIP::ssf8.xml_;_
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles de la Suscripción:").JavaMenu("Imprimir los resultados").Select @@ hightlight id_;_11660407_;_script infofile_;_ZIP::ssf7.xml_;_
	wait 5
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ValidacionFinal.png", True
	imagenToWord "Verificación de Consulta Saldo", RutaEvidencias() & "ValidacionFinal.png" 

End Sub




