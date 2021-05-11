Dim Shell
 @@ hightlight id_;_41_;_script infofile_;_ZIP::ssf1.xml_;_

Dim TipoAlta
Dim NumeroPorta, n1, n2, n3, n4, n5, n6, n7, n8, n9, Producto, Operador

TipoAlta = DataTable("e_Tipo",dtLocalSheet)
NumeroPorta = DataTable("e_NumPorta", dtLocalSheet)
Producto = DataTable("e_Producto", dtLocalSheet)
Operador = DataTable("e_Operador", dtLocalSheet)

Call SeleccionProceso()

Sub SeleccionProceso()

	Select Case ucase(TipoAlta)
		Case "POSTPAGO"
			While Window("Ejecutivo de interacción").InsightObject("InsightObject").Exist = False
				WAIT 2
			Wend
						
			Call Captura("Click Postpago","Postpago")
			Window("Ejecutivo de interacción").InsightObject("InsightObject").Click
			
			While Window("Ejecutivo de interacción").InsightObject("InsightObject_2").Exist = False
				wait 1
			Wend
			
			Call Captura("Click Calcular","Calcular")
			
			Window("Ejecutivo de interacción").InsightObject("InsightObject_2").Click
			wait 5
			While Window("Ejecutivo de interacción").InsightObject("InsightObject_3").Exist = False
				wait 1
			Wend
		
			Call Captura("Validar Score","Score")
			wait 5
			Set shell = CreateObject("Wscript.Shell") 
			shell.SendKeys "{PGDN}"
			If Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Exist =True Then
				Call ValidacionDatos()
			End If
			Call ValidaFinal()
			
		Case "PREPAGO"
			
			Call Captura("Click Prepago","Prepago")
			While Window("Ejecutivo de interacción").InsightObject("InsightObject_6").Exist = False
				wait 1
			Wend
			Window("Ejecutivo de interacción").InsightObject("InsightObject_6").Click
			wait 5
			If Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Exist = True Then
				Call ValidacionDatos()
			End If
			Call ValidaFinal()	
	
		Case "PORTABILIDAD"
			While Window("Ejecutivo de interacción").InsightObject("InsightObject_8").Exist = False
				wait 1
			Wend
			Call Captura("Click Portabilidad","Portabilidad")
			Window("Ejecutivo de interacción").InsightObject("InsightObject_8").Click
			
			While Window("Ejecutivo de interacción").InsightObject("InsightObject_10").Exist = False
				wait 1
			Wend
			Window("Ejecutivo de interacción").InsightObject("InsightObject_10").Click
			wait 1
			
			Call InsertaNumeroPorta(NumeroPorta)
			If  ucase(Producto)= "POSTPAGO" Then
				Call Tab()
				Call Down(2)
			ElseIf ucase(Producto)= "PREPAGO" Then
				Call Tab()
				Call Down(1)
			End If	
			Select Case UCASE(Operador)
				Case  "ENTEL"
					Call Tab()
					Call Down(1)
				Case "CLARO"
					Call Tab()
					Call Down(2)
				Case "WINNER SYSTEMS"
					Call Tab()
					Call Down(3)
				Case "VIETTEL"
					Call Tab()
					Call Down(4)
				Case "VIRGIN MOBILE PERU"
					Call Tab()
					Call Down(5)
				Case "OPERADOR VIRTUAL"
					Call Tab()
					Call Down(6)
			End Select
			wait 2
			Call Captura("Consulta Portabilidad","ConsPortab")
			Window("Ejecutivo de interacción").InsightObject("InsightObject_9").Click
			
			While Window("Ejecutivo de interacción").InsightObject("InsightObject_2").Exist = False
				wait 1
			Wend
			Set shell = CreateObject("Wscript.Shell") 
			shell.SendKeys "{PGDN}"
			Call Captura("Click Calcular","Calcular")
			
			Window("Ejecutivo de interacción").InsightObject("InsightObject_2").Click
			wait 5
			While Window("Ejecutivo de interacción").InsightObject("InsightObject_3").Exist = False
				wait 1
			Wend
			
			Set shell = CreateObject("Wscript.Shell") 
			shell.SendKeys "{PGDN}"
			wait 2
			Call Captura("Validar Score","Score")
			
			Set shell = CreateObject("Wscript.Shell") 
			shell.SendKeys "{PGDN}"
			If Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Exist =True Then
				Call ValidacionDatos()
			End If
			Call ValidaFinal()
	End Select
	
End Sub


Sub ValidacionDatos()
	Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Click
	WAIT 2
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN}"
	wait 1
	
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_5").Exist = False
		wait 1
	Wend	
	wait 2
'	If  True Then
'		
'	End If

	Call Captura("Validar Datos personales","ValDatosPers")
	
	Window("Ejecutivo de interacción").InsightObject("InsightObject_5").Click
	wait 3
	
End Sub
Sub ValidaFinal()
		While Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Exist = False
			wait 1
		Wend
		
		Call Captura("Validación Final","ValFinal")
		
		Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Click
End Sub
Sub Captura(Texto,Imagen)
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Imagen&".png", True
	imagenToWord Texto,RutaEvidencias() &Imagen&".png"
End Sub
Sub InsertaNumeroPorta(Numero)
	n1 = Left(Numero,1)
	n2 = left(Numero,2)
	n2 = right(n2,1)
	n3 = left(Numero,3)
	n3 = right(n3,1)
	n4 = left(Numero,4)
	n4 = right(n4,1)
	n5 = left(Numero,5)
	n5 = right(n5,1)
	n6 = left(Numero,6)
	n6 = right(n6,1)
	n7 = left(Numero,7)
	n7 = right(n7,1)
	n8 = left(Numero,8)
	n8 = right(n8,1)
	n9 = right(Numero,1)
	wait 2
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{"&n1&"}"
	shell.SendKeys "{"&n2&"}"
	shell.SendKeys "{"&n3&"}"
	shell.SendKeys "{"&n4&"}"
	shell.SendKeys "{"&n5&"}"
	shell.SendKeys "{"&n6&"}"
	shell.SendKeys "{"&n7&"}"
	shell.SendKeys "{"&n8&"}"
	shell.SendKeys "{"&n9&"}"
	wait 2
End Sub

Sub Down(cantidad)
	For Iterator = 1 To cantidad Step 1
		WAIT 1
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{DOWN}"
	Next
End Sub
Sub Tab()
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{TAB}"
End Sub
