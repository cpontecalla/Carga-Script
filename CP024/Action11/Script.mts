Dim Shell
 @@ hightlight id_;_41_;_script infofile_;_ZIP::ssf1.xml_;_

Dim TipoAlta
Dim NumeroPorta, n1, n2, n3, n4, n5, n6, n7, n8, n9

TipoAlta = DataTable("e_Tipo",dtLocalSheet)
NumeroPorta = DataTable("e_NumPorta", dtLocalSheet)
producto = DataTable("e_Producto", dtLocalSheet)
Operador = DataTable("e_Operador", dtLocalSheet)

Call SeleccionProceso()

Sub SeleccionProceso()

	Select Case ucase(TipoAlta)
		Case "POSTPAGO"
			While Window("Ejecutivo de interacción").InsightObject("InsightObject").Exist = False
				wait 1
			Wend		
			Call Captura("Click Postpago","Postpago")
			Window("Ejecutivo de interacción").InsightObject("InsightObject").Click
			
			While Window("Ejecutivo de interacción").InsightObject("InsightObject_2").Exist = False
				wait 1
			Wend
			
			Call Captura("Click Calcular","Calcular")
			
			Window("Ejecutivo de interacción").InsightObject("InsightObject_2").Click
			While Window("Ejecutivo de interacción").InsightObject("InsightObject_3").Exist = False
				wait 1
			Wend
			Call Captura("Validar Score","Score")
			wait 5
			Set shell = CreateObject("Wscript.Shell") 
			shell.SendKeys "{PGDN}"
			'If Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Exist =True Then
				Call ValidacionDatos()
			'End If
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

			wait 2
			'Window("Ejecutivo de interacción").InsightObject("InsightObject_11").Click
			Call InsertaNumeroPorta(NumeroPorta)
			
            wait 2
            Window("Ejecutivo de interacción").InsightObject("InsightObject_19").Click
            wait 2

			If  ucase(producto)= "POSTPAGO" Then
				Window("Ejecutivo de interacción").InsightObject("InsightObject_21").Click

			ElseIf ucase(producto)= "PREPAGO" Then
			Window("Ejecutivo de interacción").InsightObject("InsightObject_22").Click

				
			End If	

			wait 1

			
			Window("Ejecutivo de interacción").InsightObject("InsightObject_20").Click
			wait 2
			Select Case UCASE(Operador)
				Case  "ENTEL"
					Call Down(1)
					Set shell = CreateObject("Wscript.Shell") 
					shell.SendKeys "{ENTER}"
				Case "CLARO"
					Call Down(2)
					shell.SendKeys "{ENTER}"
				Case "WINNER SYSTEMS"
					Call Down(3)
					shell.SendKeys "{ENTER}"
				Case "VIETTEL"
					Call Down(4)
					shell.SendKeys "{ENTER}"
				Case "VIRGIN MOBILE PERU"
					Call Down(5)
					shell.SendKeys "{ENTER}"
				Case "POLPHIN TELECOM PERU"
				 	Call Down(6)
					shell.SendKeys "{ENTER}"
				Case "GUINEA MOBILE"
					Call Down(7)
					shell.SendKeys "{ENTER}"
				Case "OPERADOR VIRTUAL"
					Call Down(8)
					shell.SendKeys "{ENTER}"
			End Select
			wait 2
			Call Captura("Consulta Portabilidad","ConsPortab")
		
			Window("Ejecutivo de interacción").InsightObject("InsightObject_16").Click

			While Window("Ejecutivo de interacción").InsightObject("InsightObject_17").Exist = False
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
			wait 4
			If Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Exist =True Then
				Call ValidacionDatos()
			End If
			Call ValidaFinal()
		
	End Select
	
End Sub


			




Sub ValidacionDatos()
	If Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Exist = tRUE Then
		Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Click
	End If
	wait 2
	Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{PGDN}"
	wait 1
	If Window("Ejecutivo de interacción").InsightObject("InsightObject_12").Exist = TRUE Then
		WAIT 2
		If Window("Ejecutivo de interacción").InsightObject("InsightObject_5").Exist Then
		    Window("Ejecutivo de interacción").InsightObject("InsightObject_5").Click
		    wait 1
		End If
		If Window("Ejecutivo de interacción").InsightObject("InsightObject_13").Exist Then
			Window("Ejecutivo de interacción").InsightObject("InsightObject_13").Click
		    wait 1
	        Set shell = CreateObject("Wscript.Shell") 
	        shell.SendKeys "{LEFT 40}"
	        shell.SendKeys "{DELETE 40}"
	        wait 2
	        shell.SendKeys "prueba@gmail.com"
	    	WAIT 1
		End If
		
		
		While Window("Ejecutivo de interacción").InsightObject("InsightObject_12").Exist = False
			wait 1
		Wend	
		wait 2

	
		Call Captura("Validar Datos personales","ValDatosPers")
		Window("Ejecutivo de interacción").InsightObject("InsightObject_12").Click
	End If
	
	wait 3
	
End Sub
Sub ValidaFinal()
		While Window("Ejecutivo de interacción").InsightObject("InsightObject_18").Exist = False
			wait 1
		Wend
		

		Call Captura("Validación Final","ValFinal")
		Window("Ejecutivo de interacción").InsightObject("InsightObject_18").Click
		
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
