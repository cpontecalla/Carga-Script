
Dim val1, val2, val3, val4, val5, val6, str_dia, Num_Iter, str_TipoCliente, str_TipoDoc, str_NumDoc, intStartTime, intStopTime, var1, var2, str_Genero,  str_FechaNac, str_Nombres, str_Apellidos, str_Dpto, str_Prov, str_Distr, str_TipoVia, str_NombreVia, str_Manzana, str_Lote, var4, str_Nac, var5
Dim shell, str_Numero, b
Call SelecciondeContacto()
Call CambiodeCiclo()

Sub SelecciondeContacto()

		while(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaEdit("Número de documento").Exist)= false
			wait 1
		wend
		
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"Panel.png", True
	imagenToWord "Panel de Interacción",RutaEvidencias() &"Panel.png"
	JavaWindow("Ejecutivo de interacción").JavaButton("'CENTRO DE DIAGNOSTICO").Click

	
End Sub
Sub CambiodeCiclo()

		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO").JavaList("Cliente actual").Exist = False
			wait 1
		Wend
		Dim n
		n = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO").JavaEdit("ID del Cliente:").GetROProperty("text")
		While n = ""
			wait 1
			n = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO").JavaEdit("ID del Cliente:").GetROProperty("text")
		Wend
		wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO").JavaList("Cliente actual").Select "No se seleccionó cliente"
	wait 4
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{DOWN 1}"
	
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO").JavaEdit("Id del Cliente en Legados:").Exist = False
			wait 1
		Wend
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO").JavaEdit("ID del Cliente:").Exist = false
			wait 1
		Wend
		Dim k
		k = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO").JavaEdit("ID del Cliente:").GetROProperty("text")
		While k = ""
			wait 1
			k = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO").JavaEdit("ID del Cliente:").GetROProperty("text")
		Wend
		
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"Contacto.png", True
	imagenToWord "Contacto seleccionado",RutaEvidencias() &"Contacto.png"
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO").JavaButton("15-Residencial/Corporativo").Exist = False
		wait 1
	Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO").JavaButton("15-Residencial/Corporativo").Click
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{ENTER}"
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO_2").JavaEdit("Código de ciclo").Exist = False
			wait 1
		Wend
		Dim g
		g = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO_2").JavaEdit("Código de ciclo").GetROProperty("text")
		While g =  ""
			wait 1
			g = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO_2").JavaEdit("Código de ciclo").GetROProperty("text")
		Wend
		
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"Cambio.png", True
	imagenToWord "Cambio de ciclo",RutaEvidencias() &"Cambio.png"
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO_2").JavaButton("Lookup-Validated").Click
	
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO_3").JavaTable("SearchJTable").Exist = False
			wait 1
		Wend
	wait 5
	
	Dim F,b,c
	F = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO_3").JavaTable("SearchJTable").GetROProperty("rows")
	F = F-1
	
	For Iterator = 0 To F step 1	    
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO_3").JavaTable("SearchJTable").SelectRow ("#"&Iterator)		
		b=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO_3").JavaTable("SearchJTable").GetCellData("#"&Iterator, "#2")	    
		v = Instr(1,b,DataTable("e_CodigoCiclo", dtLocalsheet))	
		If v <> 0 Then	       
		      JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Ciclo.png", True			
		      imagenToWord "Seleccionamos ciclo",RutaEvidencias() & "Ciclo.png"	   			    	
		     Exit for 	    
		End If	
		If Iterator=F Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO_3").JavaTable("SearchJTable").SelectRow ("#0")
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Ciclo.png", True			
		    imagenToWord "Seleccionamos primera opción",RutaEvidencias() & "Ciclo.png"	   			    	
		    Exit for 	    
		End If
	Next
	
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO_3").JavaButton("Seleccionar").Click
	
	
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO_2").JavaEdit("Código de ciclo").Exist = False
			wait 1
		Wend
		Dim h
		h = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO_2").JavaEdit("Código de ciclo").GetROProperty("text")
		While h =  ""
			wait 1
			h = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO_2").JavaEdit("Código de ciclo").GetROProperty("text")
		Wend
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Guardar.png", True			
	imagenToWord "Guardamos cambio de ciclo",RutaEvidencias() & "Guardar.png"	 
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO_2").JavaButton("Guardar").Click

		While JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar una opción").Exist = False
			wait 1
		Wend
JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Mensjae.png", True			
	imagenToWord "Valdiacion de cambio",RutaEvidencias() & "Mensjae.png"	
	
	JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar una opción").JavaButton("Sí").Click

	
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO").JavaButton("Guardar").Exist = False
			wait 1
		Wend
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO").JavaButton("Guardar").Click
		

	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: 'CENTRO DE DIAGNOSTICO").JavaButton("Cerrar").Click


End Sub	
