
Dim TipoCliente, TipDocumento, NumDocumento, Nombre, Apellidos, Nacionalidad,Departamento, Provincia, Distrito, TipoVia, NombreVia, Numero, Genero

TipoCliente 	= DataTable("e_TipoCliente", dtLocalsheet)
TipDocumento 	= DataTable("e_TipoDocumento", dtLocalsheet)
NumDocumento 	= DataTable("e_NumDocumento", dtLocalsheet)
TipoIDCompania 	= DataTable("e_TipoIDCompañia", dtLocalsheet)
IDCompania		= DataTable("e_IDCompañia", dtLocalsheet)
Nombre 			= DataTable("e_Nombres", dtLocalsheet)
Apellidos 		= DataTable("e_Apellidos", dtLocalsheet)
Nacionalidad 	= DataTable("e_Nacionalidad", dtLocalsheet)
Departamento 	= ucase(DataTable("e_Departamento", dtLocalsheet))
Provincia 		= ucase(DataTable("e_Provincia", dtLocalsheet))
Distrito 		= ucase(DataTable("e_Distrito", dtLocalsheet))
TipoVia 		= ucase(DataTable("e_TipoVia", dtLocalsheet))
NombreVia 		= ucase(DataTable("e_NombreVia", dtLocalsheet))
Numero 			= DataTable("e_Numero", dtLocalsheet)
Genero 			= DataTable("e_Genero", dtLocalsheet)
FechaNac 		= DataTable("e_FechaNac", dtLocalsheet)
SubtipoCli		= DataTable("e_Sub_Tip_Cliente", dtLocalsheet)

Call EsperaCrearCliente()
Call TipCliente()
Call DetallesContacto()
Call DetalleCuenta()
Call DetalleDireccion()
Call DetallesCliente()
Call GuardarCliente()

Sub EsperaCrearCliente()
	
	While(varmenu="1") = False
		varmenu=JavaWindow("Ejecutivo de interacción").JavaMenu("Crear").JavaMenu("Cliente").GetROProperty("enabled")
	Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaMenu("Crear").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Crear").JavaMenu("Cliente").Select
    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & AltaClien&".png", True
	imagenToWord "Dar Alta al Cliente", RutaEvidencias() & AltaClien&".png"

End Sub
Sub TipCliente()

'	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaRadioButton("Cliente de la cuenta").Exist) = false
'    	wait 1
'	Wend
	Select Case DataTable("e_TipoCliente", dtLocalsheet)
		Case "Residencial"
			wait 1
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaRadioButton("Cliente de la cuenta").Set "ON"
'			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaRadioButton("Contacto del Cliente").Set "ON"
		Case "Corporativo"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaRadioButton("Cliente de la cuenta").Set "ON"
			wait 1
	End Select

End Sub
Sub DetallesContacto()
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaStaticText("Tipo de documento(st)").Exist = False
		wait 1
	Wend
	Select Case TipDocumento
		Case "DNI"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Tipo de documento").Select "DNI"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Identificación de la persona").Set NumDocumento
		 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Validar").Click
		 	wait 2
		 	Call MensContactoExistente()
		  	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Nombres").GetROProperty("text") = "")
		   		wait 1
		    Wend
		    
		    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & DetContacto&".png", True
			imagenToWord "Detalles de contacto", RutaEvidencias() & DetContacto&".png"

		Case "CE"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Tipo de documento").Select TipDocumento
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Identificación de la persona").Set NumDocumento
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Género:").Select Genero
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Fecha de Nacimiento:").Set FechaNac
			wait 2
		 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Validar").Click
		 	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaStaticText("Tipo ID de la persona").Exist = True Then
		 		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & ValDoc&".png", True
				imagenToWord "Mensaje sobre validación de Documento", RutaEvidencias() & ValDoc&".png"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		 	End If
		 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Nombres").Set Nombre
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Apellidos:").Set Apellidos
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Nacionalidad:").Select Nacionalidad
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & DetContacto&".png", True
			imagenToWord "Detalles de contacto", RutaEvidencias() & DetContacto&".png"
		Case "Pasaporte"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Tipo de documento").Select TipDocumento
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Identificación de la persona").Set NumDocumento
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Género:").Select Genero
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Fecha de Nacimiento:").Set FechaNac
			wait 2
		 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Validar").Click
		 	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaStaticText("Tipo ID de la persona").Exist = True Then
		 		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & ValDoc&".png", True
				imagenToWord "Mensaje sobre validación de Documento", RutaEvidencias() & ValDoc&".png"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		 	End If
		 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Nombres").Set Nombre
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Apellidos:").Set Apellidos
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Nacionalidad:").Select Nacionalidad
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & DetContacto&".png", True
			imagenToWord "Detalles de contacto", RutaEvidencias() & DetContacto&".png"
		Case "PTP"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Tipo de documento").Select TipDocumento
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Identificación de la persona").Set NumDocumento
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Género:").Select Genero
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Fecha de Nacimiento:").Set FechaNac
			wait 2
		 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Validar").Click
		 	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaStaticText("Tipo ID de la persona").Exist = True Then
		 		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & ValDoc&".png", True
				imagenToWord "Mensaje sobre validación de Documento", RutaEvidencias() & ValDoc&".png"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		 	End If
		 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Nombres").Set Nombre
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Apellidos:").Set Apellidos
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Nacionalidad:").Select Nacionalidad
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & DetContacto&".png", True
			imagenToWord "Detalles de contacto", RutaEvidencias() & DetContacto&".png"
		Case "SNM"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Tipo de documento").Select TipDocumento
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Identificación de la persona").Set NumDocumento
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Género:").Select Genero
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Fecha de Nacimiento:").Set FechaNac
			wait 2
		 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Validar").Click
		 	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaStaticText("Tipo ID de la persona").Exist = True Then
		 		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & ValDoc&".png", True
				imagenToWord "Mensaje sobre validación de Documento", RutaEvidencias() & ValDoc&".png"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		 	End If
		 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Nombres").Set Nombre
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Apellidos:").Set Apellidos
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Nacionalidad:").Select Nacionalidad
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & DetContacto&".png", True
			imagenToWord "Detalles de contacto", RutaEvidencias() & DetContacto&".png"
	End Select
End Sub
Sub DetalleCuenta()

	If DataTable("e_TipoCliente",dtLocalSheet)<>"Residencial" Then
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Tipo ID Compańía:").Select TipoIDCompania
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("ID de Compańía").Set IDCompania
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Validar_3").Click
		wait 1
		Call ValidaMensajes()
			While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Validar_3").GetROProperty("image label") = ""
				wait 1
			Wend
	End If
					
End Sub
Sub DetalleDireccion()
	Select Case TipDocumento
		Case "DNI"
		    If not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Provincia:").GetROProperty("text") = ""  Then
		    	Call ValProvincia()
		    End If
			If not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Distrito:").GetROProperty("text") = "" Then
				Call ValDistrito()
			End If
			If not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Nombre de Vía:").GetROProperty("text") = "" Then 
				Call ValDireccion()
			End If
	 	 Case "CE"
	 	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Departamento:").Select Departamento
	 	    wait 1
	 	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Provincia:").Set Provincia
	 	    wait 1
			Call ValProvincia()
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Distrito:").Set Distrito
			wait 1
			Call ValDistrito()
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Tipo de Vía").Select TipoVia
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Nombre de Vía:").Set NombreVia
			wait 1
			Call ValDireccion()
		Case "Pasaporte"
			 JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Departamento:").Select Departamento
	 	    wait 1
	 	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Provincia:").Set Provincia
	 	    wait 1
			Call ValProvincia()
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Distrito:").Set Distrito
			wait 1
			Call ValDistrito()
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Tipo de Vía").Select TipoVia
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Nombre de Vía:").Set NombreVia
			wait 1
			Call ValDireccion()
		Case "PTP"
			 JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Departamento:").Select Departamento
	 	    wait 1
	 	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Provincia:").Set Provincia
	 	    wait 1
			Call ValProvincia()
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Distrito:").Set Distrito
			wait 1
			Call ValDistrito()
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Tipo de Vía").Select TipoVia
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Nombre de Vía:").Set NombreVia
			wait 1
			Call ValDireccion()
		Case "SNM"
			 JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Departamento:").Select Departamento
	 	    wait 1
	 	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Provincia:").Set Provincia
	 	    wait 1
			Call ValProvincia()
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Distrito:").Set Distrito
			wait 1
			Call ValDistrito()
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Tipo de Vía").Select TipoVia
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Nombre de Vía:").Set NombreVia
			wait 1
			Call ValDireccion()
	End Select
End Sub
Sub DetallesCliente()
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Tipo de Cliente:").Select DataTable("e_TipoCliente", dtLocalsheet) 
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Subtipo de Cliente:").Select SubtipoCli
	wait 1
	If DataTable("e_TipoCliente", dtLocalSheet)<>"Residencial" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Lookup-notValidated_4").Click
			While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente").JavaTable("SearchJTable").Exist)=False
				wait 1
			Wend
			wait 1
		filas=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente").JavaTable("SearchJTable").GetROProperty("rows")
		For Iterator = 0 To filas-1 Step 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente").JavaTable("SearchJTable").SelectRow(Iterator)
			dato=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente").JavaTable("SearchJTable").GetCellData(Iterator,"#1") 
			wait 1
			If CLng(DataTable("e_Fec_Ciclo", dtLocalsheet))=dato Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&".png", True
				imagenToWord "Ciclo de Facturación", RutaEvidencias() & Ciclo&".png"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente").JavaButton("Seleccionar").Click
				wait 1
				Exit For
			End If		
			
		Next	
			If  UCASE(DataTable("e_Permitir_Acuerdo", dtLocalSheet))="SI" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaCheckBox("Permitir Acuerdo Comercial").Set "ON"
			 wait 1
		    End If
	    End If
			
			

End Sub
Sub ValProvincia()

				Set shell = CreateObject("Wscript.Shell") 
	        	shell.SendKeys "{right 100}"
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Lookup-notValidated").Click
				wait 1
				While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Lookup-notValidated").Exist = false
					wait 1
				Wend
				wait 1
				Set shell = CreateObject("Wscript.Shell") 
	        	shell.SendKeys "{right 100}"
				wait 2
				Dim c
				c=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Lookup-notValidated").GetROProperty("tagname")
				If c = "Lookup-Validated" Then
					Reporter.ReportEvent micPass, "Pass", "Provincia Validada"
'					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Prov&".png", True
'					imagenToWord "Provincia validada", RutaEvidencias() & Prov&".png"
				End If
End Sub
Sub ValDistrito()
				Set shell = CreateObject("Wscript.Shell") 
	        	shell.SendKeys "{left 100}"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Lookup-notValidated_2").Click
				While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Lookup-notValidated_2").Exist = false
					wait 1
				Wend
				Dim d
				While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Lookup-notValidated_2").GetROProperty("tagname")<>"Lookup-Validated"
					wait 1
				Wend
				d=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Lookup-notValidated_2").GetROProperty("tagname")
				If d = "Lookup-Validated" Then
					Reporter.ReportEvent micPass, "Pass", "Distrito Validado"
'					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Dist&".png", True
'					imagenToWord "Distrito validado", RutaEvidencias() & Dist&".png"
				End If
			
End Sub
Sub ValDireccion()
					If ucase(TipDocumento) = "CE" Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Número:").Set Numero
						Else 
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Número:").Set "1"
					End If
					
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Piso:").Set "1"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Interior / N° de Dpto.").Set "1"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Bloque/Manzana:").Set "1"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Lote:").Set "1"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Código de localidad:").Set "1"
					wait 1
					Set shell = CreateObject("Wscript.Shell") 
	        	    shell.SendKeys "{right 300}"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Validar_2").Click
					wait 5
					If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaStaticText("La dirección no pudo ser").Exist Then
						Call MensDireccInvalida()
					End If
					While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Validar_2").GetROProperty("image label") = ""
						wait 1
					Wend
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Direc&".png", True
					imagenToWord "Dirección validada", RutaEvidencias() & Direc&".png"
End Sub
Sub GuardarCliente()

	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & ClienCreado&".png", True
	imagenToWord "Datos del Cliente a Crear Completados", RutaEvidencias() & ClienCreado&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("ComboBoxNative$1").Select "Guardar y Ver Detalles"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Guardar").Click
	wait 1
	Call ValidaMensajes()
		While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: ' SELVA INDUSTRIAS").JavaEdit("ID de la cuenta:").Exist)Or(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: ' SELVA INDUSTRIAS").JavaEdit("ID de Contacto:").Exist))=False
			wait 1
		Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & DetalleCliente&".png", True
	imagenToWord "Detalle del Cliente Creado", RutaEvidencias() & DetalleCliente&".png"
	DataTable("e_Resultado", dtLocalsheet)="Exito"
	DataTable("e_Detalle", dtLocalsheet)="Ejecución correcta"
	If DataTable("e_TipoCliente",dtLocalsheet)<>"Residencial" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: ' SELVA INDUSTRIAS").JavaEdit("Nombre de la cuenta:").Output CheckPoint("Nombre de la cuenta:")
		wait 1
	End If
	
	lista=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: ' SELVA INDUSTRIAS").JavaList("Cliente actual").GetROProperty("items count") 
	For iterator = 0 To lista-1
		item=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: ' SELVA INDUSTRIAS").JavaList("Cliente actual").GetItemIndex(iterator) 
		If item=1 Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: ' SELVA INDUSTRIAS").JavaList("Cliente actual").Select(iterator)
			Exit For
		End If		
	Next
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: ' SELVA INDUSTRIAS").JavaEdit("ID del Cliente:").Output CheckPoint("ID del Cliente:") @@ hightlight id_;_5517145_;_script infofile_;_ZIP::ssf2.xml_;_
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cuenta: ' SELVA INDUSTRIAS").JavaButton("Cancelar").Click
	wait 1
End Sub

'---------------------------Mensajes---------------------
Sub MensDireccInvalida()
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & MensjVal&".png", True
	imagenToWord "Mensaje", RutaEvidencias() & MensjVal&".png"
	JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Departamento:").Select "LIMA"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Provincia:").Set "LIMA"
	Call ValProvincia()
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Distrito:").Set "SAN ISIDRO"
	Call ValDistrito()
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Tipo de Vía").Select "AVENIDA"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Nombre de Vía:").Set "CAMINO REAL"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Número:").Set "155"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Piso:").Set "1"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Interior / N° de Dpto.").Set "1"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Complejo de Vivienda:").Set "1"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Bloque/Manzana:").Set "1"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Lote:").Set "1"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Código de localidad:").Set "1"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Validar_2").Click
End Sub
Sub MensContactoExistente()
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & MensjValEX&".png", True
		imagenToWord "Mensaje Contacto Existente", RutaEvidencias() & MensjValEX&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		call Cancelar()
		ExitActionIteration
	End If
End Sub
Sub MensMenorEdad()
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaStaticText("El contacto debe se mayor").Exist Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & MensjValEdad&".png", True
		imagenToWord "Mensaje Contacto Menor de Edad", RutaEvidencias() & MensjValEdad&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		Call Cancelar()
		ExitActionIteration
	End If 
End Sub
Sub ValidaMensajes()
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist Then
		varmsg=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
		Select Case varmsg
			Case "La cuenta con el mismo Tipo ID y ID, ya existe"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & MensjValEX&".png", True
				imagenToWord "Mensaje Contacto Existente", RutaEvidencias() & MensjValEX&".png"
				DataTable("e_Resultado",dtLocalsheet)="Fallido"
				DataTable("e_Detalle",dtLocalsheet)=varmsg
				Reporter.ReportEvent micFail, DataTable("e_Resultado",dtLocalsheet), DataTable("e_Detalle",dtLocalsheet)
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
				call Cancelar()
				ExitActionIteration
			Case "El contacto debe se mayor a 18 ańos"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & MensjValEX&".png", True
				imagenToWord "Mensaje Contacto Existente", RutaEvidencias() & MensjValEX&".png"
				Reporter.ReportEvent micFail, DataTable("e_Resultado",dtLocalsheet), DataTable("e_Detalle",dtLocalsheet)
				DataTable("e_Resultado",dtLocalsheet)="Fallido"
				DataTable("e_Detalle",dtLocalsheet)=varmsg
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
				call Cancelar()
				ExitActionIteration
			Case "Falló la validacion de la cuenta, basada en el Tipo ID Compańía, número ID de Compańía brindado"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & MensjValEX&".png", True
				imagenToWord "Mensaje Contacto Existente", RutaEvidencias() & MensjValEX&".png"
				Reporter.ReportEvent micFail, DataTable("e_Resultado",dtLocalsheet), DataTable("e_Detalle",dtLocalsheet)
				DataTable("e_Resultado",dtLocalsheet)="Fallido"
				DataTable("e_Detalle",dtLocalsheet)=varmsg
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
				call Cancelar()
				ExitActionIteration
		End Select
	End If
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Exist Then
		varmsg=JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaObject("JPanel").GetROProperty("text")
		Select Case varmsg
			Case "<html>Contacto duplicado encontrado. Por favor valide el contacto con el botón de<br> consulta antes de guardar.<br><br></html>"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & MensjValEX&".png", True
			imagenToWord "Mensaje Contacto Existente", RutaEvidencias() & MensjValEX&".png"
			DataTable("e_Resultado",dtLocalsheet)="Fallido"
			DataTable("e_Detalle",dtLocalsheet)=varmsg
			Reporter.ReportEvent micFail, DataTable("e_Resultado",dtLocalsheet), DataTable("e_Detalle",dtLocalsheet)
			JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
			call Cancelar()
			ExitActionIteration
		End Select
	End If
	
End Sub

'-----Para CE, Pasaporte
Sub ContactoDuplicado()
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Exist  Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & MensjContDupl&".png", True
		imagenToWord "Contacto Duplicado", RutaEvidencias() & MensjContDupl&".png"
		DataTable("e_Resultado", dtLocalsheet)="Fallido"
		DataTable("e_Detalle", dtLocalsheet)="Contacto duplicado"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
		Call Cancelar()
		ExitActionIteration
	End If
End Sub	
Sub Cancelar()
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Cancelar").Click
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Guardar el formulario").JavaButton("Descartar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Guardar el formulario").JavaButton("Descartar").Click
	End If
	
End Sub




