Dim var1, varfila
Dim str_tipoDocumento
Dim str_nroDocumento
Dim CantFilas
Dim Iterator
Dim Rol
Dim Comp
Dim str_DniContacto
Dim NroDig


Dim str_IDServicio
Dim str_Plan
Dim str_Motivo
Dim str_Motivo_Text
Dim str_NroOrden
Dim str_str_TipoData
Dim str_ValOrden

str_tipoDocumento= DataTable("e_TipoDocumento", dtLocalSheet)
str_nroDocumento=DataTable("e_NumDocumento", dtLocalSheet)
str_DniContacto=DataTable("e_DniContacto", dtLocalSheet)

str_IDServicio  = DataTable("e_ID_Servicio", dtLocalSheet)
str_Plan    	= DataTable("e_Plan", dtLocalSheet)
str_Motivo		= DataTable("e_Motivo", dtLocalSheet)
str_Motivo_Text = DataTable("e_Motivo_Text", dtLocalSheet)
str_NroOrden    = DataTable("s_Nro_Orden", dtLocalSheet)
str_ValOrden    = DataTable("s_ValEstadoOrden", dtLocalSheet)

Call Busqueda_Cliente()
Call SeleccionarContacto()
Call PanelInteraccion()
Call CambioTitularidad()
'Call CambioPlanTarifario()
Call ActualizarAtributosOrden()
Call GeneracionOrden()
Call ValidaOrden()
Call DetalleActividadOrden()

Sub Busqueda_Cliente()
		wait 1
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaButton("Cerrar").Click
		End If
		
		JavaWindow("Ejecutivo de interacción").JavaButton("Find-Caller").Click
		
			While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaRadioButton("Cliente").Exist) = False
				wait 1
			Wend
			wait 1
		If str_tipoDocumento="ACUERDO" then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaRadioButton("Acuerdo de Facturación").Set
		elseif not str_tipoDocumento="SUSCRIPCION" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaRadioButton("Cliente").Set
		End If
		wait 1
		
			Select Case str_tipoDocumento
				Case "RUC"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaList("Tipo ID Compańía:").Select "RUC"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaEdit("TextFieldNative$1").SetFocus
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaEdit("TextFieldNative$1").Set str_nroDocumento
				Case "DNI"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaList("Tipo de documento").Select "DNI"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaEdit("Numero de Documento").Set str_nroDocumento
				Case "Pasaporte"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaList("Tipo de documento").Select "Pasaporte"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaEdit("Numero de Documento").Set str_nroDocumento
				Case "CE"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaList("Tipo de documento").Select "CE"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaEdit("Numero de Documento").Set str_nroDocumento
				Case "IDCLIENTE"	
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaEdit("ID del Cliente:").Set str_nroDocumento
				Case "SUSCRIPCION"
					
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaRadioButton("Suscripción").Set
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaEdit("Número de Suscripción:").Set str_nroDocumento
				Case "ACUERDO"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaEdit("ID del Acuerdo de Facturación:").Set str_nroDocumento
			End Select
			
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"BusquedaCliente"&".png", True
		imagenToWord "Busqueda Cliente", RutaEvidencias() &Num_Iter&"_"&"BusquedaCliente"&".png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaButton("Buscar ahora").Click
		wait 2
		
End Sub
Sub seleccionarActivo
	Dim i, row, estado, tipoIDCompania, rowActive, nifRow
	i=0
	While not i=4
		i=i+1
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").Exist(1) Then
			i=10
			row=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").GetROProperty("rows")
			For rowActive = 0 To row-1 Step 1
				estado=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").GetCellData(rowActive,11)
				tipoIDCompania=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").GetCellData(rowActive,12)
				Rol = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").GetCellData(rowActive,8)
				If estado="Activo" Then
				
					Select Case tipoIDCompania
					
						Case "RUC"
							If Rol = "Titular" Then				
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").SelectRow "#"&rowActive
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaButton("Seleccionar").Click
								rowActive=row
							End If
						Case "NIF"
							nifRow=rowActive
						Case "SUSCRIPCION"
							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").SelectRow "#"&rowActive
							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaButton("Seleccionar").Click
							rowActive=row
					End Select
				
'					If tipoIDCompania="RUC" Then
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").SelectRow "#"&rowActive
'						JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "CargaNumeroActivo.png", True	
'						imagenToWord "Carga de números activos", RutaEvidencias() & "CargaNumeroActivo.png"
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaButton("Seleccionar").Click
'						rowActive=row
'					ElseIf tipoIDCompania="NIF" Then
'						nifRow=rowActive
'						
'					End If
					If rowActive = row-1 Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").SelectRow "#"&rowActive
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaButton("Seleccionar").Click
					End If
					
				End If
			Next
			
		End If
	Wend
'	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "CargaNumeroActivo.png", True	
'	imagenToWord "Carga de números activos", RutaEvidencias() & "CargaNumeroActivo.png"
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaButton("Seleccionar").Click
	wait 2
End Sub
Sub SeleccionarContacto()
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("0 Registros").Exist(2) Then
 	var1 =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("0 Registros").GetROProperty("text")
	 	If (var1= "0 Registros") or (var1= "-- Registros") Then
	 		Reporter.ReportEvent micFail,"Fallido", "Nose se encuentra el RUC:"&DataTable("e_NumDocumento", dtLocalSheet)
	 		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("Cancelar").Click
	 		wait 1
	 		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No se encontro al realizar la búsqueda"
			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
			ExitActionIteration
	 	else
	 		varfila=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaTable("SearchJTable").GetROProperty("rows")
			varfila= CInt(varfila)
			wait 2
				For Iterator = 0 To varfila - 1 Step 1
					Select Case str_tipoDocumento
						Case "DNI"
							Comp = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaTable("SearchJTable").GetCellData(Iterator,9)
							If (Comp = "") Then
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaTable("SearchJTable").SelectRow "#"&Iterator
								JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SeleccionaContacto"&".png", True
								imagenToWord "Selecciona Contacto", RutaEvidencias() &Num_Iter&"_"&"SeleccionaContacto"&".png"
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("Seleccionar").Click
								wait 2
								Exit for
							End If
'							If Comp = str_DniContacto Then
'								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaTable("SearchJTable").SelectRow "#"&Iterator
'								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("Seleccionar").Click
'								wait 2
'								Exit for
'							End If
						Case "CE"
							Comp = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaTable("SearchJTable").GetCellData(Iterator,9)
							If (Comp = "") Then
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaTable("SearchJTable").SelectRow "#"&Iterator
								JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SeleccionaContacto"&".png", True
								imagenToWord "Selecciona Contacto", RutaEvidencias() &Num_Iter&"_"&"SeleccionaContacto"&".png"
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("Seleccionar").Click
								wait 2
								Exit for
							End If
						Case "RUC"
							Comp = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaTable("SearchJTable").GetCellData(Iterator,3)
							Comp = Cstr(Comp)
							NroDig = Len(Comp)
							'Se agrega 0 si son 7 digitos
							If (NroDig = "7") Then
								Comp = "0"&Comp
							End If
							If (Comp = str_DniContacto) Then
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaTable("SearchJTable").SelectRow "#"&Iterator
								JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SeleccionaContacto"&".png", True
								imagenToWord "Selecciona Contacto", RutaEvidencias() &Num_Iter&"_"&"SeleccionaContacto"&".png"
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("Seleccionar").Click
								Wait 2
								Exit for
							End If
					End Select
					If (Iterator = varfila - 1) Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaTable("SearchJTable").SelectRow "#0"
						JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SeleccionaContacto"&".png", True
						imagenToWord "Selecciona Contacto", RutaEvidencias() &Num_Iter&"_"&"SeleccionaContacto"&".png"
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("Seleccionar").Click
					End If
				Next
	 	End If
	End If
End Sub
Sub PanelInteraccion()

	t = 0
	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Ver Detalles").Exist) = False
		wait 1	
		t = t + 1
		If (t >= 30) Then
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Panel de Interacción- de manera correcta"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png", True
			imagenToWord "No cargó la pantalla -Panel de Interacción- de manera correcta",RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
		wait 1
	Wend
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"PanelInteracción"&".png", True
	imagenToWord "Panel de Interacción", RutaEvidencias() &Num_Iter&"_"&"PanelInteracción"&".png"
	t = 0
	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Exist) = False
		wait 1	
		t = t + 1
		If (t >= 30) Then
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Panel de Interacción- de manera correcta"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png", True
			imagenToWord "No cargó la pantalla -Panel de Interacción- de manera correcta",RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
		wait 1
	Wend
End Sub
Sub CambioTitularidad()
	
	JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").JavaMenu("Cambiar la titularidad").Select
	If ucase(DataTable("WIC1",dtLocalSheet)) = "SI" Then
		
RunAction "WIC1", oneIteration
	End If
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Productos").JavaButton("Buscar ahora").Exist) = False
			wait 1
		Wend

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Productos").JavaEdit("TextFieldNative$1").Set str_IDServicio
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Productos").JavaButton("Buscar ahora").Click
	
	tiempo = 0
	Do 
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Productos").JavaButton("Buscar ahora").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Productos").JavaButton("Buscar ahora").Click
			nroreg=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Productos").JavaButton("1 Registros").GetROProperty("attached text")
			tiempo = tiempo + 1
			Wait 1
		End If
		If (tiempo >= 15) Then
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No se encuentra el ID_Servicio: "&str_IDServicio&" en la busqueda por Suscripción"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErroBusquedaSuscripcion.png", True
			imagenToWord "Error Busqueda Suscripción a Cambiar de Titularidad",RutaEvidencias() &Num_Iter&"_"&"ErroBusquedaSuscripcion.png"
			Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Productos").JavaButton("Cerrar").Click
				While(JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Exist) = False
					wait 1
				Wend
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Click
				While(JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar Productos").JavaList("Motivo:").Exist) = False
					wait 1
				Wend
			JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar Productos").JavaList("Motivo:").Select "Pedido de Cliente" 
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar Productos").JavaButton("Aceptar").Click
				While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 478446254A").Exist) = False
					wait 1
				Wend
			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 478446254A").GetROProperty("text")
			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),10)
			DataTable("s_Resultado",dtLocalSheet)="Fallido"
			DataTable("s_Detalle",dtLocalSheet)="La orden "&DataTable("s_Nro_Orden",dtLocalSheet)&" se cancelo, dado que no se encontro la búsqueda"
			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet) 
			DataTable("s_ValEstadoOrden",dtLocalSheet)="Cancelado"
			wait 1
		End If
	Loop While Not(nroreg="1 Registros")
	wait 3
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Productos").JavaTable("Ver por:").SelectRow "#0"
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"BusquedaIdServicio"&".png", True
	imagenToWord "Busqueda de Suscripción", RutaEvidencias() &Num_Iter&"_"&"BusquedaIdServicio"&".png"
	wait 2

	tiempo=0
		Do 
		tiempo=tiempo+1
		var1=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Productos").JavaButton("Cambio de titularidad").GetROProperty("enabled")
			If (tiempo >= 250) Then
				DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)="No se habilito el boton 'Cambio de titularidad'"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
				ExitActionIteration
			End If
			wait 1
		Loop While Not (var1="1")
		wait 3
	
	var1=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Productos").JavaButton("Cambio de titularidad").GetROProperty("enabled")	
	If var1="0" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Productos").JavaTable("Ver por:").SelectRow "#0"
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"BusquedaIdServicio"&".png", True
		imagenToWord "Busqueda de Suscripción", RutaEvidencias() &Num_Iter&"_"&"BusquedaIdServicio"&".png"
		wait 1
	End If

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Productos").JavaButton("Cambio de titularidad").Click
	wait 3
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Productos").JavaButton("Siguiente >").Click


	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Texto del motivo:").Exist = False
		wait 1
	Wend
			

End Sub
Sub CambioPlanTarifario()
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Configuraciónydetalles"&".png", True
	imagenToWord "Configuración y detalles", RutaEvidencias() &Num_Iter&"_"&"Configuraciónydetalles"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Configuración y detalles").JavaButton("Seleccionar otras ofertas").Click
	
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Configuración y detalles").JavaButton("Buscar").Exist) = False
			wait 1
		Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaDialog("Configuración y detalles").JavaEdit("TextFieldNative$1").Set str_Plan
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaDialog("Configuración y detalles").JavaButton("Buscar").Click
	
		tiempo=0
		Do 
			If (tiempo >=40) Then
				DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)=""
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
				ExitActionIteration 
			End If
		wait 1
		Loop While JavaWindow("Ejecutivo de interacción").JavaDialog("Configuración y detalles").JavaButton("Prepago con Tarifa Única").Exist
		wait 2
	JavaWindow("Ejecutivo de interacción").JavaDialog("Configuración y detalles").JavaCheckBox("Seleccionar").Set "ON"
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SeleccionPlan"&".png", True
	imagenToWord "Seleccion del Plan Tarifario", RutaEvidencias() &Num_Iter&"_"&"SeleccionPlan"&".png"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaDialog("Configuración y detalles").JavaButton("Siguiente >").Click
	
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Configuración y detalles").JavaButton("Agregar productos").Exist) = False
			wait 1
		Wend
		
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Configuracionydetalles_2"&".png", True
	imagenToWord "Configuracion y detalles", RutaEvidencias() &Num_Iter&"_"&"Configuracionydetalles_2"&".png"
	JavaWindow("Ejecutivo de interacción").JavaDialog("Configuración y detalles").JavaButton("Siguiente >").Click
	
		tiempo=0
		Do
			If (tiempo>=40) Then
				DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)="No carga la pantalla Actualizar Atributos de Acción de Orden"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
				ExitActionIteration
			End If
			wait 1
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Exist
		wait 1
End Sub
Sub ActualizarAtributosOrden()
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Select str_Motivo
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Texto del motivo:").Set str_Motivo_Text
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActualizarAtributos"&".png", True
	imagenToWord "Actualizar Atributos de la Orden", RutaEvidencias() &Num_Iter&"_"&"ActualizarAtributos"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Siguiente >").Click
	
	tiempo=0
		Do
			If (tiempo >= 40) Then
				DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)="No carga la ventana 'Negociar Configuración del Producto Movi'"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
				ExitActionIteration
			End If
			wait 1
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Exist
		wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"NegociarConfiguración"&".png", True
	imagenToWord "Negociar Configuración", RutaEvidencias() &Num_Iter&"_"&"NegociarConfiguración"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Click
	
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").Exist = false
		wait 1
	Wend
	text = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").GetROProperty("text")
	While text = ""
		wait 1
		text = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").GetROProperty("text")
	Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"NegociarConfiguración"&".png", True
	imagenToWord "Negociar Configuración", RutaEvidencias() &Num_Iter&"_"&"NegociarConfiguración"&".png"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaRadioButton("Nuevo").Set "ON"
	wait 5
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Siguiente >").Click


'	
	
	wait 1
	
		tiempo=0
		Do
			If (tiempo >= 40) Then
				DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)="No carga la ventana 'Negociar Configuración del Producto Movi'"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
				ExitActionIteration
			End If
			wait 1
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validar y Ver Contrato").Exist
		wait 1
End Sub

Sub GeneracionOrden()

		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validar y Ver Contrato").Click
		wait 3
		If ucase(DataTable("WIC2",dtLocalSheet)) = "SI" Then
			
RunAction "WIC2", oneIteration
ExitActionIteration
		End If

		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(2) Then
			wait 3
			var1= JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
			wait 1
		End If
	

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"GeneradorContratos"&".png", True
		imagenToWord "Generador de Contratos", RutaEvidencias() &Num_Iter&"_"&"GeneradorContratos"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Close
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").Set "ON"
	End If
	
	
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Exist) = False
		wait 1
	Wend
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Resumendelaorden"&".png", True
	imagenToWord "Resumen de la orden", RutaEvidencias() &Num_Iter&"_"&"Resumendelaorden"&".png"
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
	
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 478446254A").JavaButton("Cerrar").Exist) = False
			wait 1
		Wend
	
	DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 478446254A").GetROProperty("text")
	flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
	DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),8)
	Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OrdenGenerada"&".png", True
	imagenToWord "OrdenGenerada", RutaEvidencias() &Num_Iter&"_"&"OrdenGenerada"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 478446254A").JavaButton("Cerrar").Click
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Cerrar").Click
	wait 1
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Guardar el formulario").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Guardar el formulario").JavaButton("Descartar").Click
		wait 1
	End If
End Sub
Sub ValidaOrden()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
	
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Exist) = False
			wait 1
		Wend
		wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaCheckBox("Solo órdenes pendientes").Set "OFF"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
	wait 2
	
		tiempo=0
		Do 
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
				nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("1 Registros").GetROProperty("attached text")
				tiempo=tiempo+1
				wait 1
			End If
			If (tiempo >= 180) Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar la Gestion Logistica"
					Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
					ExitActionIteration
			End If
		Loop While Not (nroreg="1 Registros")
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:") @@ hightlight id_;_26172125_;_script infofile_;_ZIP::ssf85.xml_;_
	Reporter.ReportEvent micPass,"Se valida el estado de la orden", DataTable("s_ValEstadoOrden", dtLocalSheet)
	
		tiempo = 0
		Do
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
			tiempo = tiempo +1
			wait 5
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:")
				If (tiempo >= 180) Then	
						DataTable("s_Resultado",dtLocalSheet) = "Fallido" 
						DataTable("s_Detalle", dtLocalSheet) = "La Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)&" no culmino en estado Programado"
						Reporter.ReportEvent micFail,DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
						If DataTable("s_ValEstadoOrden", dtLocalSheet) = "Enviado" Then
							Exit Do
							wait 1
						End If	
				else
					Reporter.ReportEvent micPass,"Correcto", "Se valida el estado de la orden: "&DataTable("s_Nro_Orden",dtLocalSheet)
				End If
		Loop While Not (DataTable("s_ValEstadoOrden", dtLocalSheet) = "Cerrado")
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OrdenCerrada"&".png", True
		imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"OrdenCerrada"&".png"
		DataTable("s_Resultado", dtLocalSheet) = "Éxito"
		DataTable("s_Detalle", dtLocalSheet) = "La orden culminó correctamente en el estado "&DataTable("s_ValEstadoOrden", dtLocalSheet)
		Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		wait 2
	
End Sub
Sub DetalleActividadOrden()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaList("Ver por:").Select "Acciones de orden"
	wait 3
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").DoubleClickCell 0, "#8", "LEFT"
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{ENTER}"
	wait 1
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 478471128").JavaEdit("TextAreaNative$1").Exist)=False
			wait 1
		Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 478471128").JavaTab("Ver por:").Select "Actividad"
	
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 478471128").JavaTable("SearchJTable").Exist)=False
			wait 1	
		Wend
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DetalleOrden"&".png", True
	imagenToWord "Detalle de la Orden", RutaEvidencias() &Num_Iter&"_"&"DetalleOrden"&".png"
	shell.SendKeys "{PGDN}"
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DetalleOrden_2"&".png", True
	imagenToWord "Detalle de la Orden_2", RutaEvidencias() &Num_Iter&"_"&"DetalleOrden_2"&".png"
	wait 1

	filas=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 478471128").JavaTable("SearchJTable").GetROProperty("rows")
	For Iterator = 0 To filas-1
		varselec=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 478471128").JavaTable("SearchJTable").GetCellData(Iterator,0)
	Next
	
	If varselec<>"Cerrar Acción de Orden" Then
	 	DataTable("s_Resultado",dtLocalSheet)="Fallido"
		DataTable("s_Detalle",dtLocalSheet)="La orden "&DataTable("s_Nro_Orden",dtLocalSheet)&" culmino en la actividad "&varselec&""
		Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 478471128").JavaButton("Cancelar").Click
		wait 2
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
			wait 2
		End If
		ExitActionIteration
		wait 1
	End If
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 478471128").JavaButton("Cancelar").Click

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
		wait 2
	End If
End Sub


