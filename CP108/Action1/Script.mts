'**************************************************************************************************
										'REGISTRO
'**************************************************************************************************
'Cliente 			: Telefónica del Perú
'Nombre de Script	: Crear Disputa 
'Nombre de Action	: isputa
'Creado por			: TSOFT
'Creado por			: Naimar Carolina García Perez
'Fecha de Creación	: 31/05/2019
'Descripción		: Script que sirve para generar disputa a nivel de cuenta financiera y factura
'Alcance			: 
'Pre condiciones	:
'Flags				: e_TipoCargo (Disputa Cuenta Financiera y Disputa Factura)
'**************************************************************************************************
										'HISTORIAL
'**************************************************************************************************
'Fecha de Actualización	  				Observaciones 						Actualizado por
'**************************************************************************************************
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'**************************************************************************************************


Dim CantFilas
Dim Tipo
Dim Estado
Dim IdAcuerdo
Dim e_IdAcuerdoFact
Dim e_Factura
Dim e_Monto
Dim e_Tipo
Dim NombreOf
Dim e_DescCargo
Dim e_Suscripcion
Dim Periodo
Dim VencimientoFact
Dim ini_year, ini_month, ini_day, ini_hour, ini_min, ini_sec, ini_extnow

e_Tipo 				= DataTable("e_TipoCargo", dtLocalSheet)
e_IdAcuerdoFact 	= DataTable("e_IdAcuerdo", dtLocalSheet)
e_Factura			= DataTable("e_FacturaLegal", dtLocalSheet)
e_Monto				= DataTable("e_MontoDisputa", dtLocalSheet)
e_DescCargo			= DataTable("e_DescCargo", dtLocalSheet)
e_Suscripcion		= DataTable("e_Suscripcion", dtLocalSheet)

Call SeleccionarDisputa

Sub SeleccionarDisputa
	
	Contador = Datatable.LocalSheet.GetRowCount
	wait 2
	For Iterator = 1 To Contador Step 1
		e_Tipo 				= DataTable("e_TipoCargo", dtLocalSheet)
		e_IdAcuerdoFact 	= DataTable("e_IdAcuerdo", dtLocalSheet)
		e_Factura			= DataTable("e_FacturaLegal", dtLocalSheet)
		e_Monto				= DataTable("e_MontoDisputa", dtLocalSheet)
		e_DescCargo			= DataTable("e_DescCargo", dtLocalSheet)
		e_Suscripcion		= DataTable("e_Suscripcion", dtLocalSheet)
		If e_Tipo = "Disputa Factura" Then
			Call SeleccionarAcuerdo()
			Call SeleccionDocFacturacion()
			Call GeneracionDisputaFact()
			Call ValidacionDisputa()
			Call CierraVentanas()
		ElseIf e_Tipo = "Disputa Cuenta Financiera" Then
			Call BuscarCuentaFinanciera()
			Call DisputaCuentaFinanciera()
			Call GeneracionDisputaCF()
			Call ValidacionDisputaCF()
			Call CierraVentanas()
		End If
		DataTable.SetNextRow
	Next
	
End Sub
Sub SeleccionarAcuerdo
	'Bucles que esperan la carga de la pantalla Panel de Interacción.
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Ver Detalles").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga del Panel de Interaccion"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaPanelInteraccion.png", True
			imagenToWord "Error en la Carga del Panel de Interaccion",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaPanelInteraccion.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
		End If	
	Wend
	

	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Buscar ahora").Exist = False
		Wait 1
		
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga del Panel de Interaccion"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaPanelInteraccion.png", True
			imagenToWord "Error en la Carga del Panel de Interaccion",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaPanelInteraccion.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaStaticText("Número de documento(st)").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga del Panel de Interaccion"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaPanelInteraccion.png", True
			imagenToWord "Error en la Carga del Panel de Interaccion",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaPanelInteraccion.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
		End If	
	Wend
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTab("Acciones rápidas").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTab("Acciones rápidas").Select "Resumen de Facturación"
	End If
	
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Entidades de facturación").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga Resumen de Interacción"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaResumenFacturacion.png", True
			imagenToWord "Error en la Carga Resumen de Interacción",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaResumenFacturacion.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
		End If	
	Wend
	
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Mostrar la lista").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga Resumen de Interacción"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaResumenFacturacion.png", True
			imagenToWord "Error en la Carga Resumen de Interacción",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaResumenFacturacion.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	wait 2
	
	'Se selecciona acuerdo de facturacion abierto.
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Entidades de facturación").Exist = True Then
		CantFilas =  JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Entidades de facturación").GetROProperty("rows")
	End If
	wait 2
	'Si no consigue filas no hay acuerdo que seleccionar
	If CantFilas = 0 Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga Resumen de Interacción"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaResumenFacturacion.png", True
			imagenToWord "Error en la Carga Resumen de Interacción",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaResumenFacturacion.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			ExitActionIteration
	End If
	'Se selecciona el acuerdo de facturación
	For Iterator = 0 To CantFilas Step 1
		Tipo = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Entidades de facturación").GetCellData(Iterator, "#1")
		Estado = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Entidades de facturación").GetCellData(Iterator, "#2")
		If Tipo = "Acuerdo de Facturación" and Estado = "Abierto" Then
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Entidades de facturación").SelectRow "#"&Iterator
			Exit For
			
		End If
		
	Next
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Mostrar la lista").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Mostrar la lista").Click
		wait 2
	End If
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Lista de Acuerdos de Facturaci").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga Resumen de Interacción"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaListasAcuerdos.png", True
			imagenToWord "Error en la Carga de Listas de Acuerdo de Facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaListasAcuerdos.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	'Se selecciona el acuerdo de facturación que viene cargado en el datatable
	 If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Lista de Acuerdos de Facturaci").Exist = True Then
	 	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MostrarLista.png", True
		imagenToWord "Visualizacion de Entidades de facturación LISTA",RutaEvidencias() &Num_Iter&"_"&"MostrarLista.png"
		Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
		
		CantFilas =  JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Lista de Acuerdos de Facturaci").GetROProperty("rows")
	 End If
	 wait 2
	 
	If CantFilas = 0 Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga Lista de Acuerdo de Facturación"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaListasAcuerdos.png", True
			imagenToWord "Error en la Carga de Listas de Acuerdo de Facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaListasAcuerdos.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			
	End If
	
	For Iterator = 0 To CantFilas Step 1
		IdAcuerdo = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Lista de Acuerdos de Facturaci").GetCellData(Iterator, "#5")
		IdAcuerdo = Cstr(IdAcuerdo)
		If IdAcuerdo = e_IdAcuerdoFact Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Lista de Acuerdos de Facturaci").SelectRow "#"&Iterator
			wait 2
			Set shell = CreateObject("Wscript.Shell") 
			shell.SendKeys "{ENTER}"
			wait 2
			Exit For
			
		End If
	Next
	wait 2
End Sub
Sub SeleccionDocFacturacion
'Bucles que esperan a carga de la pantalla Ver Acuerdo de Facturación
	
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaEdit("BAR ID").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver Acuerdo de Facturación"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerAcuerdoFact.png", True
			imagenToWord "Error al Cargar Ver Acuerdo de Facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorVerAcuerdoFact.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaStaticText("Nombre y Dirección de").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver Acuerdo de Facturación"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerAcuerdoFact.png", True
			imagenToWord "Error al Cargar Ver Acuerdo de Facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorVerAcuerdoFact.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaCheckBox("Producción de la Factura").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver Acuerdo de Facturación"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerAcuerdoFact.png", True
			imagenToWord "Error al Cargar Ver Acuerdo de Facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorVerAcuerdoFact.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaTab("Ningún Medio de Pago se").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver Acuerdo de Facturación"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerAcuerdoFact.png", True
			imagenToWord "Error al Cargar Ver Acuerdo de Facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorVerAcuerdoFact.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaTab("Ningún Medio de Pago se").Exist = True Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"AcuerdoFact.png", True
		imagenToWord "Visualizacion Acuerdo de Facturación",RutaEvidencias() &Num_Iter&"_"&"AcuerdoFact.png"
		Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaTab("Ningún Medio de Pago se").Select "Documentos de Facturación"	
	End If
	factura = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaEdit("Número de Factura legal:").GetROProperty("text")
	wait 2
	While e_Factura <> factura
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaButton("<").Click
		wait 1
		factura = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaEdit("Número de Factura legal:").GetROProperty("text")
	Wend
	
	
	Periodo			= JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaEdit("Periodo:").GetROProperty("text")
	VencimientoFact = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaEdit("Fecha de vencimiento:").GetROProperty("text")

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaTable("Impuesto").Exist = True Then
	 	CantFilas =  JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaTable("Impuesto").GetROProperty("rows")
	End If
	wait 2
	 
	If CantFilas = 0 Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga Lista de Documentos de Facturación"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaDocumentosFact.png", True
			imagenToWord "Error en la Carga de Documentos de Facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaDocumentosFact.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			
	End If
	For Iterator = 0 To CantFilas Step 1
		NombreOf = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaTable("Impuesto").GetCellData (Iterator, "#2")
		DescCargos = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaTable("Impuesto").GetCellData (Iterator, "#4")
		If ((NombreOf = "Prueba Movistar Tv Estándar Digital")) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaTable("Impuesto").SelectRow "#"&Iterator
			wait 2
			Select Case e_Tipo
				
				Case "Disputa Factura"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaButton("Disputar Cargo").Click
					Exit For
			End Select
			
		End If
	Next

End Sub
Sub GeneracionDisputaFact
	wait 2
			
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_2").JavaButton("Seleccione Correo electrónico").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Crear caso de facturación"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCrearCaso.png", True
			imagenToWord "Error en la Carga de pantalla Crear caso de facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_2").JavaEdit("Monto reclamado").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Crear caso de facturación"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCrearCaso.png", True
			imagenToWord "Error en la Carga de pantalla Crear caso de facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	

	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_2").JavaList("Tipo 3:").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Crear caso de facturación"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCrearCaso.png", True
			imagenToWord "Error en la Carga de pantalla Crear caso de facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
		wait 3	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_2").JavaButton("Seleccione Correo electrónico").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_2").JavaButton("Seleccione Correo electrónico").Click	
		wait 2
	End If
		
	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_3").JavaTable("SearchJTable").Exist = True Then
		CantFilas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_3").JavaTable("SearchJTable").GetROProperty("rows")
		If CantFilas >= "1"  Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_3").JavaTable("SearchJTable").SelectRow "#0"
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_3").JavaButton("Seleccionar").Click

		ElseIf CantFilas = 0  Then
				
			Valor = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_3").JavaEdit("Valor:").GetROProperty("text")
			If Valor = "" Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_3").JavaEdit("Valor:").Set "prueba@gmail.com"
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_3").JavaButton("Agregar").Click
				Wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_3").JavaTable("SearchJTable").SelectRow "#0"
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_3").JavaButton("Seleccionar").Click						
			End If
		End If
	End If
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_2").JavaList("Tipo 3:").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Crear caso de facturación"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCrearCaso.png", True
			imagenToWord "Error en la Carga de pantalla Crear caso de facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_2").JavaList("Tipo 3:").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_2").JavaList("Tipo 3:").Select "Compra de Paquete de Datos"
		wait 2				
	End If
			
			
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_2").JavaEdit("Monto reclamado").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_2").JavaEdit("Monto reclamado").Set e_Monto
	End If
			
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_2").JavaList("Razón del crédito:").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_2").JavaList("Razón del crédito:").Select "Reclamo de Facturación Cargo"
	End If
	wait 2		
			
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_2").JavaList("ComboBoxNative$1").Select "Guardar y Ver Detalles" @@ hightlight id_;_20516787_;_script infofile_;_ZIP::ssf3.xml_;_
		
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CrearCasoFact.png", True
	imagenToWord "Creación de caso de Facturación",RutaEvidencias() &Num_Iter&"_"&"CrearCasoFact.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_2").JavaButton("Guardar").Click
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaButton("Resetear").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Atributos Flexibles"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorAtrFlexibles.png", True
			imagenToWord "Error en la Carga de pantalla Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"ErrorAtrFlexibles.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaButton("Guardar").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Atributos Flexibles"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorAtrFlexibles.png", True
			imagenToWord "Error en la Carga de pantalla Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"ErrorAtrFlexibles.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Atributos Flexibles"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorAtrFlexibles.png", True
			imagenToWord "Error en la Carga de pantalla Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"ErrorAtrFlexibles.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
			
'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").SelectRow "#0"

		
	'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").DoubleClickCell "#0","#1"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#0","#1", "Pruebas"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").DoubleClickCell "#1","#1"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").SetCellData "#1","#1","1.DNI"
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").DoubleClickCell "#2","#1"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").SetCellData "#2","#1","19-Pasco"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").DoubleClickCell "#3","#1"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").SetCellData "#3","#1","04-Traslado"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").DoubleClickCell "#4","#1"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").SetCellData "#4","#1","04-Larga distancia nacional"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").DoubleClickCell "#5","#1"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").SetCellData "#5","#1","947908500"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").DoubleClickCell "#6","#1"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").SetCellData "#6","#1","09-Otros" @@ hightlight id_;_87502_;_script infofile_;_ZIP::ssf10.xml_;_
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").DoubleClickCell "#7","#1"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").SetCellData "#7","#1","09-Otros" @@ hightlight id_;_87502_;_script infofile_;_ZIP::ssf11.xml_;_
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").DoubleClickCell "#8","#1"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").SetCellData "#8","#1",e_Monto @@ hightlight id_;_87502_;_script infofile_;_ZIP::ssf12.xml_;_
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{DOWN}"
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").DoubleClickCell "#9","#1"
	wait 1
	'JavaWindow("Ejecutivo de interacción").InsightObject("InsightObject").Click
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").SetCellData "#9","#1","Sí"

'Set shell = CreateObject("Wscript.Shell") 
'	shell.SendKeys "{DOWN}"
	'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").DoubleClickCell "#10","#1"
	wait 1
		
'JavaWindow("Ejecutivo de interacción").InsightObject("InsightObject_2").Click
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").SetCellData "#10","#1",fechaSistema()	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").DoubleClickCell "#11","#1"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaTable("SearchJTable").SetCellData "#11","#1",e_Factura




	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaButton("Guardar").Exist = True Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"AtrFlexibles.png", True
		imagenToWord "Se visualiza la pantalla Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"AtrFlexibles.png"
	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación_4").JavaButton("Guardar").Click
	End If

End Sub	
Sub ValidacionDisputa
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaList("Estado").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png", True
			imagenToWord "Error en la Carga de pantalla Ver caso",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaEdit("Monto aprobado:").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png", True
			imagenToWord "Error en la Carga de pantalla Ver caso",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
		End If	
		Wend
							
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaList("Estado").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaList("Estado").Select "Fundado"
		
	End If
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaEdit("Monto aprobado:").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaEdit("Monto aprobado:").Set e_Monto
	End If
	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaButton("Guardar").Exist = True Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CasoFundado1.png", True
		imagenToWord "Caso Fundado",RutaEvidencias() &Num_Iter&"_"&"CasoFundado1.png"
	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaButton("Guardar").Click
	End If @@ hightlight id_;_13683017_;_script infofile_;_ZIP::ssf13.xml_;_
'
wait 5
'	t=0
'	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860 > Atributos").JavaButton("Guardar").Exist = False
'		wait 2
'				
'		t = t + 1
'		If (t >= 180) Then
'			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
'			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver caso"
'			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png", True
'			imagenToWord "Error en la Carga de pantalla Ver caso",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png"
'			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
'			Wait 2
'			ExitActionIteration
'					
'		End If	
'	Wend
			
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860 > Atributos").JavaButton("Guardar").Exist =True Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"EvAtributosFlexibles.png", True
		imagenToWord "Visualizacion de los atributos flexibles",RutaEvidencias() &Num_Iter&"_"&"EvAtributosFlexibles.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860 > Atributos").JavaButton("Guardar").Click
	End If

	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaButton("Guardar").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png", True
			imagenToWord "Error en la Carga de pantalla Ver caso",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaButton("Cerrar el caso").Exist = False
		Wait 2		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png", True
			imagenToWord "Error en la Carga de pantalla Ver caso",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
	
	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaButton("Cerrar el caso").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaButton("Cerrar el caso").Click
	End If

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Guardar el formulario").Exist= True Then
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Guardar el formulario").JavaButton("Guardar").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Guardar el formulario").JavaButton("Guardar").Click
		End If
	End If

	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860 > Cerrar").JavaList("Estado").Exist = False
		Wait 2
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Cerrar caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCerrarCaso.png", True
			imagenToWord "Error en la Carga de pantalla Cerrar caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCerrarCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
		
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860 > Cerrar").JavaList("Resolución:").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Cerrar caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCerrarCaso.png", True
			imagenToWord "Error en la Carga de pantalla Cerrar caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCerrarCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend	
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860 > Cerrar").JavaButton("Cerrar el caso").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Cerrar caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCerrarCaso.png", True
			imagenToWord "Error en la Carga de pantalla Cerrar caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCerrarCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
			
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860 > Cerrar").JavaButton("Cerrar el caso").Exist = True Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CerrarCaso2.png", True
		imagenToWord "Caso Cerrado",RutaEvidencias() &Num_Iter&"_"&"CerrarCaso2.png"
		Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860 > Cerrar").JavaButton("Cerrar el caso").Click
	End If
 @@ hightlight id_;_16324349_;_script infofile_;_ZIP::ssf14.xml_;_
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Facturación").JavaMenu("Cuentas financieras").Select
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaEdit("TextFieldNative$1").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Buscar Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Acuerdo de Facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaButton("Buscar ahora").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Buscar Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Acuerdo de Facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaEdit("TextFieldNative$1").Exist = True Then
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaButton("Buscar ahora").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaEdit("TextFieldNative$1").Set e_IdAcuerdoFact
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaButton("Buscar ahora").Click
		End If
	End If
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaTable("SearchJTable").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Buscar Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Acuerdo de Facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaTable("SearchJTable").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaTable("SearchJTable").SelectRow "#0"
		wait 2
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{ENTER}"
	End If
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("Buscar ahora").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Cuenta Financiera",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("StyleAuxTabbedPaneUI$1UIButton").Click @@ hightlight id_;_21031494_;_script infofile_;_ZIP::ssf15.xml_;_
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("StyleAuxTabbedPaneUI$1UIButton").Click 5,12,"LEFT" @@ hightlight id_;_21031494_;_script infofile_;_ZIP::ssf16.xml_;_
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaMenu("Créditos").Select @@ hightlight id_;_12835614_;_script infofile_;_ZIP::ssf17.xml_;_
		
	wait 2
	
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaTable("Hasta:").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Cuenta Financiera",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("Buscar ahora").Exist = True then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("Buscar ahora").Click 
			wait 5
	End If
		
	wait 2
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("<html>Contraer la fila").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("<html>Contraer la fila").Click
	End If
		
		s_Factura = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaTable("Hasta:").GetCellData(0,"#1")
		DataTable("s_IdFacturaGenerada", dtlocalSheet) = s_Factura
		DataTable("s_Resultado", dtlocalSheet) = "Exitoso"
		DataTable("s_Detalle", dtlocalSheet) = "Se ha creado el credito correctamente"
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"VerCuentaFinanciera.png", True
		imagenToWord "Se realizo el credito correctamente",RutaEvidencias() &Num_Iter&"_"&"VerCuentaFinanciera.png"
		Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
		
	wait 2
End Sub
Sub BuscarCuentaFinanciera
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Facturación").JavaMenu("Cuentas financieras").Select
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaEdit("TextFieldNative$1").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Buscar Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Acuerdo de Facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaButton("Buscar ahora").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Buscar Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Acuerdo de Facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	WAIT 2
	e_IdAcuerdoFact =  DataTable("e_IdAcuerdo", dtLocalSheet)
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaEdit("TextFieldNative$1").Exist = True Then
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaButton("Buscar ahora").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaEdit("TextFieldNative$1").Set e_IdAcuerdoFact
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaButton("Buscar ahora").Click
		End If
	End If
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaTable("SearchJTable").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Buscar Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Acuerdo de Facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaTable("SearchJTable").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaTable("SearchJTable").SelectRow "#0"
		wait 2
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{ENTER}"
	End If

	
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("Buscar ahora").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Cuenta Financiera",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	
End Sub
Sub DisputaCuentaFinanciera
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaTab("Monto anterior vencido").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Cuenta Financiera",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaTab("Monto anterior vencido").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaTab("Monto anterior vencido").Select "Lista de Facturas"	
	End If
	
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaEdit("TextFieldNative$1").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Cuenta Financiera",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	

	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("Buscar ahora").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Cuenta Financiera",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	wait 3
'	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaEdit("TextFieldNative$1").Exist = True Then
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaEdit("TextFieldNative$1").Set e_Factura
'	End If @@ hightlight id_;_10847548_;_script infofile_;_ZIP::ssf22.xml_;_
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaEdit("TextFieldNative$1_2").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaEdit("TextFieldNative$1_2").Set e_Factura
	End If
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("Buscar ahora").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaEdit("Desde:").Set "01/01/18" 
		
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("Buscar ahora").Click
	End If
	
	wait 3
	
	if JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("Número de Factura legal").Exist = True Then
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaTable("Hasta:").Exist = True Then
		'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaEdit("TextFieldNative$1").Set e_Factura

			CantFilas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaTable("Hasta:").GetROProperty("rows")
			wait 1
		End If
	End If

	
	If CantFilas = 1 Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaTable("Hasta:").SelectRow "#0"
		wait 2
		Periodo			= JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaTable("Hasta:").GetCellData("#0", "#3")
		VencimientoFact = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaTable("Hasta:").GetCellData("#0", "#8")

		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("Ingresar Disputa").Click

	End If
End Sub	
Sub GeneracionDisputaCF
	wait 2
		

	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_2").JavaButton("Seleccione Correo electrónico").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Crear caso de facturación"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCrearCaso.png", True
			imagenToWord "Error en la Carga de pantalla Crear caso de facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_2").JavaEdit("Monto reclamado").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Crear caso de facturación"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCrearCaso.png", True
			imagenToWord "Error en la Carga de pantalla Crear caso de facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	

	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_2").JavaList("Tipo 3:").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Crear caso de facturación"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCrearCaso.png", True
			imagenToWord "Error en la Carga de pantalla Crear caso de facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
	
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_2").JavaEdit("Nombres").Exist = False
		wait 1
	Wend
	wait 1
	Dim o
	o = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_2").JavaEdit("Nombres").GetROProperty("text")
	While o = ""
		wait 1
		o = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_2").JavaEdit("Nombres").GetROProperty("text")
	Wend
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_2").JavaButton("Seleccione Correo electrónico").Click	
	wait 2
	

	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_3").JavaTable("SearchJTable").Exist =  False
		wait 1
	Wend
	wait 3
	'If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_3").JavaTable("SearchJTable").Exist = True Then
	CantFilas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_3").JavaTable("SearchJTable").GetROProperty("rows")
	If CantFilas >= "1"  Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_3").JavaTable("SearchJTable").SelectRow "#0"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_3").JavaButton("Seleccionar").Click

	ElseIf CantFilas = 0  Then
				
		Valor = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_3").JavaEdit("Valor:").GetROProperty("text")
		If Valor = "" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_3").JavaEdit("Valor:").Set "prueba@gmail.com"
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_3").JavaButton("Agregar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_3").JavaTable("SearchJTable").SelectRow "#0"
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_3").JavaButton("Seleccionar").Click
		End If
	End If
	'End If
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_2").JavaList("Tipo 3:").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Crear caso de facturación"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCrearCaso.png", True
			imagenToWord "Error en la Carga de pantalla Crear caso de facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_2").JavaList("Tipo 3:").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_2").JavaList("Tipo 3:").Select "Anulación de Venta Equipo y/o SIM"
		wait 2				
	End If
	

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_2").JavaEdit("Monto reclamado").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_2").JavaEdit("Monto reclamado").Set e_Monto @@ hightlight id_;_1653816_;_script infofile_;_ZIP::ssf24.xml_;_
	End If
			
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_2").JavaList("Razón del crédito:").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_2").JavaList("Razón del crédito:").Select "Ajuste Comercial sin NC (SOL)"
	End If
	wait 2		
			
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_2").JavaList("ComboBoxNative$1").Select "Guardar y Ver Detalles" @@ hightlight id_;_1372108_;_script infofile_;_ZIP::ssf25.xml_;_
	
		
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CrearCasoFact.png", True
	imagenToWord "Creación de caso de Facturación",RutaEvidencias() &Num_Iter&"_"&"CrearCasoFact.png"
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_2").JavaButton("Guardar").Click
	
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaButton("Resetear").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Atributos Flexibles"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorAtrFlexibles.png", True
			imagenToWord "Error en la Carga de pantalla Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"ErrorAtrFlexibles.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
	t=0

	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaButton("Guardar").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Atributos Flexibles"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorAtrFlexibles.png", True
			imagenToWord "Error en la Carga de pantalla Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"ErrorAtrFlexibles.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend

	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Atributos Flexibles"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorAtrFlexibles.png", True
			imagenToWord "Error en la Carga de pantalla Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"ErrorAtrFlexibles.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend

	
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").ClickCell "#0","#1"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#0","#1",DataTable("e_MontoDisputa",dtlocalSheet)
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").ClickCell "#1","#1"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#1","#1","30/12/2020"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").ClickCell "#2","#1"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#2","#1",DataTable("e_Suscripcion",dtlocalSheet)
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").ClickCell "#3","#1"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#3","#1","Reclamo"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").ClickCell "#4","#1"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#4","#1",DataTable("e_FacturaLegal",dtlocalSheet)
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").ClickCell "#5","#1"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#5","#1","12/04/2020" @@ hightlight id_;_87502_;_script infofile_;_ZIP::ssf10.xml_;_
	wait 2
'	Set shell = CreateObject("Wscript.Shell") 
'	shell.SendKeys "{DOWN}"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").ClickCell "#6","#1"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#6","#1","09-Otros" @@ hightlight id_;_87502_;_script infofile_;_ZIP::ssf11.xml_;_
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").DoubleClickCell "#7","#1"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#7","#1","09-Otros" @@ hightlight id_;_87502_;_script infofile_;_ZIP::ssf12.xml_;_
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").DoubleClickCell "#8","#1"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#8","#1","15-Lima"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").DoubleClickCell "#9","#1"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#9","#1","Sí"
	
	
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").DoubleClickCell "#9","#1"
'	wait 1
'	JavaWindow("Ejecutivo de interacción").InsightObject("InsightObject").Click
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#9","#1",fechaSistema()

'Set shell = CreateObject("Wscript.Shell") 
'	shell.SendKeys "{DOWN}"
	'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").DoubleClickCell "#10","#1"
'	wait 1
'		
'	JavaWindow("Ejecutivo de interacción").InsightObject("InsightObject_2").Click
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#10","#1",fechaSistema()	
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").DoubleClickCell "#11","#1"
'	wait 1
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#11","#1",e_Factura
'	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaButton("Guardar").Exist = True Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"AtrFlexiblesCF.png", True
		imagenToWord "Se visualiza la pantalla Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"AtrFlexiblesCF.png"
	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaButton("Guardar").Click
	End If

End Sub
'
'	
'	wait 1
'	
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SelectRow "#10"
'	
'	'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").DoubleClickCell "#9","#1"
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#10","#1",""
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#10","#1",fechaSistema()
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").DoubleClickCell "#10","#1"
'	wait 1
'	
' @@ hightlight id_;_6155402_;_script infofile_;_ZIP::ssf31.xml_;_
'
'Set shell = CreateObject("Wscript.Shell") 
'		shell.SendKeys "{ENTER}"
' @@ hightlight id_;_6155402_;_script infofile_;_ZIP::ssf32.xml_;_
'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SelectRow "#10" @@ hightlight id_;_6155402_;_script infofile_;_ZIP::ssf33.xml_;_
'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SelectRow "#9" @@ hightlight id_;_6155402_;_script infofile_;_ZIP::ssf34.xml_;_
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:_5").JavaTable("SearchJTable").SetCellData "#10","#1",fechaSistema()
'
'wait 2
Sub ValidacionDisputaCF

	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaList("Estado").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png", True
			imagenToWord "Error en la Carga de pantalla Ver caso",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaEdit("Monto aprobado:").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png", True
			imagenToWord "Error en la Carga de pantalla Ver caso",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
		End If	
		Wend
							
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaList("Estado").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaList("Estado").Select "Fundado"
		
	End If
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaEdit("Monto aprobado:").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaEdit("Monto aprobado:").Set e_Monto
	End If
	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaButton("Guardar").Exist = True Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CasoFundado2.png", True
		imagenToWord "Caso Fundado",RutaEvidencias() &Num_Iter&"_"&"CasoFundado2.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaButton("Guardar").Click
	End If @@ hightlight id_;_13683017_;_script infofile_;_ZIP::ssf13.xml_;_
'
'	t=0
'	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860 > Atributos").JavaButton("Guardar").Exist = False
'		wait 2
'				
'		t = t + 1
'		If (t >= 180) Then
'			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
'			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver caso"
'			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png", True
'			imagenToWord "Error en la Carga de pantalla Ver caso",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png"
'			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
'			Wait 2
'			ExitActionIteration
'					
'		End If	
'	Wend
'			
'	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860 > Atributos").JavaButton("Guardar").Exist =True Then
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860 > Atributos").JavaButton("Guardar").Click
'	End If
'
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaButton("Guardar").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png", True
			imagenToWord "Error en la Carga de pantalla Ver caso",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaButton("Cerrar el caso").Exist = False
		Wait 2		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png", True
			imagenToWord "Error en la Carga de pantalla Ver caso",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
	
	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaButton("Cerrar el caso").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860*").JavaButton("Cerrar el caso").Click
	End If

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Guardar el formulario").Exist= True Then
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Guardar el formulario").JavaButton("Guardar").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Guardar el formulario").JavaButton("Guardar").Click
		End If
	End If

	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860 > Cerrar").JavaList("Estado").Exist = False
		Wait 2
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Cerrar caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCerrarCaso.png", True
			imagenToWord "Error en la Carga de pantalla Cerrar caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCerrarCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
		
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860 > Cerrar").JavaList("Resolución:").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Cerrar caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCerrarCaso.png", True
			imagenToWord "Error en la Carga de pantalla Cerrar caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCerrarCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend	
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860 > Cerrar").JavaButton("Cerrar el caso").Exist = False
		Wait 2
				
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Cerrar caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCerrarCaso.png", True
			imagenToWord "Error en la Carga de pantalla Cerrar caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCerrarCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
					
		End If	
	Wend
			
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860 > Cerrar").JavaButton("Cerrar el caso").Exist = True Then
		WAIT 5
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CerrarCaso1.png", True
		imagenToWord "Caso Cerrado",RutaEvidencias() &Num_Iter&"_"&"CerrarCaso1.png"
		Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 1860 > Cerrar").JavaButton("Cerrar el caso").Click
	End If

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("Cerrar").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("Cerrar").Click
	End If

	
 @@ hightlight id_;_12835614_;_script infofile_;_ZIP::ssf17.xml_;_
	 
	 JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Facturación").JavaMenu("Cuentas financieras").Select
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaEdit("TextFieldNative$1").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Buscar Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Acuerdo de Facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaButton("Buscar ahora").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Buscar Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Acuerdo de Facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaEdit("TextFieldNative$1").Exist = True Then
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaButton("Buscar ahora").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaEdit("TextFieldNative$1").Set e_IdAcuerdoFact
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaButton("Buscar ahora").Click
		End If
	End If
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaTable("SearchJTable").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Buscar Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Acuerdo de Facturación",RutaEvidencias() &Num_Iter&"_"&"ErrorBuscarCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaTable("SearchJTable").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaTable("SearchJTable").SelectRow "#0"
		wait 2
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{ENTER}"
	End If
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("Buscar ahora").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Cuenta Financiera",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	 
	 
	 
		
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("StyleAuxTabbedPaneUI$1UIButton_2").Click @@ hightlight id_;_21031494_;_script infofile_;_ZIP::ssf27.xml_;_
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("StyleAuxTabbedPaneUI$1UIButton_2").Click 2,8,"LEFT" @@ hightlight id_;_21031494_;_script infofile_;_ZIP::ssf28.xml_;_
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaMenu("Créditos").Select @@ hightlight id_;_740479_;_script infofile_;_ZIP::ssf29.xml_;_
	wait 2
	
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaTable("Hasta:").Exist = False
		Wait 2
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de pantalla Ver Cuenta Financiera"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorVerCtaFinanciera.png", True
			imagenToWord "Error al Cargar Ver Cuenta Financiera",RutaEvidencias() &Num_Iter&"_"&"ErrorVerCtaFinanciera.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
			
		End If	
	Wend
	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("Buscar ahora").Exist = True then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("Buscar ahora").Click 
			wait 5
	End If
		
	wait 2
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("<html>Contraer la fila").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("<html>Contraer la fila").Click
	End If
		
		s_Factura = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaTable("Hasta:").GetCellData(0,"#1")
		DataTable("s_IdFacturaGenerada", dtlocalSheet) = s_Factura
		DataTable("s_Resultado", dtlocalSheet) = "Exitoso"
		DataTable("s_Detalle", dtlocalSheet) = "Se ha creado el credito correctamente"
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"VerCuentaFinanciera.png", True
		imagenToWord "Se realizo el credito correctamente",RutaEvidencias() &Num_Iter&"_"&"VerCuentaFinanciera.png"
		Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
		
	wait 2		
End Sub
Sub CierraVentanas
	
	'Cierre de ventanas
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("Cerrar").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Cuenta Financiera:").JavaButton("Cerrar").Click	
		wait 1
	End If
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaButton("Cerrar").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaButton("Cerrar").Click
		wait 1	
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaButton("Cerrar").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver Acuerdo de Facturación").JavaButton("Cerrar").Click
		wait 1
	End If

End Sub	
Function fechaSistema()
	ini_year = Year(Now)
	ini_month = Month(Now)
	If (ini_month<=9) Then
		ini_month=("0"&ini_month)
	End If
	ini_day = Day(Now)
	If (ini_day<=9) Then
		ini_day=("0"&ini_day)
	End If
	ini_hour = Hour(Now)
	If (ini_hour<=9) Then
		ini_hour=("0"&ini_hour)
	End If
	ini_min = Minute(Now)
	If (ini_min<=9) Then
		ini_min=("0"&ini_min)
	End If
	ini_sec = Second(Now)
	If (ini_sec<=9) Then
		ini_sec=("0"&ini_sec)
	End If

	fechaSistema = ini_day&"/"&ini_month&"/"&ini_year&" "&ini_hour&":"&ini_min&":"&ini_sec	
End Function


