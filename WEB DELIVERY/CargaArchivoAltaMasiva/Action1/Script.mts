

Call PanelIteraccion()
Call VisionGeneralCliente()
Call SeleccionamosAcuerdo()
Call AltaMasiva()
Call SolicitudActivacionMasiva()
Call EnvioSolicitudMasiva()
Call SolicitarOrdenesMasivas()
Call BuscarSolicitudMasiva()
Call InformacionSolicitudMasiva()
Call ActividadOrden()



Sub PanelIteraccion()
	wait 1
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").Exist = false
		wait 1
	Wend
	wait 2
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "PanelInteraccion.png", True
	imagenToWord "Panel de Iteracción:", RutaEvidencias() & "PanelInteraccion.png"
	wait 2
End Sub
Sub VisionGeneralCliente()
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Visión General del Cliente").Select
	wait 1
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("BAR ID").Exist = false
		wait 1
	Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ResumenClienteEmpresarial.png", True
	imagenToWord "Resumen de Cliente Empresarial:", RutaEvidencias() & "ResumenClienteEmpresarial.png"
	wait 2

End Sub
Sub SeleccionamosAcuerdo()
	Do
	Valor = ""
		Fila = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("BAR ID").GetROProperty("rows")
		For i=0 to Fila-1 Step 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("BAR ID").ClickCell i, 1
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("BAR ID").PressKey "C",micCtrl
				JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").PressKey "V",micCtrl
				Valor = JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").GetROProperty("value")
				Valor2 = DataTable("e_NombreAcuerdoComercial", dtLocalSheet)
		If Valor = Valor2 Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("BAR ID").ActivateRow "#"&i
			Reporter.ReportNote "Se encontro el valor"
        	JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").Set ""
        	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "SeleccionamosAcuerdo.png", True
			imagenToWord "Seleccionamos Acuerdo Comercial:", RutaEvidencias() & "SeleccionamosAcuerdo.png"
			wait 2
   		Exit For
    	End If
		If i = Fila-1 Then
			Reporter.ReportNote "No se encontró el valor, se repetira la acción"
        	JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").Set ""
    	Exit For
		End If
    	JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").Set ""
		Next
	Loop While Not Valor = Valor2
	wait 2
End Sub
Sub AltaMasiva()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaButton("Alta Masiva").Click
	wait 1
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Activación masiva - Selecciona").JavaButton("<html>Dar de alta</html>").Exist = false
		wait 1
	Wend
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ActivaciónMasiva.png", True
	imagenToWord "Activación Masiva:", RutaEvidencias() & "ActivaciónMasiva.png"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Activación masiva - Selecciona").JavaButton("<html>Dar de alta</html>").Click

End Sub
Sub SolicitudActivacionMasiva()
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Solicitud de Activación").JavaEdit("TextFieldNative$1").Exist = false
		wait 1
	Wend
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Solicitud de Activación").JavaEdit("Plan Elige Todo+ S/ 199.90").Exist = False
		wait 1
	Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Solicitud de Activación").JavaEdit("Plan Elige Todo+ S/ 199.90").Set DataTable("e_CodVendedor", dtLocalSheet)
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Solicitud de Activación").JavaButton("Validar").Click
	wait 5
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Solicitud de Activación").JavaEdit("TextFieldNative$1").Set DataTable("e_ID_Reserva", dtLocalSheet)
	wait 1
	
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Solicitud de Activación").JavaButton("Seleccione Archivo").Click
	wait 1
	While JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Archivo").JavaButton("Examinar …").Exist = false
		wait 1
	Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Archivo").JavaButton("Examinar …").Click
	wait 1
	While JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Archivo").JavaDialog("Abrir").JavaEdit("Nombre del archivo:").Exist = false
		wait 1
	Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Archivo").JavaDialog("Abrir").JavaEdit("Nombre del archivo:").Set DataTable("e_RutaArchivoMasivo", dtLocalSheet)
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Archivo").JavaDialog("Abrir").JavaButton("Abrir").Click
	While JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Archivo").JavaButton("Examinar …").Exist = false
		wait 1
	Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ArchivoSeleccionado.png", True
	imagenToWord "Archivo seleccionado:", RutaEvidencias() & "ArchivoSeleccionado.png"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Archivo").JavaButton("Aceptar").Click
	wait 5
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN}"
	Dim estado
	estado = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Solicitud de Activación").JavaStaticText("Archivo se cargó y se").GetROProperty("text")
	If estado = "Archivo se cargó y se valida con éxito" Then
		DataTable("s_Resultado", dtlocalSheet) = "Exito"
		DataTable("s_Detalle", dtlocalSheet) = estado
		Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ArchivoExito.png", True
		imagenToWord "Carga del Archivo Exitoso:", RutaEvidencias() & "ArchivoExito.png"
		Else 
		DataTable("s_Resultado", dtlocalSheet) = "Fallido"
		DataTable("s_Detalle", dtlocalSheet) = estado
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ArchivoIncorrecto.png", True
		imagenToWord "Carga del Archivo Fallido:", RutaEvidencias() & "ArchivoIncorrecto.png"
		ExitActionIteration
	End If
	wait 1

End Sub
Sub EnvioSolicitudMasiva()
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Solicitud de Activación").JavaButton("Enviar solicitud masiva").Click
	wait 1
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("La confirmación de la").JavaStaticText("El archivo de activación").Exist = false
		wait 1
	Wend
	wait 1 
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Confirmacion.png", True
	imagenToWord "Confirmación de Activación Masiva:", RutaEvidencias() & "Confirmacion.png"
	wait 1
	Dim id 
	id = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("La confirmación de la").JavaStaticText("61044(st)").GetROProperty("text")
	DataTable("s_IdSolicitud", dtlocalSheet) = id
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("La confirmación de la").JavaButton("Cerrar").Click


End Sub
Sub SolicitarOrdenesMasivas()
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Solicitar Ordenes Masivas").Select
	wait 1
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar solicitudes masivas").JavaButton("Buscar ahora").Exist = false
		wait 1
	Wend
	wait 1
		
End Sub
Sub BuscarSolicitudMasiva()
	wait 1 
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar solicitudes masivas").JavaEdit("TextFieldNative$1").Set DataTable("s_IdSolicitud", dtlocalSheet)
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar solicitudes masivas").JavaButton("Buscar ahora").Click
	wait 1
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar solicitudes masivas").JavaTable("SearchJTable").Exist = false
		wait 1
	Wend
	wait 1
	Dim NumRegistros
	NumRegistros = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar solicitudes masivas").JavaButton("1 Registros").GetROProperty("text")
	If NumRegistros = "1 Registros"  Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar solicitudes masivas").JavaTable("SearchJTable").SelectRow "#0"
	End If
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "IdEncontrado.png", True
	imagenToWord "ID de Solicitud Masiva Encontrada:", RutaEvidencias() & "IdEncontrado.png"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar solicitudes masivas").JavaButton("Aceptar").Click
	wait 1
End Sub
Sub InformacionSolicitudMasiva()
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Información de solicitud").JavaEdit("<html><b>ID de Solicitud").Exist = false
		wait 1
	Wend
	Dim hab
	hab = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Información de solicitud").JavaEdit("<html><b>ID de Solicitud").GetROProperty("text")
	While hab = ""
		wait 1 
		hab = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Información de solicitud").JavaEdit("<html><b>ID de Solicitud").GetROProperty("text")
	Wend
	wait 1
	Dim CantLineas,CantLineasProcesadas,CantLineasFallidas
	CantLineas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Información de solicitud").JavaEdit("<html><b>Total de líneas:").GetROProperty("text")
	CantLineasProcesadas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Información de solicitud").JavaEdit("<html><b>Total de líneas").GetROProperty("text")
	CantLineasFallidas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Información de solicitud").JavaEdit("<html><b>Total de líneas_2").GetROProperty("text")
	While not ((CantLineas = CantLineasProcesadas) or (CantLineas = CantLineasFallidas))
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Información de solicitud").JavaButton("Actualizar").Click
		wait 5
		CantLineas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Información de solicitud").JavaEdit("<html><b>Total de líneas:").GetROProperty("text")
		CantLineasProcesadas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Información de solicitud").JavaEdit("<html><b>Total de líneas").GetROProperty("text")
		CantLineasFallidas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Información de solicitud").JavaEdit("<html><b>Total de líneas_2").GetROProperty("text")
	Wend
	If CantLineas = CantLineasProcesadas Then
		DataTable("s_Resultado", dtlocalSheet) = "Éxito"
		DataTable("s_Detalle", dtlocalSheet) = "Linea(s) procesadas con éxito"
		DataTable("s_CantLineasProcesadas", dtlocalSheet) = CantLineasProcesadas
		Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "LineasProcesadas.png", True
		imagenToWord "Lineas Procesadas con Éxito:", RutaEvidencias() & "LineasProcesadas.png"
		ElseIf CantLineas = CantLineasFallidas Then
		DataTable("s_Resultado", dtlocalSheet) = "Fallido"
		DataTable("s_Detalle", dtlocalSheet) = "Linea(s) no procesadas con éxito"
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "LineasNoProcesadas.png", True
		imagenToWord "Lineas NO Procesadas con Éxito:", RutaEvidencias() & "LineasNoProcesadas.png"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Información de solicitud").JavaButton("Detalles").Click
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles de solicitud").JavaTable("SearchJTable").Exist = false
			wait 1
		Wend
		wait 1
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "DetalleFallido.png", True
		imagenToWord "Detalle de la Solicitud Masiva:", RutaEvidencias() & "DetalleFallido.png"
		ExitActionIteration
	End If
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Información de solicitud").JavaButton("Detalles").Click
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles de solicitud").JavaTable("SearchJTable").Exist = false
		wait 1
	Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "DetalleExito.png", True
	imagenToWord "Detalle de la Solicitud Masiva:", RutaEvidencias() & "DetalleExito.png"
	wait 1
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles de solicitud").JavaTable("SearchJTable").ClickCell 0, 6
'	wait 2
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles de solicitud").JavaTable("SearchJTable").PressKey "C",micCtrl
'	JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").PressKey "V",micCtrl
'	Dim Valor
'	Valor = JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").GetROProperty("value")
'	DataTable("s_NumeroOrden", dtlocalSheet) = Valor&"A"
'	
	
End Sub


Sub ActividadOrden()

	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles de solicitud").JavaTable("SearchJTable").GetROProperty("rows")	
	filas = filas - 1
	For Iterator = 0 To filas step 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles de solicitud").JavaTable("SearchJTable").ClickCell Iterator, 6
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{ENTER}"
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 1047916A").JavaEdit("ID de acción de orden:").Exist = false
			wait 1
		Wend
		Dim h
		h = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 1047916A").JavaEdit("ID de acción de orden:").GetROProperty("text")
		While h = ""
			wait 1
			h = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 1047916A").JavaEdit("ID de acción de orden:").GetROProperty("text")
		Wend
		wait 1
		Dim t
		t = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 1047916A").JavaButton("959537A").GetROProperty("text")
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "AccionOrden.png", True
		imagenToWord "Detalle de Acción de Orden: "&t, RutaEvidencias() & "AccionOrden.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 1047916A").JavaTab("Nombre del cliente:").Select "Actividad"
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 1047916A").JavaTable("SearchJTable").Exist = false
			wait 1
		Wend
		wait 4
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Actividad.png", True
		imagenToWord "Detalle de la Actividad de la Orden: "&t, RutaEvidencias() & "Actividad.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 1047916A").Close
		wait 3
		If Iterator = filas Then
			Exit For
		End If
	Next
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles de solicitud").Close
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Información de solicitud").Close

End Sub




