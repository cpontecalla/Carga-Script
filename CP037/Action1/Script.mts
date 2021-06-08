Dim t, nroreg
Dim str_IDServicio
Dim str_EstadoServ
Dim str_Motivo
Dim str_TextoMotivo
Dim str_NomCliente
Dim str_ApeCliente
Dim str_Tipo_Doc
Dim str_Num_Doc
Dim str_TipoData

Num_Iter	     = Environment.Value("ActionIteration")
str_IDServicio   = DataTable("e_ID_Servicio",dtLocalSheet)
str_EstadoServ	 = DataTable("e_Estado",dtLocalSheet)
str_Motivo   	 = DataTable("e_Motivo",dtLocalSheet)
str_TextoMotivo  = DataTable("e_Motivo_Text",dtLocalSheet)
str_FechaReanuda = DataTable("e_Fecha_Reanudacion",dtLocalSheet)
str_TipoData 	 = DataTable("e_Tipo_De_DATA",dtLocalSheet)

Call PanelInteraccion()
Call BuscarSuscripcion()
Call DetallesProducto()
Call ActualizarAtributos()
Call ResumenOrden()
'If DataTable("e_Ambiente","Login") Then
	Call EmpujeOrden()
'End If
Call BuscarOrden()
Call DetalleActividadOrden()

Sub PanelInteraccion()
	
	t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Ver Detalles").Exist) = False
			wait 1	
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Panel de Interacción- de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png", True
				imagenToWord "No cargó la pantalla -Panel de Interacción- de manera correcta",RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
			wait 1
		Wend
		
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Exist) = False
			wait 1	
			t = t + 1
			If (t >= 180) Then
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
Sub BuscarSuscripcion()

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Ver Detalles").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"PanelInteraccion.png", True
		imagenToWord "PanelInteraccion.png",RutaEvidencias() &Num_Iter&"_"&"PanelInteraccion.png"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTab("Acciones rápidas").Select "Suscripciones"
		wait 1
	End If
	
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaRadioButton("Sólo contacto").Exist) = False
			wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Suscripciones- de manera correcta"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorPantallaSuscripciones_"&Num_Iter&".png", True
				imagenToWord "Error Pantalla Suscripciones",RutaEvidencias() & "ErrorPantallaSuscripciones_"&Num_Iter&".png"
				ExitTestIteration
			End If
		Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaRadioButton("Cliente").Set "ON"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaEdit("TextFieldNative$1").Set str_IDServicio
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaEdit("TextFieldNative$1_2").Set "Activo"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Buscar ahora").Click
	wait 2
	
	tiempo = 0
		Do 
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Buscar ahora").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Buscar ahora").Click
				nroreg=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("1 Registros").GetROProperty("attached text")
				tiempo = tiempo + 1
				Wait 1
			End If
			If (tiempo >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No se encuentra el ID_Servicio: "&str_IDServicio&" en la busqueda por Suscripción"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErroBusquedaSuscripcion.png", True
				imagenToWord "Error Busqueda Suscripcion",RutaEvidencias() &Num_Iter&"_"&"ErroBusquedaSuscripcion.png"
				Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Click
				wait 1
				If JavaWindow("Ejecutivo de interacción").JavaDialog("Guardar el formulario").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaDialog("Guardar el formulario").JavaButton("Descartar").Click
					wait 1
				End If
				ExitActionIteration
			End If
		
		Loop While Not(nroreg="1 Registros")
		
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Acciones rápidas").SelectRow "#0"
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"IDServicioEncontrado.png", True
	imagenToWord "ID de Servicio Encontrado.png",RutaEvidencias() &Num_Iter&"_"&"IDServicioEncontrado.png"
	Wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Ver Productos Asignados").Click
	Wait 2
			
End Sub
Sub DetallesProducto()

		tiempo = 0
		Do 
			tiempo = tiempo + 1
			If (tiempo >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "La pantalla -Detalles del Producto- no cargó de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorDetallesProducto.png", True
				imagenToWord "La pantalla -Detalles del Producto- no cargó de manera correcta", RutaEvidencias() &Num_Iter&"_"&"ErrorDetallesProducto.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			Else
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DetallesProducto.png", True
				imagenToWord "Pantalla -Detalles del Producto- cargó de manera correcta", RutaEvidencias() &Num_Iter&"_"&"DetallesProducto.png"
				Reporter.ReportEvent micPass, "Exito","La pantalla Detalles de Producto cargo correctamente"
			End If
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Calcular Penalidad").Exist(2)
		Wait 2
	
	If JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").JavaMenu("Suspender").GetROProperty("enabled")="1" Then
		JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").JavaMenu("Suspender").Select
	Else
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		DataTable("s_Detalle", dtLocalSheet) = "No se puede Suspender al número: "&str_IDServicio&" ya que la opción Suspender esta deshabilitada"
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OpcionSuspenderDeshabilitada.png", True
		imagenToWord "No se puede Suspender al número: "&str_IDServicio&" ya que la opción Suspender esta deshabilitada", RutaEvidencias() &Num_Iter&"_"&"OpcionSuspenderDeshabilitada.png"
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
		ExitActionIteration
	End If
	
End Sub
Sub FlujoWIC()

	If DataTable("e_WIC_ValidaCli", dtLocalsheet)="SI" Then
	End If
	
End Sub
Sub ActualizarAtributos()
	
		tiempo = 0
		Do
		tiempo = tiempo + 1 
		If tiempo >= 60 Then
		    DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		    DataTable("s_Detalle", dtLocalSheet) =  "Error al cargar la pantalla, no cargo la pantalla Detalles del producto"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) ,  DataTable("s_Detalle", dtLocalSheet)
			else
			Reporter.ReportEvent micPass, "Exito", "Cargo correctamente la pantalla Detalles del producto"
		End If
		wait 2
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").Exist) or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist))
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").Exist(1) Then
		var1=JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").JavaTable("Las siguientes acciones").GetCellData(0,0)
		var1=Replace(var1,"<html>","")
	 	var1=Replace(var1,"</html>","")
	 	DataTable("s_Resultado", dtLocalSheet) = "Fallido"
	    DataTable("s_Detalle", dtLocalSheet) = var1
	 	Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) ,DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").JavaButton("Rechazar solicitud de").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
		wait 2
		ExitActionIteration
	End If
	wait 1
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MensajeValidacion.png", True
		imagenToWord "Mensaje de Validacion", RutaEvidencias() &Num_Iter&"_"&"MensajeValidacion.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		wait 1
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		wait 1
	End If
	
	Dim Iterator
	Count = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").GetROProperty ("items count")
	For Iterator = 1 To Count-1
	 	rs = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").GetItem (Iterator)
		If rs = DataTable("e_Motivo", dtLocalSheet) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Select DataTable("e_Motivo", dtLocalSheet)
			Exit for
		ElseIf Iterator = Count-1 Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Select "Pedido de Cliente"
			Exit for
		End if	
	Next
	
	


	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Texto del motivo:").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Texto del motivo:").Set str_TextoMotivo
	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Fecha de reanudación:").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Fecha de reanudación:").Select str_FechaReanuda
		wait 1
	End If
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Activo Desde:").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Activo Desde:").Select str_FechaReanuda
		wait 1
	End If
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActualizarAtributos.png", True
	imagenToWord "Actualizar Atributos", RutaEvidencias() &Num_Iter&"_"&"ActualizarAtributos.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Siguiente >").Click

		t = 0
		While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist)) = False
			Wait 1	
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Resumen de la Orden- de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorResumenOrden.png", True
				imagenToWord "No cargó la pantalla -Resumen de la Orden- de manera correcta",RutaEvidencias() &Num_Iter&"_"&"ErrorResumenOrden.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
		Wait 1

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist Then
		varprb=JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaObject("JPanel").GetROProperty("text")
		varprb=Replace(varprb,"<html>"," ")
		varprb=Replace(varprb,"</html>"," ")
		varprb=Replace(varprb,"&#8203;"," ")
		varprb=Replace(varprb,"<br>"," ")
		varprb=Replace(varprb,"<br>"," ")
		varprb=Replace(varprb,"&nbsp;"," ")
		varprb=Mid(varprb, 2, 32)
		If varprb="Problema de entrada del usuario" Then
			DataTable("s_Resultado",dtLocalSheet) = "Fallido"
			DataTable("s_Detalle",dtLocalSheet) = var1
			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet) 
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MensajeValidación"&".png" , True
			imagenToWord "Mensaje de Validación", RutaEvidencias() &Num_Iter&"_"&"MensajeValidación"&".png"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Cerrar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Actualizar Atributos de").JavaList("Motivo:").Select "Cancelar a Pedido de Cliente"
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Actualizar Atributos de").JavaButton("Aceptar").Click
			wait 2
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
				wait 2
			End If
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
				wait 2	
			End If
			wait 3
			ExitActionIteration
		End If
	End If

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(2) Then
		var1=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
		var1 = Replace(var1,"<html>","")
		var1 = Replace(var1,"</html>","")
		DataTable("s_Resultado",dtLocalSheet) = "Fallido"
		DataTable("s_Detalle",dtLocalSheet) = var1
		Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet) 
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MensajeValidación"&".png" , True
		imagenToWord "Mensaje de Validación", RutaEvidencias() &Num_Iter&"_"&"MensajeValidación"&".png"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Actualizar Atributos de").JavaList("Motivo:").Select "Cancelar a Pedido de Cliente"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Actualizar Atributos de").JavaButton("Aceptar").Click
		wait 2
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
			wait 2
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
			wait 2	
		End If
		wait 3
		ExitActionIteration
	End If
	''En la pantalla "Resumen de la Orden"
	''Valida que en el arbol esten todos elementos en estado "Reanudado"
	'cantFilas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").GetROProperty ("rows")
	'For i = 2 To cantFilas -2 Step 1
	'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").SelectRow "#"&i
	'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").PressKey "C",micCtrl
	'	JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").PressKey "V",micCtrl
	'	valor = JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").GetROProperty ("text")
	'	flag = InStr(valor, "Suspendido")
	'	If flag = 0 Then
	'		Reporter.ReportEvent micFail, "Estados", "Uno de los elementos no quedo en estado Suspendido"
	'	End If
	'	JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").Set ""
	'Next
End Sub
Sub ResumenOrden()
'
'		tiempo = 0
'		Do
'			'While((JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist)) = False
'			While((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist)) = False
'				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validar y Ver Contrato").Click
'				tiempo = tiempo + 1
				wait 3
'			Wend
'			If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist Then
'				wait 3
'				var1= JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("attached text")
'				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MensajeContrato.png", True
'				imagenToWord "Mensaje Contrato", RutaEvidencias() &Num_Iter&"_"&"MensajeContrato.png"
'				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
'				wait 2
'			End If
'			wait 1
'			If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist(2) Then
'				JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
'				wait 2
'			End If
'			If tiempo>=60 Then
'				DataTable("s_Detalle",dtLocalSheet) = "Fallido"
'				DataTable("s_Resultado",dtLocalSheet) = "Error de Contrato, no se a cargado el contrato correctamente"
'				Reporter.ReportEvent micFail, DataTable("s_Detalle",dtLocalSheet), DataTable("s_Resultado",dtLocalSheet)
'				ExitActionIteration
'			else
'				Reporter.ReportEvent micPass,"Contrato Exitoso","Se a cargado el contrato correctamente"
'			End If
'			wait 1
'		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist(1)) Or (var1="Contratos no Generados") Or (var1="0"))
'		wait 3
	
	tiempo=0
	Do
		tiempo=tiempo+1
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist(2) Then
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ContratoCargado.png", True
			imagenToWord "Contrato cargado",RutaEvidencias() &Num_Iter&"_"&"ContratoCargado.png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Close
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").Set "ON"
			wait 1
			Exit Do
		Else 
			Exit Do
		End If
	wait 1
	Loop While Not (tiempo=30)

	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ResumenOrden.png", True
	imagenToWord "Resumen de la Orden",RutaEvidencias() &Num_Iter&"_"&"ResumenOrden.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
	wait 2

		t = 0
		While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist)) = False
			Wait 1	
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la orden generada de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorOrdenGenerada.png", True
				imagenToWord "No cargó la orden generada de manera correcta",RutaEvidencias() &Num_Iter&"_"&"ErrorOrdenGenerada.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
		Wait 1

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Aceptar").Click
	End If

'	Select Case DataTable("e_Ambiente", "Login [Login]")
'		Case "UAT8"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),8)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
'			Wait 2
'		Case "UAT4"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
'			Wait 2
'		Case "UAT6"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
'			Wait 2
'		Case "UAT10"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
'			Wait 2
'		Case "UAT13"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
'			Wait 2
'		Case "PROD"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),10)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
'			Wait 2				
'	End Select
	Dim text
    text = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").GetROProperty("text")
    DataTable("s_Nro_Orden", dtLocalSheet) = RTRIM(LTRIM((replace(text,"Orden",""))))
   	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden Generada"&".png", True
	imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Orden Generada"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Exist(3) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
	End If
	wait 3

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
	End If
End  Sub
Sub EmpujeOrden()

	If (str_TipoData = "DATA LOGICA") Then
	
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select @@ hightlight id_;_10269309_;_script infofile_;_ZIP::ssf3.xml_;_
		
			t = 0
			While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Exist) = False
				Wait 1	
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Depósito de Órdenes- de manera correcta"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorDepositoOrdenes.png", True
					imagenToWord "No cargó la pantalla -Depósito de Órdenes- de manera correcta"&Num_Iter,RutaEvidencias() &Num_Iter&"_"&"ErrorDepositoOrdenes.png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
				End If
			Wend
			Wait 1
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Select "Tareas pendientes del equipo" @@ hightlight id_;_9869075_;_script infofile_;_ZIP::ssf5.xml_;_
		Wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").SetFocus
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
		Wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
		
			tiempo = 0
			Do 
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist(1) Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
					nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("1 Registros").GetROProperty("attached text")
					tiempo = tiempo + 1
					Wait 1
				End If
				If (tiempo >= 180) Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar el Empuje de la Orden"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OrdenNoEncontrada.png", True
					imagenToWord "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar el Empuje de la Orden"&Num_Iter, RutaEvidencias() &Num_Iter&"_"&"OrdenNoEncontrada.png"
					Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
					Wait 2
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
					End If
					ExitActionIteration
				End If
			
			Loop While Not(nroreg="1 Registros")
			Wait 1
			
'			tiempo = 0
'			Do
'				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
'					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
'					Wait 2
'					tiempo = tiempo + 1
'					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").Output CheckPoint("Equipo usuario:")
'					varValidaRespuestaCumplimiento = Environment("s_ValidaManejarRespuestaCumplimiento")
'					Wait 1
'				End If
'				
'				If (tiempo >= 180) Then
'					DataTable("s_Resultado",dtLocalSheet) = "Fallido"
'					DataTable("s_Detalle",dtLocalSheet) = "La actividad 'Manejar Respuesta de Cumplimiento' no cargó"	
'					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"RespuestaCumplimientoNoCargo.png", True
'					imagenToWord "La actividad 'Manejar Respuesta de Cumplimiento' no cargó"&Num_Iter,RutaEvidencias() &Num_Iter&"_"&"RespuestaCumplimientoNoCargo.png"
'					Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
'					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
'					Wait 2
'					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
'					Wait 2
'					ExitTestIteration
'				End If 
'			Loop While Not varValidaRespuestaCumplimiento = "Manejar Respuesta de Cumplimiento"
'			Wait 2
		
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DepositoOrdenes.png", True
		imagenToWord "Depósito de Órdenes"&Num_Iter,RutaEvidencias() &Num_Iter&"_"&"DepositoOrdenes.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").SelectRow "#0"	
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Gestión manual").Click
		
			While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaList("Estado de la gestión manual:").Exist)=False
				wait 1
			Wend
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaList("Estado de la gestión manual:").Select "Cumplimiento Completo"
		Wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaList("Motivo de la gestión manual").Select "Manejo manual: Manejo Manual OSS"
		Wait 2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"GestionManual.png", True
		imagenToWord "-Estado de la gestión manual- de manera correcta"&Num_Iter,RutaEvidencias() &Num_Iter&"_"&"GestionManual.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Enviar").Click
	
			While (JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").Exist) = False
				wait 1	
			Wend
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
	End If
End  Sub
Sub BuscarOrden()
	
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Exist) = False
			Wait 1	
		Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaCheckBox("Solo órdenes pendientes").Set "OFF"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
	wait 6
		tiempo = 0
		Do
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
			Wait 5
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:")
			
			tiempo = tiempo + 1
			If (tiempo >= 60)	Then		
				'Error no cambia el estado de la orden a "Cerrado"
				DataTable("s_Resultado",dtLocalSheet) = "Fallido"
				DataTable("s_Detalle",dtLocalSheet) = "La orden no culminó en estado Cerrado, favor de revisar la orden"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OrdenNoCerro.png", True
				imagenToWord "La orden no culminó en estado Cerrado, favor de revisar la orden"&Num_Iter,RutaEvidencias() &Num_Iter&"_"&"OrdenNoCerro.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet) , DataTable("s_Detalle",dtLocalSheet)
					If DataTable("s_ValEstadoOrden", dtLocalSheet) = "Enviado" Then
						Exit Do
						wait 1
					End If	
			Else
				Reporter.ReportEvent micPass, "Se valida el estado de la orden", DataTable("s_ValEstadoOrden", dtLocalSheet)
			End If
		Loop While Not DataTable("s_ValEstadoOrden", dtLocalSheet) = "Cerrado"
		DataTable("s_Resultado",dtLocalSheet) = "Éxito"
		DataTable("s_Detalle",dtLocalSheet) = "Se realizó el Corte APC correctamente"
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OrdenCerrada.png", True
		imagenToWord "Se realizó el Corte APC correctamente"&Num_Iter,RutaEvidencias() &Num_Iter&"_"&"OrdenCerrada.png"
		Reporter.ReportEvent micPass,"Orden Finalizada","La orden finalizó correctamente"
		Wait 2
 @@ hightlight id_;_8937941_;_script infofile_;_ZIP::ssf7.xml_;_
End Sub
Sub DetalleActividadOrden()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaList("Ver por:").Select "Acciones de orden"
	wait 3
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").DoubleClickCell 0, "#8", "LEFT"
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{ENTER}"
	wait 1
	
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 789395A").JavaEdit("Fecha de vencimiento:").Exist)=False
			wait 1
		Wend
		
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 789395A").JavaTab("Nombre del cliente:").Select "Actividad"
		
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 789395A").JavaTable("SearchJTable").Exist)=False
			wait 1	
		Wend
		
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_1"&".png", True
	imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_1"&".png"
	
	shell.SendKeys "{PGDN}"
	wait 1
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_2"&".png", True
	imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_2"&".png"
	
	filas=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 789395A").JavaTable("SearchJTable").GetROProperty("rows")
	For Iterator = 0 To filas-1
		varselec=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 789395A").JavaTable("SearchJTable").GetCellData(Iterator,0)
	Next
	
	If varselec<>"Actualizar Descuento" Then
	 	DataTable("s_Resultado",dtLocalSheet)="Fallido"
		DataTable("s_Detalle",dtLocalSheet)="La orden "&DataTable("s_Nro_Orden",dtLocalSheet)&" no culmino en estado Cerrado, falló en la Actividad "&varselec&""
		Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 789395A").JavaButton("Cancelar").Click

		wait 2
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
			wait 2
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(2) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
		End If
		ExitActionIteration
		wait 1
	End If
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 789395A").JavaButton("Cancelar").Click

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
		wait 2
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
	End If
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
		wait 1
	End If
	
'	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Exist Then
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
'		wait 1
'	End If
	
End Sub


