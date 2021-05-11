Dim Num_Iter, vardepo, nroreg, varValidaRespuestaCumplimiento, cantFilas, var1
Dim shell, Iterator, filas, varselec
Dim str_IDServicio
Dim str_EstadoServ
Dim str_Motivo
Dim str_TextoMotivo
Dim str_TipoData

Num_Iter	     = Environment.Value("ActionIteration")
str_IDServicio   = DataTable("e_ID_Servicio",dtLocalSheet)
str_EstadoServ   = DataTable("e_Estado", dtLocalSheet)
str_Motivo       = DataTable("e_Motivo", dtLocalSheet)
str_TextoMotivo  = DataTable("e_Motivo_Text", dtLocalSheet)
str_TipoData     = DataTable("e_Tipo_De_DATA", dtLocalSheet)

Call PanelInteraccion()
Call ProductosAsignados()
Call DetallesProducto()
Call FlujoWIC()
Call ActualizarAtributos()
Call NegociarConfiguracion()
Call ResumenOrden()
'If DataTable("e_Ambiente", "Login") Then
	Call EmpujeOrden()
'End If
Call BuscarOrden()
Call DetalleActividadOrden()

Sub PanelInteraccion()
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Exist) = False
			Wait 1	
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Panel de Interacción- de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png", True
				imagenToWord "No cargó la pantalla -Panel de Interacción- de manera correcta",RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
	
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Buscar ahora").Exist) = False
			Wait 1	
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Panel de Interacción- de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png", True
				imagenToWord "No cargó la pantalla -Panel de Interacción- de manera correcta",RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
End Sub
Sub ProductosAsignados()
	
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Productos Asignados").Select @@ hightlight id_;_29580658_;_script infofile_;_ZIP::ssf15.xml_;_
	
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Exist) = False
			Wait 1	
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Productos asignados- de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorProductosAsignados.png", True
				imagenToWord "No cargó la pantalla -Productos asignados- de manera correcta",RutaEvidencias() &Num_Iter&"_"&"ErrorProductosAsignados.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaEdit("TextFieldNative$1_2").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaEdit("TextFieldNative$1_2").Set str_IDServicio
	
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaCheckBox("Incluir órdenes pendientes").Exist) = False
			Wait 1	
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó el control -Incluir órdenes pendientes- de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorControlIncluirOrdenesPendientes.png", True
				imagenToWord "No cargó el control -Incluir órdenes pendientes- de manera correcta",RutaEvidencias() &Num_Iter&"_"&"ErrorControlIncluirOrdenesPendientes.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
	
	If DataTable("e_ID_Servicio", dtLocalSheet) = "" Then
		DataTable("e_ID_Servicio", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaTable("Tabla").GetCellData (0,2)	
	End If
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaList("ComboBoxNative$1").Select str_EstadoServ
	Wait 2
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Click	
		
		tiempo=0
			Do 
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Click
					nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("0 Registros").GetROProperty("label")
					tiempo=tiempo+1
					wait 1
				End If
				If (tiempo >= 20) Then
						DataTable("s_Resultado", dtLocalSheet) = "Fallido"
						DataTable("s_Detalle", dtLocalSheet) = "No se encuentra el número: "&DataTable("e_ID_Servicio", dtLocalSheet)&" en la consulta"
						Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
						JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorNroBuscado.png", True
						imagenToWord "Error Nro Buscado",RutaEvidencias() &Num_Iter&"_"&"ErrorNroBuscado.png"
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
						ExitActionIteration
				End If
			Loop While Not(nroreg="1 Registros")
			wait 1

	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"IDServicioEncontrado.png", True
	imagenToWord "Id de Servicio: "&str_IDServicio&" encontrado", RutaEvidencias() &Num_Iter&"_"&"IDServicioEncontrado.png"	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaTable("Tabla").DoubleClickCell 0, "#2","LEFT"
	Wait 1
	
	'cantFilas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaTable("Tabla Detalle").GetROProperty("rows")
	''Obtiene el nuemero de IMEI
	'For i = cantFilas To 1 Step -1
	'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaTable("Tabla Detalle").SelectRow "#"&i-1
	'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaTable("Tabla Detalle").PressKey "C",micCtrl @@ hightlight id_;_19063758_;_script infofile_;_ZIP::ssf6.xml_;_
	'	JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").PressKey "V",micCtrl @@ hightlight id_;_14186289_;_script infofile_;_ZIP::ssf7.xml_;_
	'	valor = JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").GetROProperty ("value")
	'	If valor = "Número IMEI" Then
	'		DataTable("s_IMEI", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaTable("Tabla Detalle").GetCellData (i-1,2)
	'		Exit For
	'		Reporter.ReportEvent micPass, "Busqueda IMEI", "Se encontró el IMEI en el arbol y se guardó en la tabla"
	'	End If
	'	If i = 1 Then
	'		Reporter.ReportEvent micFail, "Busqueda IMEI", "No se encontró el IMEI en el arbol"
	'		ExitTest
	'	End If
	'Next
End Sub
Sub DetallesProducto()
	
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Calcular Penalidad").Exist) = False
			Wait 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "La pantalla -Detalles del Producto- no cargó de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorDetallesProducto.png", True
				imagenToWord "La pantalla -Detalles del Producto- no cargó de manera correcta", RutaEvidencias() &Num_Iter&"_"&"ErrorDetallesProducto.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
				Else
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DetallesProducto.png", True
				imagenToWord "Pantalla -Detalles del Producto- cargó de manera correcta", RutaEvidencias() &Num_Iter&"_"&"DetallesProducto.png"
				Reporter.ReportEvent micPass, "Éxito","La pantalla Detalles de Producto cargo correctamente"
			End If	
		Wend
	
	JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").Select
	If JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").JavaMenu("Reanudar").GetROProperty("enabled") = "1" Then
		JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").JavaMenu("Reanudar").Select
	Else
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
	    DataTable("s_Detalle", dtLocalSheet) = "No se puede Reanudar el Servicio al número: "&str_IDServicio&", ya que la opción Reanudar esta deshabilitada"
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OpcionReanudarDeshabilitada.png", True
		imagenToWord "No se puede Suspender al número: "&str_IDServicio&" ya que la opción Reanudar esta deshabilitada", RutaEvidencias() &Num_Iter&"_"&"OpcionSuspenderDeshabilitada.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
		ExitActionIteration
	End If
	
		tiempo = 0
		Do
			tiempo = tiempo + 1 
			If (tiempo >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) =  "Error al cargar, no cargó la pantalla Actualizar Atributos"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorActualizarAtributos.png", True
				imagenToWord "Error al cargar, no cargó la pantalla Actualizar Atributos", RutaEvidencias() &Num_Iter&"_"&"ErrorActualizarAtributos.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			Else
				Reporter.ReportEvent micPass, "Éxito", "La pantalla -Actualizar Atributos- cargó correctamente"
			End If
			Wait 2
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").Exist))
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").Exist(1) Then
		var1=JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").JavaTable("Las siguientes acciones").GetCellData(0,0)
		var1=Replace(var1,"<html>","")
		var1=Replace(var1,"</html>","")
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		DataTable("s_Detalle", dtLocalSheet) = var1
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) ,DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MensajeOrdenPendiente.png", True
		imagenToWord var1, RutaEvidencias() &Num_Iter&"_"&"MensajeOrdenPendiente.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").JavaButton("Rechazar solicitud de").Click
		Wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
		Wait 2
		ExitActionIteration
	End If
End Sub
Sub FlujoWIC()

	If DataTable("e_WIC_ValidaCli", dtLocalsheet)="SI" Then
		RunAction "WIC", oneIteration
	End If
	
End Sub
Sub ActualizarAtributos()
	
		tiempo = 0
		Do
			tiempo = tiempo + 1 
			If (tiempo >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) =  "Error al cargar, no cargó la pantalla Actualizar Atributos"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorActualizarAtributos.png", True
				imagenToWord "Error al cargar, no cargó la pantalla Actualizar Atributos", RutaEvidencias() &Num_Iter&"_"&"ErrorActualizarAtributos.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			Else
				Reporter.ReportEvent micPass, "Éxito", "La pantalla -Actualizar Atributos- cargó correctamente"
			End If
			Wait 2
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Exist))
	wait 2
	
	Dim Iterator
	Count = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").GetROProperty ("items count")
	For Iterator = 1 To Count-1
	 	rs = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").GetItem (Iterator)
		If rs =  DataTable("e_Motivo", dtLocalSheet) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Select DataTable("e_Motivo", dtLocalSheet)
			Exit for
		ElseIf Iterator = Count-1 Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Select "Pedido de Cliente"
			Exit for
		End if	
	Next
	
	Wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Texto del motivo:").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Texto del motivo:").Set str_TextoMotivo
	Wait 2
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActualizarAtributos.png", True
	imagenToWord "Actualizar Atributos", RutaEvidencias() &Num_Iter&"_"&"ActualizarAtributos.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Siguiente >").Click

End Sub
Sub NegociarConfiguracion()
	
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Exist) = False
			Wait 1	
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Negociar Configuración- de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorNegociarConfiguracion.png", True
				imagenToWord "No cargó la pantalla -Negociar Configuración- de manera correcta",RutaEvidencias() &Num_Iter&"_"&"ErrorNegociarConfiguracion.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
		Wait 1
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"NegociarConfiguracion.png", True
	imagenToWord "Actualizar Atributos", RutaEvidencias() &Num_Iter&"_"&"NegociarConfiguracion.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Click
	
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validar y Ver Contrato").Exist) = False
			wait 1
		Wend
		wait 2
	
End Sub
Sub ResumenOrden()
	
'	cantFilas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").GetROProperty ("rows")
'	For i = 2 To cantFilas -2 Step 1
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").SelectRow "#"&i
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").PressKey "C",micCtrl
'		JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").PressKey "V",micCtrl
'		valor = JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").GetROProperty ("text")
'		flag = InStr(valor, "Reanudado")
'		If flag = 0 Then
'			Reporter.ReportEvent micFail, "Estados", "Uno de los elementos no quedo en estado Reanudado"
'			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ElementoNoReanudado.png", True
'			imagenToWord "Uno de los elementos no quedo en estado Reanudado",RutaEvidencias() &Num_Iter&"_"&"ElementoNoReanudado.png"
'		End If
'		JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").Set ""
'	Next
'	Wait 4
	
		tiempo = 0
		Do
			While((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist)) = False
				wait 1
				tiempo = tiempo + 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validar y Ver Contrato").Click
				If DataTable("e_WIC_ContrCli",dtLocalSheet)="SI" Then
						RunAction "WIC2", oneIteration
					Exit Do
				End If
				
				wait 3
			Wend
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist(2) Then
				wait 3
				var1= JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("attached text")
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MensajeContrato.png", True
				imagenToWord "Mensaje Contrato", RutaEvidencias() &Num_Iter&"_"&"MensajeContrato.png"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
				wait 2
			End If
		
				If (tiempo >= 60) Then
					DataTable("s_Detalle",dtLocalSheet) = "Fallido"
					DataTable("s_Resultado",dtLocalSheet) = "Error de Contrato, no se a cargado el contrato correctamente"
					Reporter.ReportEvent micFail, DataTable("s_Detalle",dtLocalSheet), DataTable("s_Resultado",dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorContratoNoCargado.png", True
					imagenToWord "Error de Contrato, no se ha cargado el contrato correctamente",RutaEvidencias() &Num_Iter&"_"&"ErrorContratoNoCargado.png"
					ExitActionIteration
				Else
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ContratoCargado.png", True
					imagenToWord "Contrato cargado correctamente",RutaEvidencias() &Num_Iter&"_"&"ContratoCargado.png"
					Reporter.ReportEvent micPass,"Contrato Exitoso","Se ha cargado el contrato correctamente"
				End If
				Wait 1
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist(1)) Or (var1="0") Or (var1="Contratos no Generados"))
		Wait 3
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ContratoCargado.png", True
		imagenToWord "Contrato cargado",RutaEvidencias() &Num_Iter&"_"&"ContratoCargado.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Close
		Wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").Set "ON"
		Wait 1
	End If
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ResumenOrden.png", True
	imagenToWord "Resumen de la Orden",RutaEvidencias() &Num_Iter&"_"&"ResumenOrden.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
	Wait 2
	
'	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
'		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Aceptar").Click
'	End If
'	
	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").Exist) = False
		Wait 1	
	Wend
	
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
'	DataTable("s_Resultado", dtLocalSheet) = "Éxito"
'	DataTable("s_Detalle", dtLocalSheet) = "Se ejecutó la Reconexión APC correctamente"
'	
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
	Wait 2

End Sub
Sub EmpujeOrden()

	
	If str_TipoData = "DATA LOGICA" Then
	
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select @@ hightlight id_;_10269309_;_script infofile_;_ZIP::ssf3.xml_;_
		Wait 5
		
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
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
						Wait 2
						ExitActionIteration
						Wait 2
					End If
			Loop While Not(nroreg="1 Registros")
			Wait 1
		
			tiempo = 0
			Do
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
					Wait 2
					tiempo = tiempo + 1
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").Output CheckPoint("Equipo usuario:") @@ hightlight id_;_33211627_;_script infofile_;_ZIP::ssf1.xml_;_
					varValidaRespuestaCumplimiento = Environment("s_ValidaManejarRespuestaCumplimiento")
					Wait 1
				End If
					If (tiempo >= 180) Then
						DataTable("s_Resultado",dtLocalSheet) = "Fallido"
						DataTable("s_Detalle",dtLocalSheet) = "La actividad 'Manejar Respuesta de Cumplimiento' no cargó"	
						JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"RespuestaCumplimientoNoCargo.png", True
						imagenToWord "La actividad 'Manejar Respuesta de Cumplimiento' no cargó"&Num_Iter,RutaEvidencias() &Num_Iter&"_"&"RespuestaCumplimientoNoCargo.png"
						Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
						Wait 2
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
						wait 2
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
						Wait 2
						ExitTestIteration
					End If 
			Loop While Not varValidaRespuestaCumplimiento = "Manejar Respuesta de Cumplimiento"
			Wait 2
	
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DepositoOrdenes.png", True
		imagenToWord "Depósito de Órdenes"&Num_Iter,RutaEvidencias() &Num_Iter&"_"&"DepositoOrdenes.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").SelectRow "#0"
		Wait 1
			
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Gestión manual").Click
	
			While (JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaList("Estado de la gestión manual:").Exist) = False
				Wait 1	
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
			Wait 1	
		Wend
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
	End If

End Sub
Sub BuscarOrden()

	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
		
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Exist) = False
			Wait 1	
		Wend
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
	Wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaCheckBox("Solo órdenes pendientes").Set "OFF"
	Wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
	Wait 8
	
		tiempo = 0
		Do
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
			Wait 5
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:")
			
			tiempo = tiempo +1
			If (tiempo >= 60) Then		
				DataTable("s_Resultado",dtLocalSheet) = "Fallido"
				DataTable("s_Detalle",dtLocalSheet) = "La orden no culmino en estado Cerrado, favor de revisar la orden"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OrdenNoCerro.png", True
				imagenToWord "La orden no culminó en estado Cerrado, favor de revisar la orden"&Num_Iter,RutaEvidencias() &Num_Iter&"_"&"OrdenNoCerro.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet) , DataTable("s_Detalle",dtLocalSheet)
					If DataTable("s_ValEstadoOrden", dtLocalSheet) = "Enviado" Then
						Exit Do
						wait 1
					End If
				'ExitActionIteration
			Else
				Reporter.ReportEvent micPass, "Se valida el estado de la orden", DataTable("s_ValEstadoOrden", dtLocalSheet)
			End If
		Loop While Not DataTable("s_ValEstadoOrden", dtLocalSheet) = "Cerrado"
		DataTable("s_Resultado",dtLocalSheet) = "Éxito"
		DataTable("s_Detalle",dtLocalSheet) = "Se realizó la Reconexión APC correctamente"
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OrdenCerrada.png", True
		imagenToWord "Se realizó el Corte APC correctamente"&Num_Iter,RutaEvidencias() &Num_Iter&"_"&"OrdenCerrada.png"
		Reporter.ReportEvent micPass,"Orden Finalizada","La orden finalizó correctamente"
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
	
	If varselec<>"Cerrar Acción de Orden" Then
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
		wait 1
	End If
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
		wait 1
	End If
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
		wait 1
	End If

End Sub
