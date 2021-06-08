Dim shell, filas, Iterator, varselec, varhab
Dim str_EstControl
Dim str_EstServicio
Dim str_IDServicio
Dim str_Motivo

str_EstServicio = DataTable("e_Estado", dtLocalSheet)
str_IDServicio  = DataTable("e_ID_Servicio",dtLocalSheet)
str_Motivo      = DataTable("e_Motivo", dtLocalSheet)
str_TxtMotivo   = DataTable("e_Motivo_Text", dtLocalSheet)
str_TipoData    = DataTable("e_Tipo_De_DATA", dtLocalSheet)

Call PanelInteraccion()
Call ProductosAsignados()
Call DetallesProducto()
Call ActualizarAtributos()
Call ResumenOrden()
'If DataTable("e_Ambiente", "Login")<>"PROD" Then
	Call EmpujeOrden()	
'End If
Call ValidaOrden()
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
			imagenToWord "No cargó la pantalla -Panel de Interacción- de manera correcta.png",RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
	Wend 
	wait 1
	
	t = 0
	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Ver todo").Exist) = False
		wait 1
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Panel de Interacción- de manera correcta"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png", True
			imagenToWord "No cargó la pantalla -Panel de Interacción- de manera correcta.png",RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
	Wend 
	wait 1
	
	t = 0
	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Buscar ahora").Exist) = False
		wait 1
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Panel de Interacción- de manera correcta"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png", True
			imagenToWord "No cargó la pantalla -Panel de Interacción- de manera correcta.png",RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
	Wend 
	wait 1
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Ver Detalles").Exist(1) Then
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Buscar ahora").Exist(1) Then
				JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select	
					str_EstControl = JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").GetROProperty("enabled")			
					If (str_EstControl = "1") Then
						JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Productos Asignados").Select @@ hightlight id_;_29580658_;_script infofile_;_ZIP::ssf15.xml_;_
						wait 2
					End If
			End If
	End If
	
End Sub
Sub ProductosAsignados()
	
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaEdit("TextFieldNative$1_2").Exist) = False
			wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Productos Asignados- de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorProductosAsignados.png", True
				imagenToWord "No cargó la pantalla -Productos Asignados- de manera correcta.png",RutaEvidencias() &Num_Iter&"_"&"ErrorProductosAsignados.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend 
		wait 1
	
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaTable("Tabla").Exist) = False
			wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Productos Asignados- de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorProductosAsignados.png", True
				imagenToWord "No cargó la pantalla -Productos Asignados- de manera correcta.png",RutaEvidencias() &Num_Iter&"_"&"ErrorProductosAsignados.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend 
		wait 1
	
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Exist) = False
			wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Productos Asignados- de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorProductosAsignados.png", True
				imagenToWord "No cargó la pantalla -Productos Asignados- de manera correcta.png",RutaEvidencias() &Num_Iter&"_"&"ErrorProductosAsignados.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend 
		wait 1
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaEdit("TextFieldNative$1_2").Exist(1) Then
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Exist(1) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaEdit("TextFieldNative$1_2").SetFocus
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaEdit("TextFieldNative$1_2").Set DataTable("e_ID_Servicio",dtLocalSheet) @@ hightlight id_;_10077622_;_script infofile_;_ZIP::ssf16.xml_;_
			wait 2
		End If
	End If
	
	If DataTable("e_ID_Servicio", dtLocalSheet) = "" Then
		DataTable("e_ID_Servicio", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaTable("Tabla").GetCellData (0,2)	
	End If
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaList("ComboBoxNative$1").Select str_EstServicio
	wait 2

	tiempo=0
		Do 
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Click
				varbusq=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("0 Registros").GetROProperty("attached text")
				tiempo=tiempo+1
				wait 1
			End If
			If (tiempo >= 20) Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No se encuentra el Id de Servicio: "&str_IDServicio&" en el estado Activo" 
					Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorBusquedaIDServicio.png", True
					imagenToWord "No cargó el ID de Servicio en Búsqueda.png",RutaEvidencias() &Num_Iter&"_"&"ErrorBusquedaIDServicio.png"
					ExitActionIteration
			End If
		Loop While Not(varbusq="1 Registros")
		wait 1

		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"IDServicioEncontrado.png", True
		imagenToWord "ID de Servicio Encontrado.png",RutaEvidencias() &Num_Iter&"_"&"IDServicioEncontrado.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaTable("Tabla").DoubleClickCell 0, "#2","LEFT"
		wait 4
End Sub
Sub DetallesProducto()
		
		tiempo = 0
		Do
			tiempo = tiempo + 1 
			If (tiempo >= 180) Then
			    DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			    DataTable("s_Detalle", dtLocalSheet) =  "Error al cargar la pantalla, no cargo la pantalla Detalles del producto"
			    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorDetallesProducto.png", True
				imagenToWord "No cargó la pantalla -Detalles del Producto- de manera correcta.png",RutaEvidencias() &Num_Iter&"_"&"ErrorDetallesProducto.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			Else
				Reporter.ReportEvent micPass, "Éxito", "Cargó correctamente la pantalla Detalles del producto"
			End If
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Calcular Penalidad").Exist
		
	JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").Select
	wait 1
	If (JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").JavaMenu("Dar de baja").GetROProperty("enabled") = "1") Then
		JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").JavaMenu("Dar de baja").Select
	Else
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
	    DataTable("s_Detalle", dtLocalSheet) = "No se puede dar de baja al número: "&str_IDServicio&", ya que la opción Dar de Baja esta deshabilitada"
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
		ExitActionIteration
	End If
	
End Sub
Sub ActualizarAtributos()
	
		tiempo = 0
		Do
			wait 1
			tiempo = tiempo + 1 
			If (tiempo >= 80) Then
			    DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			    DataTable("s_Detalle", dtLocalSheet) =  "Error al cargar la pantalla, no cargó la pantalla Actualizar Atributos"
			    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorActualizarAtributos.png", True
				imagenToWord "No cargó la pantalla -Actualizar Atributos- de manera correcta.png",RutaEvidencias() &Num_Iter&"_"&"ErrorActualizarAtributos.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			Else
				Reporter.ReportEvent micPass, "Éxito", "Cargó correctamente la pantalla Actualizar Atributos"
			End If
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").Exist) or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Exist))
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").Exist(1) Then
		var1=JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").JavaTable("Las siguientes acciones").GetCellData(0,0)
		var1=Replace(var1,"<html>","")
	 	var1=Replace(var1,"</html>","")
	 	DataTable("s_Resultado", dtLocalSheet) = "Fallido"
	    DataTable("s_Detalle", dtLocalSheet) = var1
	    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorOrdenPendiente.png", True
		imagenToWord "Error existe una Orden Pendiente.png",RutaEvidencias() &Num_Iter&"_"&"ErrorOrdenPendiente.png"
	 	Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").JavaButton("Rechazar solicitud de").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
		wait 2
		ExitActionIteration
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Exist(2) Then
		Dim x, count
		count = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").GetROProperty ("items count")
		For Iterator = 0 To count-1 Step 1
			x = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").GetItem ("#"&Iterator)
			If x = str_Motivo Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Select str_Motivo
				Exit for 
			End If
			If Iterator = count-1  Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Select "Pedido de Cliente"
			End If
		Next
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Texto del motivo:").SetFocus
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Texto del motivo:").Set str_TxtMotivo
		wait 2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActualizarAtributos.png", True
		imagenToWord "Actualizar Atributos.png",RutaEvidencias() &Num_Iter&"_"&"ActualizarAtributos.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Siguiente >").Click
	End If
	wait 2
	
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist) = False
			Wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No se habilitó el botón -Validade y Ver Contrato-"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorbtnValidaVerContrato_"&Num_Iter&".png", True
				imagenToWord "No se habilitó el botón -Validade y Ver Contrato_"&Num_Iter,RutaEvidencias() & "ErrorbtnValidaVerContrato_"&Num_Iter&".png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
		Wait 1
	
	''En la pantalla "Resumen de la Orden"
	''Valida que en el arbol esten todos elementos en estado "Reanudado"
	'cantFilas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").GetROProperty ("rows")
	'For i = 2 To cantFilas -2 Step 1
	'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").SelectRow "#"&i
	'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").PressKey "C",micCtrl
	'	wait 1
	'	JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").PressKey "V",micCtrl
	'	wait 1
	'	valor = JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").GetROProperty ("text")
	'	flag = InStr(valor, "Removido")
	'	If flag = 0 Then
	'		Reporter.ReportEvent micFail, "Estados", "Uno de los elementos no quedo en estado Removido "
	'	End If
	'	JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").Set ""
	'Next
	'wait 4
	
End Sub
Sub ResumenOrden()
	
	varhab=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").GetROProperty("enabled")
	If  varhab<>"0" Then
		tiempo = 0
			Do
				While((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist)) = False
					wait 1
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Click
						tiempo = tiempo + 1
					wait 3
				Wend
			
	'			'Click "Validade y Ver Contrato"
	'			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist(2) Then
	'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Click
	'				Wait 5
	'				'WIC Genera Contrato
	'			End If
				
				If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist(3) Then
					wait 3
					var1= JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("attached text")
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Mensaje.png", True
					imagenToWord "Mensaje.png",RutaEvidencias() &Num_Iter&"_"&"Mensaje.png"
					JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
				End If
	'			If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist(2) Then
	'				JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
	'				wait 2
	'			End If
				wait 1
					If (tiempo >= 180) Then
						DataTable("s_Detalle",dtLocalSheet) = "Fallido"
						DataTable("s_Resultado",dtLocalSheet) = "No se a cargado el contrato correctamente"
						Reporter.ReportEvent micFail, DataTable("s_Detalle",dtLocalSheet), DataTable("s_Resultado",dtLocalSheet)
						ExitActionIteration
					Else
						Reporter.ReportEvent micPass,"Contrato Exitoso","Se ha cargado el contrato correctamente"
				End If
			Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist) Or (var1="0")Or (var1="Contratos no Generados"))
			wait 3
		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist(2) Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Close
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").Set "ON"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Contrato.png", True
			imagenToWord "Contrato.png",RutaEvidencias() &Num_Iter&"_"&"Contrato.png"
		End If
	End If
		
	'Bucle que espera "Enviar orden"
	t = 0
	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Exist) = False
		Wait 1
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No se habilitó el botón -Enviar orden-"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorbtnEnviarOrden_"&Num_Iter&".png", True
			imagenToWord "No se habilitó el botón -Enviar orden_"&Num_Iter, RutaEvidencias() & "ErrorbtnEnviarOrden_"&Num_Iter&".png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
	Wend
	Wait 1
	
	'Click en "Enviar orden"
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Exist(2) Then
		'Damos clic en el boton "Enviar Orden"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
		Wait 3
	End If
	
	'Bucle que espera el envío de la orden	
	While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").Exist)Or(JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist))=False
		wait 1
	Wend
	
	'Mensaje de validación
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Aceptar").Click
	End If
	
	'Captura de la orden generada
'	Select Case DataTable("e_Ambiente", "Login [Login]")
'		Case "UAT8"
'			wait 3
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),8)
'			str_NroOrden = DataTable("s_Nro_Orden", dtLocalSheet) 
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
'			wait 2
'		Case "UAT4"
'			wait 3
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			str_NroOrden = DataTable("s_Nro_Orden", dtLocalSheet)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
'			wait 2
'		Case "UAT6"
'			wait 3
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			str_NroOrden = DataTable("s_Nro_Orden", dtLocalSheet)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
'			wait 2	
'		Case "UAT10"
'			wait 3
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			str_NroOrden = DataTable("s_Nro_Orden", dtLocalSheet)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
'			wait 2	
'		Case "UAT13"
'			wait 3
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			str_NroOrden = DataTable("s_Nro_Orden", dtLocalSheet)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
'			wait 2
'		Case "PROD"
'			wait 3
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),10)
'			str_NroOrden = DataTable("s_Nro_Orden", dtLocalSheet)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
'			wait 2			
'	End Select
	Dim text
    text = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").GetROProperty("text")
    DataTable("s_Nro_Orden", dtLocalSheet) = RTRIM(LTRIM((replace(text,"Orden",""))))
    wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden Generada"&".png", True
	imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Orden Generada"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
	End If
	
End Sub
Sub EmpujeOrden()

	If str_TipoData = "DATA LOGICA" Then
	
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select @@ hightlight id_;_10269309_;_script infofile_;_ZIP::ssf3.xml_;_

		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Exist)=False
			wait 1
		Wend           
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Select "Tareas pendientes del equipo" @@ hightlight id_;_9869075_;_script infofile_;_ZIP::ssf5.xml_;_
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").SetFocus
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
		
		tiempo=0
			Do 
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
					var = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("1 Registros").GetROProperty("text")
					tiempo=tiempo+1
					wait 1
				End If
				
				If (tiempo >= 120) Then
						DataTable("s_Resultado", dtLocalSheet) = "Fallido"
						DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar el Empuje de la Orden"
						Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
						JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorOrdenBuscada_"&Num_Iter&".png", True
						imagenToWord "Grupo de Órdenes no encuentra datos: "&DataTable("s_Nro_Orden", dtLocalSheet), RutaEvidencias() & "ErrorOrdenBuscada_"&Num_Iter&".png"
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
						wait 2
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
						wait 2
						ExitActionIteration
						wait 2
				End If
			Loop While Not(var="1 Registros")
			wait 1
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").SelectRow "#0"
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "OrdenEncontrada.png", True
		imagenToWord "Orden Encontrada",RutaEvidencias() & "OrdenEncontrada.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Gestión manual").Click
	
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaList("Estado de la gestión manual:").Exist) = False
			wait  1
		Wend
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaList("Estado de la gestión manual:").Select "Cumplimiento Completo"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaList("Motivo de la gestión manual").Select "Manejo manual: Manejo Manual OSS"
		wait 1
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "EstadoGestionManual.png", True
		imagenToWord "Estado Gestion Manual",RutaEvidencias() & "EstadoGestionManual.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Enviar").Click
		
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist)=False
			wait 1
		Wend
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
		End If
		
	End If
End Sub
Sub ValidaOrden()
	
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Exist)= False
		wait 1
	Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaCheckBox("Solo órdenes pendientes").Set "OFF"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
	
		tiempo=0
		Do 
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
				varbusqord = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("1 Registros").GetROProperty("attached text")
				tiempo=tiempo+1
				wait 1
			End If
			
			If (tiempo >= 180) Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar el Cierre de la Orden"
					Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "OrdenNoEncontrada.png", True
					imagenToWord "Orden No Encontrada",RutaEvidencias() & "OrdenNoEncontrada.png"
					
					ExitActionIteration
					wait 2
			End If
		Loop While Not(varbusqord="1 Registros")
		wait 1

	
		tiempo = 0
		Do
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:") @@ hightlight id_;_24634058_;_script infofile_;_ZIP::ssf1.xml_;_
			tiempo = tiempo + 1
				If (tiempo>=120) Then		
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "La Orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" no culmino en estado Cerrado"
					Reporter.ReportEvent micFail,"Error al finalizar la orden","Es probable que la orden termine con tiempo excedido"
					If DataTable("s_ValEstadoOrden", dtLocalSheet) = "Enviado" Then
						Exit Do
						wait 1
					End If	
					'ExitActionIteration
				else
				Reporter.ReportEvent micPass, "Se valida el estado de la orden",  DataTable("s_ValEstadoOrden", dtLocalSheet)
				End If
			Loop While Not DataTable("s_ValEstadoOrden", dtLocalSheet) = "Cerrado"
			DataTable("s_Resultado", dtLocalSheet)="Exito"
			DataTable("s_Detalle", dtLocalSheet)="La orden finalizó correctamente"
			Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)

		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Cerrada_"&".png", True
		imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"Orden_Cerrada_"&".png"
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
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
	End If
	
End Sub


