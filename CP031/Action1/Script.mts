Dim var1, varbusqord, shell, filas, Iterator, varselec, varhab
Dim str_EstControl
Dim str_IDServicio
Dim str_Motivo
Dim str_TxtMotivo
Dim str_TipoData
Dim str_NroOrden

str_IDServicio = DataTable("e_ID_Servicio", dtLocalSheet)
str_Estado     = DataTable("e_Estado", dtLocalSheet)
str_Motivo     = DataTable("e_Motivo", dtLocalSheet)
str_TxtMotivo  = DataTable("e_Motivo_Text", dtLocalSheet)
str_TipoData   = DataTable("e_Tipo_De_DATA", dtLocalSheet)

Call PanelInteraccion()
Call IngresoNumero()
'Call ProductosAsignados()
Call DetallesProducto()
Call FlujoWic()
Call ActualizarAtributos()
Call ParametrizaUsuario()
Call NegociarConfiguracion()
Call NegociarDistribucion()
Call ResumenOrden()
'If DataTable("e_Ambiente","Login") Then
	Call EmpujaOrden()
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
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaStaticText("Número de documento(st)").Exist(1) Then
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Ver Detalles").Exist(2) Then
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Ver todo").Exist(1) Then
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Buscar ahora").Exist(1) Then
					JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
					str_EstControl = JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Productos Asignados").GetROProperty("enabled")
				End If
			End If
		End If
	End If
	
'	If (str_EstControl = "1") Then
'		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
'		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Productos Asignados").Select @@ hightlight id_;_29580658_;_script infofile_;_ZIP::ssf15.xml_;_
'		wait 2
'	End If
	
End Sub
Sub IngresoNumero()
		wait 5
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTab("Acciones rápidas").Select "Suscripciones"
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaEdit("TextFieldNative$1").Exist) = False
			wait 1
		Wend
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaEdit("TextFieldNative$1").Set str_IDServicio
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Buscar ahora").Click
	
		tiempo=0
		Do 
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Buscar ahora").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Buscar ahora").Click
				nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("1 Registros").GetROProperty("attached text")
				tiempo=tiempo+1
				wait 1
			End If
				If (tiempo >= 180) Then
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Num_Registro_"&Num_Iter&".png", True
					imagenToWord "Error_Num_Registro_"&Num_Iter,RutaEvidencias() & "Num_Registro_"&Num_Iter&".png"
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "Tiene muchos Registros, que se procedió a detener el flujo."
					Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
				End If
		Loop While Not(nroreg="1 Registros")
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Acciones rápidas").ClickCell 0, "#0", "LEFT"
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "BusquedaSuscripcion.png", True
		imagenToWord "Busqueda de Suscripcion",RutaEvidencias() & "BusquedaSuscripcion.png"
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Ver Productos Asignados").Click
		
End Sub
Sub ProductosAsignados()
	
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

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaEdit("TextFieldNative$1_2").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaEdit("TextFieldNative$1_2").Set str_IDServicio @@ hightlight id_;_10077622_;_script infofile_;_ZIP::ssf16.xml_;_
	wait 2
	
	'Si no se completa el parametro p_ID_Servicio, el script toma automaticamente el primera de la tabla
'	If DataTable("e_ID_Servicio", dtLocalSheet) = "" Then
'		DataTable("e_ID_Servicio", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaTable("Tabla").GetCellData (0,2)	
'	End If
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaList("ComboBoxNative$1").Select str_Estado
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Check CheckPoint("Buscar ahora_2") @@ hightlight id_;_25359929_;_script infofile_;_ZIP::ssf18.xml_;_
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Click
	wait 2
	
		tiempo=0
		Do 
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Click
				nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("0 Registros").GetROProperty("label")
				tiempo=tiempo+1
				wait 1
			End If
			If (tiempo >= 30) Then
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
	imagenToWord "ID de Servicio Encontrado.png",RutaEvidencias() &Num_Iter&"_"&"IDServicioEncontrado.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaTable("Tabla").DoubleClickCell 0, "#2","LEFT"
	wait 4

End  Sub
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
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Calcular Penalidad").Exist
		Wait 2
	
	
	JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").Select
	wait 2
	If (JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").JavaMenu("Reinstalación").GetROProperty("enabled")  = "1") Then
		JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").JavaMenu("Reinstalación").Select
	Else
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
	    DataTable("s_Detalle", dtLocalSheet) = "No se puede dar de baja al número: "&str_IDServicio&", ya que la opción Reinstalación  esta deshabilitada"
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OpcionReinstalaciónDeshabilitada.png", True
		imagenToWord "No se puede Reinstalar el número: "&str_IDServicio&" ya que la opción 'Reinstalación' esta deshabilitada", RutaEvidencias() &Num_Iter&"_"&"OpcionReinstalaciónDeshabilitada.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
		ExitActionIteration
	End If
	
		tiempo = 0
		Do
			wait 1
			tiempo = tiempo + 1 
			If (tiempo >= 180) Then
			    DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			    DataTable("s_Detalle", dtLocalSheet) =  "Error al cargar la pantalla, no cargó la pantalla Detalles del producto"
			    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorActualizarAtributos.png", True
				imagenToWord "No cargó la pantalla -Actualizar Atributos- de manera correcta.png",RutaEvidencias() &Num_Iter&"_"&"ErrorActualizarAtributos.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			Else
				Reporter.ReportEvent micPass, "Éxito", "Cargó correctamente la pantalla Detalles del producto"
			End If
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").Exist))
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").Exist(1) Then
		var1=JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").JavaTable("Las siguientes acciones").GetCellData(0,0)
		var1=Replace(var1,"<html>","")
	 	var1=Replace(var1,"</html>","")
	 	DataTable("s_Resultado", dtLocalSheet) = "Fallido"
	    DataTable("s_Detalle", dtLocalSheet) = var1
	 	Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").JavaButton("Rechazar solicitud de").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
		wait 2
		ExitActionIteration
	End If
	
End Sub
Sub FlujoWic()
	
	If DataTable("e_WIC_ValidaCli",dtLocalSheet)="SI" Then
		RunAction "WIC_1", oneIteration
	End  If
'		'En la ventana Sistema de información y validación del cliente
'		'Damos clic en el boton "Continuar"
'		If UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar").UIAButton("Continuar").Exist(1) Then
'			UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar").UIAButton("Continuar").Click
'		End If
'		
'		'Seleccionamos la opción "Boleta de pago - Fijo"
'		UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAComboBox("Tipo de sustento:").UIAButton("Abrir").Click
'		UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAComboBox("Tipo de sustento:").Select "Boleta de pago - Fijo"
'		
'		'Seleccionamos la opción "Ingreso neto último mes"
'		UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAComboBox("Sustento:").UIAButton("Abrir").Click
'		UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAComboBox("Sustento:").Select "Ingreso neto último mes"
'		
'		'Ingresamos el Valor de sustento
'		UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAEdit("Valor de sustento:").SetValue "5000"
'		
'		'Seleccionamos la opción "Soles"
'		UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAComboBox("Moneda :").UIAButton("Abrir").Click
'		UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAComboBox("Moneda :").Select "Soles"
'		
'		'Damos clic en el boton "Calcular"
'		UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAButton("Calcular").Click
'		wait 10
'		
'		'Control de Mensaje de Error	
'		
'		If JavaDialog("Error Message").Exist(5) Then
'			var1 = JavaDialog("Error Message").JavaObject("JPanel").GetROProperty("text")
'			JavaDialog("Error Message").JavaButton("Cancelar").Click   
'			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
'			DataTable("s_Detalle", dtLocalSheet) = var1
'			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
'			'wait 5
'			'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
'			'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
'			wait 5
'			ExitActionIteration
'		End If	
'		
'		'Bucle que controla que Calcule existosamente
'		tiempo= 0
'		Do
'			wait 1
'			tiempo= tiempo+1
'			If (tiempo >= 180) Then
'				DataTable("s_Resultado",dtLocalSheet) = "Fallido"
'				DataTable("s_Detalle",dtLocalSheet) = "Cálculo Incorrecto"
'				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
'			Else
'				Reporter.ReportEvent micPass, "Exito", "Cálculo correcto"
'			End If
'		Loop While Not (UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar").UIAObject("Score").Exist(1))
'		
'		'Damos clic en el boton "Continuar"
'		UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar").UIAButton("Continuar_2").Click
'		wait 4
'		'Damos clic en el boton "Continuar"
'		UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAButton("Continuar").Click
'		wait 4
'	End If
'	
End Sub
Sub ActualizarAtributos()
	
		tiempo = 0
		Do
			wait 1
			tiempo = tiempo + 1 
			If (tiempo >= 180) Then
			    DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			    DataTable("s_Detalle", dtLocalSheet) =  "Error al cargar la pantalla, no cargó la pantalla Detalles del producto"
			    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorActualizarAtributos.png", True
				imagenToWord "No cargó la pantalla -Actualizar Atributos- de manera correcta.png",RutaEvidencias() &Num_Iter&"_"&"ErrorActualizarAtributos.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			Else
				Reporter.ReportEvent micPass, "Éxito", "Cargó correctamente la pantalla Detalles del producto"
			End If
		Loop While Not (JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").Exist(1) Or JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Exist(2))
		wait 2
		Dim Iterator
	Count = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").GetROProperty ("items count")
	For Iterator = 1 To Count-1
	 	rs = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").GetItem (Iterator)
		If rs = str_TipodeCambio Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Select str_TipodeCambio
			Exit for
		ElseIf Iterator = Count-1 Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Select "Pedido de Cliente"
			Exit for
		End if	
	Next
	
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Texto del motivo:").SetFocus
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Texto del motivo:").Set str_TxtMotivo
		wait 2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActualizarAtributos.png", True
		imagenToWord "Actualizar Atributos.png",RutaEvidencias() &Num_Iter&"_"&"ActualizarAtributos.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Siguiente >").Click
		wait 5
	
End Sub
Sub ParametrizaUsuario()
	
		t=0
		Do
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Parametriza Usuario- de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorParametrizaUsuario.png", True
				imagenToWord "No cargó la pantalla -Parametriza Usuario- de manera correcta.png",RutaEvidencias() &Num_Iter&"_"&"ErrorParametrizaUsuario.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
			wait 1
		'Loop While Not (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaButton("Siguiente >").Exist(1) Or JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").Exist(1))
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Exist) Or(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").Exist))
		
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Exist Then
			wait 2	
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Set
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaButton("Siguiente >").Click
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"ParametrizaProducto.png", True
			imagenToWord "Parametriza el Producto del Usuario",RutaEvidencias() &Num_Iter&"ParametrizaProducto.png"
			wait 2
		End If
End Sub
Sub NegociarConfiguracion()
	
		tiempo = 0
		Do
			wait 1
			tiempo = tiempo + 1
			If (tiempo >= 180) Then
				DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)="No cargó la pantalla -Negociar Configuración del Producto Móvil- de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorNegociarConfiguracion.png", True
				imagenToWord "No cargó la pantalla -Negociar Configuración del Producto Móvil- de manera correcta.png",RutaEvidencias() &Num_Iter&"_"&"ErrorNegociarConfiguracion.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").Close
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
				wait 2
				ExitActionIteration
			Else
				Reporter.ReportEvent micPass, "Éxito", "Cargó correctamente la pantalla Negociar Configuración del Producto Móvil"
			End If
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").Exist(1)
		
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"NegociarConfiguracion.png", True
	imagenToWord "Negociar Configuración del Producto Móvil.png",RutaEvidencias() &Num_Iter&"_"&"NegociarConfiguracion.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Click
	
End Sub
Sub NegociarDistribucion()
	
		tiempo = 0
		Do
			wait 1
			tiempo = tiempo + 1
			If (tiempo>=180) Then
				DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)="No cargó la ventana -Negociar Distribución- o -Resumen de la Orden-"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorNegociarDistribucionOResumenOrden.png", True
				imagenToWord "No cargó la ventana -Negociar Distribución- o -Resumen de la Orden-.png",RutaEvidencias() &Num_Iter&"_"&"ErrorNegociarDistribucionOResumenOrden.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
			Else
				Reporter.ReportEvent micPass, "Éxito", "La ventana 'Negociar Distribución' o 'Resumen de la orden' cargó correctamente"
			End If
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaRadioButton("Nuevo").Exist) or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").Exist))
		
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaRadioButton("Nuevo").Exist Then
		wait 4
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaRadioButton("Nuevo").Set "ON"
		wait 2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"NegociarDistribucion.png", True
		imagenToWord "-Negociar Distribución-.png",RutaEvidencias() &Num_Iter&"_"&"NegociarDistribucion.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Siguiente >").Click
		wait 2
	End If
	
End Sub
Sub ResumenOrden()

		tiempo = 0
		Do
			wait 1	
			tiempo = tiempo + 1
			If (tiempo>=180) Then
				DataTable("s_Resultado",dtLocalSheet) = "Fallido"
				DataTable("s_Detalle",dtLocalSheet) = "No cargó la ventana -Resumen de la orden-"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorResumenOrden.png", True
				imagenToWord "No cargó la ventana -Resumen de la orden-.png",RutaEvidencias() &Num_Iter&"_"&"ErrorResumenOrden.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
			Else
				Reporter.ReportEvent micPass, "Éxito", "La ventana 'Resumen de la Orden' cargó correctamente"
			End If
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").Exist(4)
		
	
'	cantFilas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").GetROProperty ("rows")
'	For i = 2 To cantFilas -2 Step 1
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").SelectRow "#"&i
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").PressKey "C",micCtrl
'		wait 1
'		JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").PressKey "V",micCtrl
'		wait 1
'		valor = JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").GetROProperty ("text")
'		flag = InStr(valor, "Rehabilitar")
'		
'		If flag = 0 Then
'			Reporter.ReportEvent micFail, "Estados", "Uno de los elementos no quedo en estado Removido "
'		End If
'		JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").Set ""
'	Next
'	wait 4
	varhab=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validar y Ver Contrato").GetROProperty("enabled")
	wait 1
	If varhab<>"0" Then

		tiempo = 0
		Do
			wait 1
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").Exist(1) Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validar y Ver Contrato").Click
			End If
			wait 4
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(2) Then
				var1= JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("attached text")
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Mensaje.png", True
				imagenToWord "Mensaje.png",RutaEvidencias() &Num_Iter&"_"&"Mensaje.png"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
			End If
			wait 4
			
			If (tiempo >= 180) Then
				DataTable("s_Detalle",dtLocalSheet) = "Fallido"
				DataTable("s_Resultado",dtLocalSheet) = "Error de Contrato, no se a cargado el contrato correctamente"
				Reporter.ReportEvent micFail, DataTable("s_Detalle",dtLocalSheet), DataTable("s_Resultado",dtLocalSheet)
				ExitActionIteration
			Else
				Reporter.ReportEvent micPass,"Contrato Exitoso","Se ha cargado el contrato correctamente"
			End If
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist) Or (var1="0") Or (var1="Contratos no Generados"))
		wait 3
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Close
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").Set "ON"
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Contrato.png", True
		imagenToWord "Contrato.png",RutaEvidencias() &Num_Iter&"_"&"Contrato.png"
		wait 3
	End If
	End  If
	
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
	wait 3
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Aceptar").Click
	End If
	wait 3
	
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
'	DataTable("s_Resultado", dtLocalSheet) = "Éxito"
'	DataTable("s_Detalle", dtLocalSheet) = "Se envió la orden "&str_NroOrden&" correctamente"
'

	Dim text
    text = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").GetROProperty("text")
    DataTable("s_Nro_Orden", dtLocalSheet) = RTRIM(LTRIM((replace(text,"Orden",""))))
   	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden Generada"&".png", True
	imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Orden Generada"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2386368A").JavaButton("Cerrar").Click
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
	End If
	
End Sub
Sub EmpujaOrden()
	
	If str_TipoData = "DATA LOGICA" Then
	
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select @@ hightlight id_;_10269309_;_script infofile_;_ZIP::ssf3.xml_;_
		wait 5
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Select "Tareas pendientes del equipo" @@ hightlight id_;_9869075_;_script infofile_;_ZIP::ssf5.xml_;_
		wait 5
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").SetFocus
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden",dtLocalSheet) 
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
	
			t = 0
			While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("1 Registros").GetROProperty("attached text") <> "-- Registros") = False
				wait 1
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtLocalSheet) = "Error Data"
					DataTable("s_Detalle", dtLocalSheet) = "Grupo de Órdenes no encuentra datos: "&DataTable("s_Nro_Orden", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorOrdenBuscada_"&Num_Iter&".png", True
					imagenToWord "Grupo de Órdenes no encuentra datos: "&DataTable("s_Nro_Orden", dtLocalSheet), RutaEvidencias() & "ErrorOrdenBuscada_"&Num_Iter&".png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					ExitTestIteration
				End If
			Wend
			wait 1
	
		var = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("1 Registros").GetROProperty("attached text")
		Select Case var
			Case "1 Registros"
				wait 3
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "OrdenEncontrada.png", True
				imagenToWord "Orden Encontrada",RutaEvidencias() & "OrdenEncontrada.png"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").SelectRow "#0" @@ hightlight id_;_13657528_;_script infofile_;_ZIP::ssf6.xml_;_
			Case "0 Registros"
				wait 3
				DataTable("s_Resultado",dtLocalSheet) = "Fallido" 
				DataTable("s_Detalle", dtLocalSheet) = "Grupo de ordenes no encuentra datos: "&str_NroOrden
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Orden0Registros.png", True
				imagenToWord "Orden 0 Registros",RutaEvidencias() & "Orden0Registros.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			Case "-- Registros"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Orden--Registros.png", True
				imagenToWord "Orden -- Registros",RutaEvidencias() & "Orden--Registros.png"
				Reporter.ReportEvent micFail, "Error Data: "&str_NroOrden, "Grupo de ordenes no encuentra datos"
				ExitActionIteration
		End Select
		
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Gestión manual").Click
	
			t = 0
			While (JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaList("Estado de la gestión manual:").Exist) = False
				wait 1
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No cargó el Pop Up para realizar la Gestión Manual"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorGestionManual.png", True
					imagenToWord "No cargó el Pop Up para realizar la Gestión Manual",RutaEvidencias() & "ErrorGestionManual.png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
				End If
			Wend
			wait 1
	
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaList("Estado de la gestión manual:").Select "Cumplimiento Completo"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaList("Motivo de la gestión manual").Select "Manejo manual: Manejo Manual OSS"
		wait 2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "EstadoGestionManual.png", True
		imagenToWord "Estado Gestion Manual",RutaEvidencias() & "EstadoGestionManual.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Enviar").Click
		wait 3
		
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(2) Then
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "EmpujeOK.png", True
			imagenToWord "Empuje OK",RutaEvidencias() & "EmpujeOK.png"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
		End If
		wait 3
	
	End If
	
End Sub
Sub ValidaOrden()

	wait 5
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
	wait 3
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden",dtLocalSheet)
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaCheckBox("Solo órdenes pendientes").Set "OFF"
	wait 1
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
	
		t = 0
		Do
			wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó ninguna orden en la búsqueda"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorBusquedaOrden.png", True
				imagenToWord "No cargó ninguna orden en la búsqueda",RutaEvidencias() & "ErrorBusquedaOrden.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitTestIteration
			End If
		Loop While Not (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("1 Registros").GetROProperty("attached text") = "1 Registros")
		wait 1
	
	varbusqord = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("1 Registros").GetROProperty("attached text")
	Select Case varbusqord 
		Case "1 Registros"
			wait 10
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:")
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "OrdenEncontrada.png", True
			imagenToWord "Orden Encontrada",RutaEvidencias() & "OrdenEncontrada.png"
		    Reporter.ReportEvent micPass,"Existe la Orden", DataTable("s_Nro_Orden", dtLocalSheet)
			wait 2
		Case "0 Registros"
			wait 2
			DataTable("s_Resultado",dtLocalSheet) = "Fallido"
			DataTable("s_Detalle",dtLocalSheet) = "No se encuentra el Nro.Orden: "&str_NroOrden
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "OrdenNoEncontrada.png", True
			imagenToWord "Orden No Encontrada",RutaEvidencias() & "OrdenNoEncontrada.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
			ExitActionIteration
	End Select
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:")
	
		tiempo = 0
		Do
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
			wait 3
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:")
			
			tiempo = tiempo + 1
			If (tiempo >= 180) Then		
				DataTable("s_Resultado",dtLocalSheet) = "Fallido"
				DataTable("s_Detalle",dtLocalSheet) = "La orden "&DataTable("s_Nro_Orden",dtLocalSheet)&" no culminó en estado Cerrado"
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
		DataTable("s_Detalle",dtLocalSheet) = "Se realizó la reinstalación correctamente"
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "OrdenCerrada"&Num_Iter&".png", True
		imagenToWord "Orden Cerrada",RutaEvidencias() & "OrdenCerrada.png"
		Reporter.ReportEvent micPass,"Orden Finalizada","La orden finalizó correctamente"
		wait 5
End Sub
Sub DetalleActividadOrden()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaList("Ver por:").Select "Acciones de orden"
	wait 3
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").DoubleClickCell 0, "#8", "LEFT"
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{ENTER}"
	wait 1
	
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 767572A").JavaEdit("Fecha de vencimiento:").Exist)=False
		wait 1
	Wend

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 767572A").JavaTab("Nombre del cliente:").Select "Actividad"
	
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 767572A").JavaTable("SearchJTable").Exist)=False
		wait 1	
	Wend
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_1"&".png", True
	imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_1"&".png"
	
	shell.SendKeys "{PGDN}"
	wait 1
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_2"&".png", True
	imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_2"&".png"
	
	filas=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 767572A").JavaTable("SearchJTable").GetROProperty("rows")
	For Iterator = 0 To filas-1
		varselec=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 767572A").JavaTable("SearchJTable").GetCellData(Iterator,0)
	Next
	
	If varselec<>"Cerrar Acción de Orden" Then
	 	DataTable("s_Resultado",dtLocalSheet)="Fallido"
		DataTable("s_Detalle",dtLocalSheet)="La orden "&DataTable("s_Nro_Orden",dtLocalSheet)&" culmino en la Actividad "&varselec&""
		Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 767572A").JavaButton("Cancelar").Click

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
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 767572A").JavaButton("Cancelar").Click

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
		wait 2
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
	End If
	
End Sub

