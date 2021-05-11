
'Flag CAEQ_COMBO e_MotivoCambio="CAEQ_EQUIPO Y SIM"
'Flag CAEQ_SIM   e_MotivoCambio="CAEQ_SIM"
'Flag CAEQ_Equipo   e_MotivoCambio="CAEQ_EQUIPO"

Dim var1, var2, str_var3, str_var4, var4, varlog, filas, varhab, varValidaRespuestaCumplimiento, var6, varlog3, var8, nroreg
Dim str_TipodeCambio
Dim str_MotivoCambio
Dim str_EquipoMovil
Dim str_idDispositivo
Dim str_Tipo_Data_Eqp
Dim vardisp
Dim str_tipoSIM
Dim str_idSim
Dim str_Ambiente
Dim str_numeroID
Dim str_ValPenalidad
Dim str_MedioPago
Dim str_Cuotas
Dim str_Financiamiento
Dim str_CambioPlan 
Dim str_Plan

Num_Iter 	   		=   Environment.Value("ActionIteration") 
str_IDServicio     	= 	DataTable("e_ID_Servicio", dtLocalSheet)
str_Ambiente        =   DataTable("e_ambiente",dtLocalSheet)
str_TipodeCambio    =   DataTable("e_TipodeCambio", dtLocalSheet)
str_MotivoCambio    =   DataTable("e_MotivoCambio", dtLocalSheet)
str_Financiamiento	=	DataTable("e_Financiamiento", dtLocalSheet)
str_EquipoMovil     = 	DataTable("e_Equipo_Movil", dtLocalSheet)
str_idDispositivo  	= 	DataTable("e_ID_Dispositivo", dtLocalSheet)
str_idSim			= 	DataTable("e_SerieSIM", dtLocalSheet)
str_Tipo_Data 		= 	DataTable("e_Tipo_Data", dtLocalSheet)
str_tipoSIM         =   DataTable("e_tipoSIM",dtLocalSheet)
str_ValPenalidad 	= 	DataTable("e_ValidaPenalidad", dtLocalSheet)
str_MedioPago		=  	DataTable("e_MedioPago", dtLocalSheet)
str_Cuotas			=  	DataTable("e_Cuotas", dtLocalSheet)
str_NroOrden        = 	DataTable("s_Nro_Orden", dtLocalSheet)
str_CambioPlan      =   DataTable("e_CambioPlan", dtLocalSheet)
str_Plan            =   DataTable("e_Plan", dtLocalSheet)
str_metodo_entrega  =   DataTable("Met_Entrega", dtLocalSheet)
str_Financiamiento  =   ucase(str_Financiamiento)
varlog3				=  "<html>Cambiar</html>"

Call PanelInteraccion()
Call IngresoNumero()


Call esperaCarga
'Call verificacionSolicitud
'Call ingresoNum
Call detallesProductos
Call DetalledelaLinea
Call ParametrosCambio
Call RecursosCambio
Call TipoEnvio
'Call reemplazarCargo
Call GeneracionOrden
Call EnviarOrden()
If str_metodo_entrega <> "Delivery" Then
	Call GestionLogistica
   Call EmpujeOrden
   Call OrdenCerrado
End If

Call cerrarVentanas
Sub esperaCarga
	wait 6
End Sub
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
Sub detallesProductos
Dim rowOrdenCnt, strNombre
	t = 0
	While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaTable("Antigüedad de línea:").Exist) = False)
		wait 1	
		
		t = t + 1
		If (t >= 180) Then
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Detalles_Producto_"&Num_Iter&".png", True
			imagenToWord "Error_Detalles_Producto_"&Num_Iter,RutaEvidencias() & "Detalles_Producto_"&Num_Iter&".png"
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No cargó el control -Detalles del Producto Asignado- de manera correcta"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
	Wend

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaTab("Antigüedad de línea:").Select "Configuración"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaTab("Antigüedad de línea:").Type micRight
	strNombre=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaTab("Antigüedad de línea:").GetROProperty("value")
	If strNombre="Conexiones" or strNombre="Conexiones [Ninguna]" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaTab("Antigüedad de línea:").Type micRight
	Else 
		If strNombre="Órdenes pendientes [Ninguna]" Then
			rowOrdenCnt=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaTable("Antigüedad de línea:").GetROProperty("rows")
		Else 
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Cant_Ordenes_"&Num_Iter&".png", True
			imagenToWord "Error_Cant_Ordenes_"&Num_Iter,RutaEvidencias() & "Cant_Ordenes_"&Num_Iter&".png"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaTab("Antigüedad de línea:").Select "Órdenes pendientes"
			rowOrdenCnt=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaTable("Antigüedad de línea:").GetROProperty("rows")
			
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "El número a buscar tiene "&rowOrdenCnt&" órdenes pendientes."
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
	End If	
	
	If rowOrdenCnt > 0 Then
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		DataTable("s_Detalle", dtLocalSheet) = "El número "&str_IDServicio&" posee Órdenes pendientes"
		Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "OrdenPend.png", True
		imagenToWord "El Numero posee Orden Pendiente",RutaEvidencias() & "OrdenPend.png"
		ExitActionIteration
	Else 
		DataTable("s_Resultado", dtLocalSheet) = "Exitoso"
		DataTable("s_Detalle", dtLocalSheet) = "El número "&str_IDServicio&" no posee Ordenes pendientes"
		Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "SinOrdenPend.png", True
		imagenToWord "El Numero no posee Orden Pendiente",RutaEvidencias() & "SinOrdenPend.png"
	End If
	
	
	If (DataTable("e_CalcPenalidad", dtLocalSheet) = "SI") Then	
		Call CalcularPenalidad
	End If

End Sub
Sub DetalledelaLinea()
	JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").JavaMenu("Cambiar").Select
	Call esperaCarga
	
	If DataTable("e_WIC_ValidaCli",dtLocalSheet)="SI" Then
	

RunAction "WIC", oneIteration

	End If
	t = 0
	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de_2").JavaList("Motivo:").Exist) = False
		wait 1
		t = t + 1
		If (t >= 180) Then
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Actualizar_Accion_Orden_"&Num_Iter&".png", True
			imagenToWord "Error_Actualizar_Accion_Orden_"&Num_Iter,RutaEvidencias() & "Actualizar_Accion_Orden_"&Num_Iter&".png"
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Actualizar atributos de Acción de Orden- de manera correcta"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
			ExitActionIteration
		End If
	Wend
End Sub
Sub ParametrosCambio()
	
	'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de_2").JavaList("Motivo:").Select str_TipodeCambio
	Dim Iterator
	Count = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de_2").JavaList("Motivo:").GetROProperty ("items count")
	
	For Iterator = 1 To Count-1
	 	rs = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de_2").JavaList("Motivo:").GetItem (Iterator)
		If rs = str_TipodeCambio Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de_2").JavaList("Motivo:").Select str_TipodeCambio
			Exit for
		ElseIf Iterator = Count-1 Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de_2").JavaList("Motivo:").Select "Pedido de Cliente"
			Exit for
		End if	
	Next
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de_2").JavaEdit("Texto del motivo:").Set str_MotivoCambio @@ hightlight id_;_1371164_;_script infofile_;_ZIP::ssf14.xml_;_
	wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "PantallaActualizarAtributos.png", True
	imagenToWord "Pantalla Actualizar Atributos",RutaEvidencias() & "PantallaActualizarAtributos.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Siguiente >").Click @@ hightlight id_;_15762962_;_script infofile_;_ZIP::ssf17.xml_;_
	
	
	Dim i
	i=0
	While not i=180
		i=i+1
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Exist(1) Then
			i=180
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Set
			
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaButton("Siguiente >").Click
			t=0
			While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Buscar").Exist) = False
				wait 1
				t = t + 1
				If (t >= 180) Then
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Negociar_Configuracion_"&Num_Iter&".png", True
					imagenToWord "Error_Negociar_Configuracion_"&Num_Iter,RutaEvidencias() & "Negociar_Configuracion_"&Num_Iter&".png"
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Negociar Configuración- de manera correcta"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
					ExitActionIteration
				End If
			Wend		
		ElseIf JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Buscar").Exist(1) Then
			i=180
			wait 1
		Else 
			If i=180 Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Parametriza_Producto_"&Num_Iter&".png", True
				imagenToWord "Error_Parametriza_Producto_"&Num_Iter,RutaEvidencias() & "Parametriza_Producto_"&Num_Iter&".png"
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Parametriza el Producto- de manera correcta"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaButton("Cerrar").Click
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
				wait 2
				ExitActionIteration
			End If
		End If
		
	Wend
		
		
End Sub
Sub IngresodeNumero()
	wait 3
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Productos Asignados").Select @@ hightlight id_;_28643574_;_script infofile_;_ZIP::ssf1.xml_;_
	t=0
	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaEdit("TextFieldNative$1_2").Exist) = False
		wait 1	
		t = t + 1
		If (t >= 180) Then
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Productos_Asignados_"&Num_Iter&".png", True
			imagenToWord "Error_Productos_Asignados_"&Num_Iter,RutaEvidencias() & "Productos_Asignados_"&Num_Iter&".png"
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Productos asignados- de manera correcta"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
			ExitActionIteration
		End If
	Wend

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaEdit("TextFieldNative$1_2").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaEdit("TextFieldNative$1_2").Set str_IDServicio
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Click @@ hightlight id_;_31932278_;_script infofile_;_ZIP::ssf4.xml_;_
	wait 2

		tiempo=0
		Do 
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Click
				nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("1 Registros").GetROProperty("attached text")
				tiempo=tiempo+1
				wait 1
			End If
			If (tiempo >= 180) Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Numero_Registro_"&Num_Iter&".png", True
				imagenToWord "Error_Numero_Registro_"&Num_Iter,RutaEvidencias() & "Numero_Registro_"&Num_Iter&".png"
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "Tiene muchos Registros que se detuvo el flujo."
				Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
				ExitActionIteration
			End If
		Loop While Not(nroreg="1 Registros")
		wait 1

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaTable("Ver por:").DoubleClickCell 0, "#2", "LEFT"
	wait 2	

	t = 0
	While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Calcular Penalidad").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").Exist)) = False
		wait 1	
		
		t = t + 1
		If (t >= 180) Then
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorBotónCalcularPenalidad_"&Num_Iter&".png", True
			imagenToWord "Error Botón Calcular Penalidad_"&Num_Iter,RutaEvidencias() & "ErrorBotónCalularPenalidad_"&Num_Iter&".png"
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No cargó el control -Detalles del Producto Asignado- de manera correcta"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
	Wend
	Wait 1
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").Exist(3) Then
		varlog2 = JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").JavaTable("Las siguientes acciones").GetCellData(0,0)
	  	DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		DataTable("s_Detalle", dtLocalSheet) = varlog2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "MensajeError.png", True
		imagenToWord "Mensaje Error",RutaEvidencias() & "MensajeError.png"
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").JavaButton("Rechazar solicitud de").Click
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
		wait 2
		ExitActionIteration
	End If
End Sub
Sub Carga()
	While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
		wait 1
	Wend
End Sub
Sub EliminarEquipoSIM()

	dim Iterator , filas, b
	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetROProperty("rows")
	For Iterator = 0 To filas-1
	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow ("#"&Iterator)
		j = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetCellData("#"&Iterator, "#1")
		j= Replace(j, "<html>","") 
		j = Replace(j, "</html>","") 
		j = left(j,4)
		If j = "TSPE" Then
		   Iterator = Iterator-2
		   JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow ("#"&Iterator)
		   wait 1
		   b = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Eliminar").GetROProperty("enabled")
		   If b="0" Then
			   	Iterator = Iterator-1
			   	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow ("#"&Iterator)
			   	b = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Eliminar").GetROProperty("enabled")
			   	If b="1" Then
			   		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "EliminarDisp.png", True
					imagenToWord "Se elimina dispositivo",RutaEvidencias() & "EliminarDisp.png"
			   	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Eliminar").Click
			   	    Call Carga()
					wait 2
			   	End If
		   	ElseIf b="1" Then
		   	    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "EliminarDisp.png", True
				imagenToWord "Se elimina dispositivo",RutaEvidencias() & "EliminarDisp.png"
		   	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Eliminar").Click
		   	    Call Carga()
		   	   wait 2
		   End If
		  	
		  	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow ("#"&Iterator)
		  	b = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Eliminar").GetROProperty("enabled")
		   If b = "1" Then
		   	   JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "EliminarTarjSIM.png", True
				imagenToWord "Se elimina TARJETA SIM",RutaEvidencias() & "EliminarTarjSIM.png"
		   	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Eliminar").Click
		   	    Call Carga()
				wait 2
					   	    
		   ElseIf b = "0" Then
		   		Iterator = Iterator+1
		   		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow ("#"&Iterator)
		   		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "EliminarTarjSIM.png", True
				imagenToWord "Se elimina TARJETA SIM",RutaEvidencias() & "EliminarTarjSIM.png"
		   	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Eliminar").Click
		   	   Call Carga()
		   		wait 2
		   End If
			Exit for
		End If
	Next

End Sub
Sub EliminarSIM()
	dim Iterator , filas, b
	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetROProperty("rows")
	For Iterator = 0 To filas-1
	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow ("#"&Iterator)
		j = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetCellData("#"&Iterator, "#1")
		j= Replace(j, "<html>","") 
		j = Replace(j, "</html>","") 
		j = left(j,4)
		If j = "TSPE" Then
		   Iterator = Iterator-1
		   JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow ("#"&Iterator)
		   b = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Eliminar").GetROProperty("enabled")
		   If b = "1" Then
		   	   JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "EliminarTarjSIM.png", True
				imagenToWord "Se elimina TARJETA SIM",RutaEvidencias() & "EliminarTarjSIM.png"
		   	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Eliminar").Click
		   	    Call Carga()
				wait 2
		   End If
			Exit for
		End If
	Next
End Sub
Sub EliminarEquipo()
			
	dim Iterator , filas, b
	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetROProperty("rows")
	For Iterator = 0 To filas-1
	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow ("#"&Iterator)
		j = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetCellData("#"&Iterator, "#1")
		j= Replace(j, "<html>","") 
		j = Replace(j, "</html>","") 
		j = left(j,4)
		If j = "TSPE" Then
		   Iterator = Iterator-2
		   JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow ("#"&Iterator)
		   wait 1
		   b = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Eliminar").GetROProperty("enabled")
		   If b="0" Then
			   	Iterator = Iterator-1
			   	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow ("#"&Iterator)
			   	b = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Eliminar").GetROProperty("enabled")
			   	If b="1" Then
			   		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "EliminarDisp.png", True
					imagenToWord "Se elimina dispositivo",RutaEvidencias() & "EliminarDisp.png"
			   	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Eliminar").Click
			   	    Call Carga()
					wait 2
			   	End If
		   	ElseIf b="1" Then
		   	    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "EliminarDisp.png", True
				imagenToWord "Se elimina dispositivo",RutaEvidencias() & "EliminarDisp.png"
		   	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Eliminar").Click
		   	    Call Carga()
		   	   wait 2
		   End If
			Exit for
		End If
	Next
	End Sub	
Sub IngresoGrupoSIM()
	dim Iterator , filas, j,h
	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetROProperty("rows")
	For Iterator = filas-1 To 0 step -1
	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow ("#"&Iterator)
		j = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetCellData("#"&Iterator, "#1")
	    h = Instr(1,j,"Grupo de SIM")
	    If h <> 0 Then
	    	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell "#"&Iterator, "#1", "LEFT" 
	    	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData "#"&Iterator, "#1","Estandar"
	    	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "IngresoGrupoSIM.png", True
			imagenToWord "Ingresamos Grupo SIM",RutaEvidencias() & "IngresoGrupoSIM.png"
	    	Exit for 
	    End If
	Next 
End Sub	
Sub IngresoTipoSIM()
	
	dim Iterator , filas, j,h
	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetROProperty("rows")
	For Iterator = filas-1 To 0 step -1
	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow ("#"&Iterator)
		j = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetCellData("#"&Iterator, "#1")
	    h = Instr(1,j,"Tipo de SIM")
	    If h <> 0 Then
	        JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell "#"&Iterator, "#1", "LEFT" 
	    	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData "#"&Iterator, "#1",str_tipoSIM
	    	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "IngresoTipoSIM.png", True
			imagenToWord "Ingresamos Tipo SIM",RutaEvidencias() & "IngresoTipoSIM.png"
	   		
	    	Exit for 
	    End If
	Next 
End Sub	
Sub IngresoIMEI()
	dim Iterator , filas, j,h
	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetROProperty("rows")
	For Iterator = filas-1 To 0 step -1
	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow ("#"&Iterator)
		j = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetCellData("#"&Iterator, "#1")
	    h = Instr(1,j,"Número IMEI")
	    If h <> 0 Then
	    	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell "#"&Iterator, "#1", "LEFT" 
	    	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData "#"&Iterator, "#1",str_idDispositivo
	    	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "IngresoIMEI.png", True
			imagenToWord "Ingresamos IMEI",RutaEvidencias() & "IngresoIMEI.png"
	    	Exit for 
	    End If
	Next 
End Sub	
Sub InsertarDispositivo()
	dim Iterator , filas, j,h
	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:_2").GetROProperty("rows")
	For Iterator = filas-1 To 0 step -1
	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:_2").SelectRow ("#"&Iterator)
		j = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:_2").GetCellData("#"&Iterator, "#1")
	    h = Instr(1,j,"ID del Equipo")
	    If h <> 0 Then
	    	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:_2").DoubleClickCell "#"&Iterator, "#2", "LEFT"
	    	Exit for 
	    End If
	Next 
End Sub	
Sub RecursosCambio()
	If str_MotivoCambio="CAEQ_EQUIPO Y SIM" Then
		Call EliminarEquipoSIM()
	ElseIf str_MotivoCambio="CAEQ_SIM" Then
		Call EliminarSIM()
	ElseIf str_MotivoCambio="CAEQ_EQUIPO" Then
		Call EliminarEquipo()
	End If
	
	If str_MotivoCambio="CAEQ_EQUIPO Y SIM" Then
		wait 2
		'Agrega SIM
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Expand "#0;Móvil;SIM y Dispositivo"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;SIM y Dispositivo;Tarjeta SIM"
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "SeleccionOpcionTarjetaSIM.png", True
		imagenToWord "Selección Opción Tarjeta SIM",RutaEvidencias() & "SeleccionOpcionTarjetaSIM.png"
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").GetROProperty("enabled") = "1") = False
			wait 1
			t = t + 1
			If (t >= 180) Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorBotónAgregar_"&Num_Iter&".png", True
				imagenToWord "Error Botón Agregar_"&Num_Iter,RutaEvidencias() & "ErrorBotónAgregar_"&Num_Iter&".png"
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No se deshabilitó botón -Agregar-"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
		Call Carga()
		wait 2
		''Se Agrega Dispositivo
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;SIM y Dispositivo;Dispositivo"
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "SeleccionOpcionEquipo.png", True
		imagenToWord "Selección Opción Dispositivo",RutaEvidencias() & "SeleccionOpcionEquipo.png"
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").GetROProperty("enabled") = "1") = False
			wait 1
			t = t + 1
			If (t >= 180) Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorBotónEliminar_"&Num_Iter&".png", True
				imagenToWord "Error Botón Eliminar_"&Num_Iter,RutaEvidencias() & "ErrorBotónEliminar_"&Num_Iter&".png"
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No se deshabilitó botón -Agregar-"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
		Call Carga()
		wait 2

	ElseIf str_MotivoCambio="CAEQ_SIM" Then
		'Se Agrega SIM
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Expand "#0;Móvil;SIM y Dispositivo"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;SIM y Dispositivo;Tarjeta SIM"
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "SeleccionOpcionTarjetaSIM.png", True
		imagenToWord "Selección Opción Tarjeta SIM",RutaEvidencias() & "SeleccionOpcionTarjetaSIM.png"
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").GetROProperty("enabled") = "1") = False
			wait 1
			t = t + 1
			If (t >= 180) Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorBotónAgregar_"&Num_Iter&".png", True
				imagenToWord "Error Botón Agregar_"&Num_Iter,RutaEvidencias() & "ErrorBotónAgregar_"&Num_Iter&".png"
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No se deshabilitó botón -Agregar-"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
		Call Carga()
		wait 2
	ElseIf str_MotivoCambio="CAEQ_EQUIPO" Then
		''Agrega IMEI
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Expand "#0;Móvil;SIM y Dispositivo"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;SIM y Dispositivo;Dispositivo"
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "SeleccionOpcionEquipo.png", True
		imagenToWord "Selección Opción Dispositivo",RutaEvidencias() & "SeleccionOpcionEquipo.png"
		
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").GetROProperty("enabled") = "1") = False
			wait 1
			t = t + 1
			If (t >= 180) Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorBotónEliminar_"&Num_Iter&".png", True
				imagenToWord "Error Botón Eliminar_"&Num_Iter,RutaEvidencias() & "ErrorBotónEliminar_"&Num_Iter&".png"
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No se deshabilitó botón -Agregar-"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
		Call Carga()
		wait 2
	End If
	
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
	Call Carga()
	wait 2		
	varlog2 = "<html>Falta el atributo obligatorio Número IMEI de Dispositivo. Ingres"
	If (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist) Then
		If (Left((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable_2").GetCellData(0,1)),70) = varlog2) Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		End If
	End If
	
	If str_MotivoCambio="CAEQ_EQUIPO Y SIM" Then
	
''''''''Seleccion Opciones de SIM
		wait 2
		Call IngresoTipoSIM()
		wait 2
        call IngresoGrupoSIM()
        wait 2
        call IngresoIMEI()
        
    ElseIf str_MotivoCambio="CAEQ_EQUIPO"  Then
        wait 2
        call IngresoGrupoSIM()
        wait 2
        call IngresoIMEI()
        
    ElseIf str_MotivoCambio="CAEQ_SIM" Then
        JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "CambioSIM.png", True
		imagenToWord "Cambiamos Tarjeta SIM ",RutaEvidencias() & "CambioSIM.png"
	End If		

	
	If str_MotivoCambio="CAEQ_EQUIPO Y SIM" or str_MotivoCambio="CAEQ_EQUIPO" Then

		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Mostrar atributos:").Select "Todo" @@ hightlight id_;_31325070_;_script infofile_;_ZIP::ssf93.xml_;_
		wait 3
		call InsertarDispositivo()
		
	''Pantalla de Equipo
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaButton("Buscar").Exist) = False
			wait  1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "La pantalla de Filtro y Selección del Equipo no cargó de manera correcta"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorFiltroSelecciónEquipo_"&Num_Iter&".png", True
				imagenToWord "Error Filtro Selección Equipo_"&Num_Iter,RutaEvidencias() & "ErrorFiltroSelecciónEquipo_"&Num_Iter&".png"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaButton("Cancelar").Click	
				wait 4
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
				wait 4
				If JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").Exist(5) Then
					JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Click
					wait 3
				End If
				If JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_3").Exist(5) Then
					JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_3").JavaList("Motivo:").Select "Pedido de Cliente"
					wait 2
					JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_3").JavaButton("Aceptar").Click
					wait 5
				End If
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar_2").Exist(3) Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar_2").Click
					wait 1
				End If
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Exist(3) Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
					wait 1
				End If
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Exist(3) Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
					wait 1
				End If
			End If		
		Wend 
	
	
		''Selección de Equipo
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaObject("PanelNative$JXPanel").Exist Then
			'Bucle que espera la carga de los equipos
			t = 0
			While ((JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaStaticText("Mostrando 1-6 de 20 equipos").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaStaticText("No hay Dispositivos en").Exist)) = False
				wait 1
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "La pantalla de Filtro y Selección del Equipo no cargó de manera correcta"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorFiltroSelecciónEquipo_"&Num_Iter&".png", True
					imagenToWord "Error Filtro Selección Equipo_"&Num_Iter,RutaEvidencias() & "ErrorFiltroSelecciónEquipo_"&Num_Iter&".png"
					JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaButton("Cancelar").Click	
					wait 4
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
					wait 4
					If JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").Exist(5) Then
						JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Click
						wait 3
					End If
					If JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_3").Exist(5) Then
						JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_3").JavaList("Motivo:").Select "Cancelar a Pedido de Cliente"
						wait 2
						JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_3").JavaButton("Aceptar").Click
						wait 5
					End If
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar_2").Exist(3) Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar_2").Click
						wait 1
					End If
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Exist(3) Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
						wait 1
					End If
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Exist(3) Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
						wait 1
					End If
				End If
			Wend
			
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaEdit("TextFieldNative$1").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaEdit("TextFieldNative$1").Set str_EquipoMovil
			End If
			wait 2
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaButton("Buscar").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaButton("Buscar").Click
			End If
			
			If ucase(str_EquipoMovil = "HUAWEI P10 NEGRO") Then
			   
				While ((JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaStaticText("Mostrando 1-6 de 10 equipos(st").Exist) OR (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist))= false
					wait 1
				Wend
				
				If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
					Call avisoError	
				Else 
					If JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaButton("LOB_Close").Exist Then
						JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaButton("LOB_Close").Click
					End If
					wait 2
					JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaCheckBox("Seleccionar_3").Set "ON"
					wait 2
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "SeleccionEquipo.png", True
					imagenToWord "Se selecciona el equipo",RutaEvidencias() & "SeleccionEquipo.png"
				End If
				
			Else 
				While ((JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaCheckBox("Seleccionar").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist)) = false
					wait 1
				Wend
				If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
					Call avisoError	
				Else 
					wait 2
					JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaCheckBox("Seleccionar").Set "ON"
					wait 2
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "SeleccionEquipo.png", True
					imagenToWord "Se selecciona el equipo",RutaEvidencias() & "SeleccionEquipo.png"
				End If
			End If
		
	
'			If (JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist) Then
'				Call avisoError	
'			End If
	
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaButton("Agregar").Click
			
		Else
		
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaEdit("TextFieldNative$1").SetFocus
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaEdit("TextFieldNative$1").Set str_EquipoMovil
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaButton("Buscar").Click
			Wait 2
			t = 0
			While ((JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaButton("Equipo").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaDialog("Mensaje").Exist)) = False
				Wait 1
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "El equipo "&str_Dispositivo&" no ha sido encontrado en lista"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "EquipoNoEncontrado_"&Num_Iter&".png", True
					imagenToWord "Equipo No Encontrado_"&Num_Iter,RutaEvidencias() & "EquipoNoEncontrado_"&Num_Iter&".png"
					JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaButton("Cancelar").Click	
					wait 2
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
					wait 2
					If JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").Exist(5) Then
						JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Click
						wait 5
					End If
					If JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_3").Exist(5) Then
						JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_3").JavaList("Motivo:").Select "Cancelar a Pedido de Cliente"
						wait 2
						JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_3").JavaButton("Aceptar").Click
						wait 5
					End If
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar_2").Exist(3) Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar_2").Click
						wait 2
					End If
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Exist(3) Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
						wait 2
					End If
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Exist(3) Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
						wait 2
					End If
					ExitActionIteration	
				End If			
			Wend
			
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaButton("Equipo").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaCheckBox("Seleccionar").Set "ON"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "EquipoSeleccionado.png", True
				imagenToWord "Equipo Seleccionado",RutaEvidencias() & "EquipoSeleccionado.png"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_2").JavaButton("Agregar").Click
				wait 2
			End If
		End If
	End If
	
	
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Exist = False
		wait 1
	Wend	
	
	wait 5
	
	Call Carga()
	
	If  ucase(str_CambioPlan) = "SI" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Reemplazar oferta").Click
		While JavaWindow("Ejecutivo de interacción").JavaDialog("null Móvil (Orden 930703A").JavaObject("PanelNative$JXPanel").Exist = False
			wait 1
		Wend
		JavaWindow("Ejecutivo de interacción").JavaDialog("null Móvil (Orden 930703A").JavaEdit("TextFieldNative$1").Set str_Plan 
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("null Móvil (Orden 930703A").JavaButton("Buscar").Click
		While JavaWindow("Ejecutivo de interacción").JavaDialog("null Móvil (Orden 930703A").JavaCheckBox("Seleccionar").Exist = false
			wait 1
		Wend
		JavaWindow("Ejecutivo de interacción").JavaDialog("null Móvil (Orden 930703A").JavaCheckBox("Seleccionar").Set "ON"
		wait 2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "SeleccionPlan.png", True
	    imagenToWord "Seleccionamos Plan Móvil",RutaEvidencias() & "SeleccionPlan.png"
	    JavaWindow("Ejecutivo de interacción").JavaDialog("null Móvil (Orden 930703A").JavaButton("Siguiente >").Click
	    While JavaWindow("Ejecutivo de interacción").JavaDialog("null Móvil (Orden 930703A").JavaTable("(Nuevo)").Exist = False
	    	wait 1
	    Wend
	    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "DetallePlan.png", True
	    imagenToWord "Detalle del Plan",RutaEvidencias() & "DetallePlan.png"
	    JavaWindow("Ejecutivo de interacción").JavaDialog("null Móvil (Orden 930703A").JavaButton("Siguiente >").Click
	    Call Carga()
	End If
	

	If DataTable("e_CR3988", dtLocalSheet) = "SI" Then
		wait 1
        
	End If

	
	IF ucase(str_Financiamiento) = "SI"  Then
	     If str_Cuotas = 18 Then
		     JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "MOVISTAR-18 cuotas"
		     Set shell = CreateObject("Wscript.Shell")
		     shell.SendKeys "{RIGHT 100}"
		     JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar_2").Click
		     Call Carga()
		     wait 1
             JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Calcular").Click
             Call Carga()
		     wait 1
		     JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Financiamiento18cuotas.png", True
	         imagenToWord "Financiamiento de 18 cuotas",RutaEvidencias() & "Financiamiento18cuotas.png"
	         
		     ElseIf str_Cuotas = 12 Then
		     JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "MOVISTAR-12 cuotas"
		     Set shell = CreateObject("Wscript.Shell")
		     shell.SendKeys "{RIGHT 100}"
		     JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar_2").Click
		     Call Carga()
		     wait 1
             JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Calcular").Click
             Call Carga()
		     wait 1
		     JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Financiamiento12cuotas.png", True
	         imagenToWord "Financiamiento de 12 cuotas",RutaEvidencias() & "Financiamiento12cuotas.png"
	 
	     End If
	ElseIf ucase(str_Financiamiento) = "NO" and DataTable("e_CR3988", dtLocalSheet) = "NO" Then 
	     JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "Contado"
	     Set shell = CreateObject("Wscript.Shell")
		 shell.SendKeys "{RIGHT 100}"
		 JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar_2").Click
		 Call Carga()
		 wait 1
         JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Calcular").Click
         Call Carga()
		 wait 1
		 JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "FinanciamientoContado.png", True
	     imagenToWord "Financiamiento contado",RutaEvidencias() & "FinanciamientoContado.png"
	End If
       
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "PantallaNegociarConfiguración.png", True
	imagenToWord "Pantalla Negociar Configuración",RutaEvidencias() & "PantallaNegociarConfiguración.png"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Click
	Call esperaCarga
End Sub	
Sub avisoError
	While((JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist)) = False
			wait 1
	Wend
		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").JavaButton("OK").Exist(3) Then
	   	 	JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").JavaButton("OK").Click
	   	 	wait 2
		End  If
End Sub
Sub acuerdoFacturacion
''Validar lo adicional en producción Pelao
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaRadioButton("Única factura").Set "ON"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("ID del Acuerdo de Facturación:").WaitProperty "value", Not Empty, 30000 @@ hightlight id_;_20967626_;_script infofile_;_ZIP::ssf87.xml_;_
	
	If LCase(str_ambiente)="desarrollo" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaRadioButton("Nuevo").Set
	End If @@ hightlight id_;_28483837_;_script infofile_;_ZIP::ssf25.xml_;_
	wait 3
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ValidaciónAcuerdoFacturación.png", True
	imagenToWord "Validación Acuerdo Facturación",RutaEvidencias() & "ValidaciónAcuerdoFacturación.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaButton("Siguiente >").Click
	Call esperaCarga
End Sub
Sub pagoInmediato
			t = 0
		While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaButton("Pago inmediato").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist)) = False
			Wait 1
			t = t + 1
			If (t >= 180) Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorBotónPagoInmediato_"&Num_Iter&".png", True
				imagenToWord "Error Botón Pago Inmediato_"&Num_Iter,RutaEvidencias() & "ErrorBotónPagoInmediato_"&Num_Iter&".png"
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No se habilitó el botón -Pago Inmediato- de la siguiente pantalla"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist Then
			var1 = JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaStaticText("<html>No se puede reactivar").GetROProperty("label")
			var1 = Replace(var1,"<html>","")
			var1 = Replace(var1,"</html>","")
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = var1
			Reporter.ReportEvent micFail, "Error de Resumen de Contrato", "No se ha cargado el resumen correctamente"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "MensajeProblema.png", True
			imagenToWord "Mensaje Problema",RutaEvidencias() & "MensajeProblema.png"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaButton("Cerrar").Click
			Wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
			Wait 2
		    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
		    Wait 2
			ExitActionIteration
		End If
		If str_Financiamiento = "SI" or DataTable("e_CR3988", dtLocalSheet) = "SI"  Then
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaCheckBox("Financiamiento Externo").Set "ON"
'			wait 1
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaCheckBox("Financiamiento Externo").Set "OFF"
'			wait 1
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaCheckBox("Financiamiento Externo").Set "ON"
			t=0
			While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaEdit("Importe de Cuota Mayor:").Exist) = False
				wait 1
				t = t + 1
				If (t >= 180) Then
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorCargaFinanciamiento_"&Num_Iter&".png", True
					imagenToWord "Error Carga Financiamiento_"&Num_Iter,RutaEvidencias() & "ErrorCargaFinanciamiento_"&Num_Iter&".png"
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No cargó la ventana -Pago Inmediato Financimiento- de manera correcta"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
				End If
			Wend
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaEdit("Importe de Cuota Mayor:").Set 1
			wait 2
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "SolicitudFinanciamiento.png", True
			imagenToWord "Solicitud Financiamiento",RutaEvidencias() & "SolicitudFinanciamiento.png"
		End If
		
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaButton("Límite de Compra").Exist(2) Then
			var6 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaButton("Límite de Compra").GetROProperty("enabled")
			If (var6 >= 1) Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaButton("Límite de Compra").Click
				While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaButton("Pago inmediato").GetROProperty("enabled") = 0
					wait 1
				Wend
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaButton("Pago inmediato").Click 
				wait 3
			End If
		End If
	t=0
	While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Exist)  or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaStaticText("'Seleccione la casilla").Exist))= False
		wait 1
		t = t + 1
		If (t >= 180) Then
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorCargaPagoInmediato_"&Num_Iter&".png", True
			imagenToWord "Error Carga Pago Inmediato_"&Num_Iter,RutaEvidencias() & "ErrorCargaPagoInmediato_"&Num_Iter&".png"
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No cargó la ventana -Pago Inmediato- de manera correcta"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
	Wend
	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaStaticText("'Seleccione la casilla").Exist Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "MensajeFinanzaExterna.png", True
		imagenToWord "Mensaje de Finanza Externa",RutaEvidencias() & "MensajeFinanzaExterna.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaButton("Siguiente >").Click
		Exit SUB

	End If
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Select "Factura"
	wait 2
	var8 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Numero RUC").GetROProperty("text")
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Numero RUC").GetROProperty("text") = "" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Select "Boleta"
	End If
	
	wait 1
	Dim Iterator
	Count = 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").GetROProperty ("items count")

	For Iterator = 0 To Count-1
	 	rs = 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").GetItem (Iterator)
	 
		If rs = DataTable("e_MedioPago", dtLocalSheet) Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Select DataTable("e_MedioPago", dtLocalSheet)
				If str_MedioPago = "Pago a la Factura" Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Cantidad de cuotas:").Select str_Cuotas
					wait 1
				    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Calcular").Click
				End If
			Exit for
		ElseIf Iterator = Count-1 Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Select "Externo"
			Exit for
		End if	
	Next
	wait 1
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ProcesarPagoInmediato.png", True
	imagenToWord "Procesar Pago Inmediato",RutaEvidencias() & "ProcesarPagoInmediato.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Enviar").Click
	
	t=0
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaButton("Pago inmediato").Exist) = true
		wait 1
		t = t + 1
		If (t >= 180) Then
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorBotónPagoInmediato2_"&Num_Iter&".png", True
			imagenToWord "Error Carga Botón PagoInmediato Finalizado proceso de Pago_"&Num_Iter,RutaEvidencias() & "ErrorBotónPagoInmediato2_"&Num_Iter&".png"
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No cargó la ventana -Negociar Pago- de manera correcta"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
	Wend
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "NegociarPago.png", True
	imagenToWord "Negociar Pago Inmediato",RutaEvidencias() & "NegociarPago.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaButton("Siguiente >").Click
	
'		t = 0 
'		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("En tienda").Exist) = False
'			Wait 1
'			
'			t = t + 1
'			If (t >= 180) Then
'				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorRadiobuttonEnTienda_"&Num_Iter&".png", True
'				imagenToWord "Error RadioButton En Tienda_"&Num_Iter,RutaEvidencias() & "ErrorRadiobuttonEnTienda_"&Num_Iter&".png"
'				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
'				DataTable("s_Detalle", dtLocalSheet) = "No se habilitó la opción -En tienda- de la siguiente pantalla"
'				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
'				ExitActionIteration
'			End If
'		Wend
End Sub
Sub TipoEnvio()
    While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("En tienda").Exist = false
    	wait 1
    Wend
	Select Case str_metodo_entrega
		Case "En tienda"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("En tienda").Set "ON"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MetodoPago"&".png" , True
				imagenToWord "Método de Pago", RutaEvidencias() &Num_Iter&"_"&"MetodoPago"&".png"
				wait 2
		Case "Delivery"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("Delivery").Set "ON"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png" , True
				imagenToWord "Canal de Venta", RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png"
				
					While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Exist) = False
							wait 1
					Wend
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Lookup-notValidated_2").Click


			
					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaTable("SearchJTable").Exist) = False
						wait 1
					Wend
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaTable("SearchJTable").SelectRow "#0"
				wait 1
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DireccionEnvio"&".png" , True
				imagenToWord "Dirección de Envio", RutaEvidencias() &Num_Iter&"_"&"DireccionEnvio"&".png"
				
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Click
				wait 1
					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Exist) = False
						wait 1
					Wend
				wait 2
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Buscar detalles de contacto").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Buscar detalles de contacto").Click
				wait 1
					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Exist) = False
						wait 1
					Wend
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaList("ComboBoxNative$1").Select "DNI"
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaEdit("TextFieldNative$1").Set "95141994"
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Buscar ahora").Click
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaTable("SearchJTable").SelectRow "#0"
				wait 1
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DireccionReceptor"&".png" , True
				imagenToWord "Direccion de Receptor", RutaEvidencias() &Num_Iter&"_"&"DireccionReceptor"&".png"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Click
				wait 1
				End If
				
					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Exist) = False
						wait 1
					Wend
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Set "QA"
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Número de teléfono del").Set "994361186"
				wait 2
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"EntregaDelivery"&".png", True
				imagenToWord "Entrega Delivery", RutaEvidencias() &Num_Iter&"_"&"EntregaDelivery"&".png"
				wait 1
		Case "Recojo en otra tienda"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("Recojo en otra tienda").Set "ON"
			wait 2
	End Select
	Wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Siguiente >").Click
	Call esperaCarga

	Dim i
	i=0
	While not i=180
		i=i+1
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("ID del Acuerdo de Facturación:").Exist(1) Then
			i=180
			Reporter.ReportEvent micPass, "Éxito en carga de ID del Acuerdo Financiero", "Se ha cargado correctamente los datos del ID del Acuerdo Financiero"	
			Call acuerdoFacturacion
			Call pagoInmediato
		ElseIf JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaObject("WindowsInternalFrameTitlePane").Exist(1) Then
			i=180
			Reporter.ReportEvent micPass, "Éxito en carga de ID del Acuerdo Financiero", "Se ha cargado correctamente los datos del ID del Acuerdo Financiero"	
			Call acuerdoFacturacion
			Call pagoInmediato
		ElseIf JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaButton("Pago inmediato").Exist(1) Then
			i=180
			Call pagoInmediato
		ElseIf JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist Then
			Call pagoInmediato
			
		ElseIf JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist Then
			i = 180
		Else 
			If i=180 Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorVentanaAcuerdoFacturacion_"&Num_Iter&".png", True
				imagenToWord "Error Ventana Acuerdo Facturación_"&Num_Iter,RutaEvidencias() & "ErrorVentanaAcuerdoFacturacion_"&Num_Iter&".png"
				Reporter.ReportEvent micFail, "Error en carga de pantalla Negociar Distribución de Cargos", "No cargó la pantalla que contiene el ID del Acuerdo Financiero"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorPantallaAcuerdoFinanciero.png", True
				imagenToWord "Error Pantalla Acuerdo Financiero",RutaEvidencias() & "ErrorPantallaAcuerdoFinanciero.png"
				ExitActionIteration
			End If
			
		End If
	Wend	
	

	t = 0
	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist) = False
		Wait 1
		t = t + 1
		If (t >= 180) Then
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorBotonVerContrato_"&Num_Iter&".png", True
			imagenToWord "Error Botón Ver Contrato_"&Num_Iter,RutaEvidencias() & "ErrorBotonVerContrato_"&Num_Iter&".png"
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No se habilitó el botón -Ver Contrato- de la siguiente pantalla"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
	Wend
	'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("Mostrar detalles de precio").Set "ON"
End Sub
Sub reemplazarCargo
Dim monto
	If str_Financiamiento = "NO" Then
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").Exist(3) Then
			filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").GetROProperty("rows")
			For Iterator = 0 To filas-1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").SelectRow "#"&Iterator
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").PressKey "C",micCtrl
					JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").PressKey "V",micCtrl
					str_titulo=JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").GetROProperty("text")
					str_titulo=Left(str_titulo,102+31)
					str_titulo = Replace(str_titulo,"Elemento    Actividad    Fecha de vencimiento    Único    Mensual    Mensual anterior    Impuesto     ","")
					str_titulo=Left(str_titulo,31)
					
				If (str_titulo = "Precio de Dispositivo Corporate") Then		
								
					monto=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").GetCellData (Iterator,3)
					If not monto="0,00" Then
						While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Reemplazar cargo").GetROProperty("enabled") = "0") = true
							wait 1
							t = t + 1
							If (t >= 180) Then
								DataTable("s_Resultado", dtLocalSheet) = "Fallido"
								DataTable("s_Detalle", dtLocalSheet) = "No se habilitó botón -Reemplazar cargo-"
								JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorBotónReemplazarCargoInhabilitado_"&Num_Iter&".png", True
								imagenToWord "Error Botón Reemplazar cargo Inhabilitado_"&Num_Iter,RutaEvidencias() & "ErrorBotónReemplazarCargoInhabilitado_"&Num_Iter&".png"
								Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
								ExitActionIteration
							End If
						Wend
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Reemplazar cargo").Click
						While (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").JavaEdit("Monto:").GetROProperty("enabled") = "0") = true
							wait 1
							t = t + 1
							If (t >= 180) Then
								DataTable("s_Resultado", dtLocalSheet) = "Fallido"
								DataTable("s_Detalle", dtLocalSheet) = "No se habilitó la ventana de Resumen de la Orden."
								JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorVentanaResumenOrden_"&Num_Iter&".png", True
								imagenToWord "Error Ventana Resumen de la Orden no encontrada_"&Num_Iter,RutaEvidencias() & "ErrorVentanaResumenOrden_"&Num_Iter&".png"
								Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
								ExitActionIteration
							End If
						Wend
						JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").JavaEdit("Monto:").Set "S/0,00"
						JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").JavaButton("Aceptar").Click
					End If
					Exit For
				End If
			Next
		End If
	End If
		'Bucle que espera "Validade y Ver Contrato"
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
End Sub
Sub GeneracionOrden()

	tiempo = 0
	Do
		While((JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist)) = False
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Click
			If DataTable("e_WIC_ContrCli", dtLocalSheet) = "SI" Then
				'Call EnviarOrden()
				

RunAction "WIC2", oneIteration

				Exit Do
			End If
		Wend
		
'		'Click "Validade y Ver Contrato"
'		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist(2) Then
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Click
'			Wait 5
'			'WIC Genera Contrato
'			If DataTable("e_WIC_ContrCli", dtLocalSheet) = "SI" Then
'				RunAction "WIC2 [WIC2_ContyValiIdenti]", oneIteration
'				Call EnviarOrden()
'				Exit Sub
'			End If
'			
'		End If
					
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist(3) Then
			var1 = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("attached text")
	   	 	JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		End  If
		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist(2) Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
		End If
		
			If tiempo>=180 Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorContrato_"&Num_Iter&".png", True
				imagenToWord "Error Carga Contrato_"&Num_Iter,RutaEvidencias() & "ErrorContrato_"&Num_Iter&".png"
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No se a cargado el contrato correctamente"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			else
				Reporter.ReportEvent micPass,"Contrato Exitoso","Se ha cargado el contrato correctamente"
			End If
	
	Loop While Not (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist or (var1 = "Contratos no Generados") or (var1 = "0"))

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist(1) Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Close
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").Set "ON"
	End If
	'Call EnviarOrden()
	
End Sub
Sub EnviarOrden()
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
	
	'Click en "Enviar orden"
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
	End If

	'Control de Mensajde de Validación
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(2) Then
		varlog = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable_3").GetCellData(0,0)
		If (varlog3 = varlog) Then
			DataTable("s_Resultado", dtLocalSheet) = "Mensaje de Validación"
			DataTable("s_Detalle", dtLocalSheet) = varlog2
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "MensajeValidacion.png", True
			imagenToWord "Mensaje de Validación",RutaEvidencias() & "MensajeValidacion.png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Aceptar").Click
		End If
	End If
	
	'Bucle que espera el envío de la orden
	t = 0
	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").Exist) = False
		wait 1
		t = t + 1
		If (t >= 180) Then
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorCargaNroOrden_"&Num_Iter&".png", True
			imagenToWord "Error Carga Nro Orden_"&Num_Iter,RutaEvidencias() & "ErrorCargaNroOrden_"&Num_Iter&".png"
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No cargó la ventana -Nro de Orden- de manera correcta"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
	Wend
	
'	Select Case DataTable("e_Ambiente", "Login [Login]")
'		Case "UAT8"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),8)
'			str_NroOrden = DataTable("s_Nro_Orden", dtLocalSheet) 
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Click
'			wait 2
'		Case "UAT4"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			str_NroOrden = DataTable("s_Nro_Orden", dtLocalSheet)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Click
'			wait 2
'		Case "UAT6"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			str_NroOrden = DataTable("s_Nro_Orden", dtLocalSheet)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Click
'			wait 2	
'		Case "UAT10"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			str_NroOrden = DataTable("s_Nro_Orden", dtLocalSheet)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Click
'			wait 2
'		Case "UAT13"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			str_NroOrden = DataTable("s_Nro_Orden", dtLocalSheet)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Click
'			wait 2	

    Dim text
    text = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").GetROProperty("text")
    DataTable("s_Nro_Orden", dtLocalSheet) = RTRIM(LTRIM((replace(text,"Orden",""))))
    WAIT 1
    
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").CaptureBitmap RutaEvidencias() & "Orden_Generada_.png", True
    imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_.png"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Click
    
	DataTable("s_Resultado", dtLocalSheet) = "Éxito"
	DataTable("s_Detalle", dtLocalSheet) = "Se envió la orden "&str_NroOrden&" correctamente"
End Sub
Sub GestionLogistica()
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
	
	t=0
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Exist)=False
		wait 1
		t = t + 1
		If (t >= 180) Then
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorCargaBuscarOrden_"&Num_Iter&".png", True
			imagenToWord "Error Carga Buscar Orden_"&Num_Iter,RutaEvidencias() & "ErrorCargaBuscarOrden_"&Num_Iter&".png"
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No cargó la ventana -Buscar orden- de manera correcta"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
	Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)									
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
	wait 1
	
	
		tiempo=0
		Do 
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
				nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("1 Registros").GetROProperty("attached text")
				tiempo=tiempo+1
				wait 1
			Else 
				If (tiempo >= 180) Then
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorVerCantRegistro_"&Num_Iter&".png", True
					imagenToWord "Error Cantidad Registro por Orden_"&Num_Iter,RutaEvidencias() & "ErrorVerCantRegistro_"&Num_Iter&".png"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar el Empuje de la Orden"
					Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
					ExitActionIteration
				End If
			End If
		Loop While Not (nroreg="1 Registros")
		wait 2
	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").SelectRow "#0"
		
		tiempo=0
			Do
				If (DataTable("s_Detalle", dtLocalSheet)="Por favor rellenar todas las identificaciones de equipos") or (DataTable("s_Detalle", dtLocalSheet)<>"Por favor rellenar todas las identificaciones de equipos") Then
					If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Cancelar").Exist(1) Then
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Cancelar").Click
						wait 2
					End If
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Gestionar logística").Click
					tiempo=tiempo+1
					wait 1
					t=0
					While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").Exist) = False
						wait 1
						t = t + 1
						If (t >= 180) Then
							JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorCargaSolicitarOrden_"&Num_Iter&".png", True
							imagenToWord "Error Carga Solicitar Orden_"&Num_Iter,RutaEvidencias() & "ErrorCargaSolicitarOrden_"&Num_Iter&".png"
							DataTable("s_Resultado", dtLocalSheet) = "Fallido"
							DataTable("s_Detalle", dtLocalSheet) = "No cargó la ventana -Buscar: Solicitar Orden- de manera correcta"
							Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
							ExitActionIteration
						End If
					Wend
					
					vardisp=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").GetCellData (1,4)
					If vardisp<>str_idDispositivo Then
						If str_MotivoCambio="CAEQ_EQUIPO Y SIM" Then
							JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#1","#4"
							Set shell = CreateObject("Wscript.Shell") 
							shell.SendKeys "{ENTER}"
							JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#1","#4",str_idDispositivo
							wait 1
							JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#2","#4"
							Set shell = CreateObject("Wscript.Shell") 
							shell.SendKeys "{ENTER}"
							JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#2","#4",str_idSim
							wait 1
						ElseIf str_MotivoCambio="CAEQ_SIM" Then
							JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#1","#4"
							Set shell = CreateObject("Wscript.Shell") 
							shell.SendKeys "{ENTER}"
							JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#1","#4",str_idSim
							wait 1
						ElseIf str_MotivoCambio="CAEQ_EQUIPO" Then
							JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#1","#4"
							Set shell = CreateObject("Wscript.Shell") 
							shell.SendKeys "{ENTER}"
							JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#1","#4",str_idDispositivo
							wait 1
						End If
					else
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#2","#4"
						Set shell = CreateObject("Wscript.Shell") 
						shell.SendKeys "{ENTER}"
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#2","#4",str_idSim
						wait 1
					End If
					JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Ingreso_Materiales_"&".png", True
					imagenToWord "Ingreso de Materiales", RutaEvidencias() &Num_Iter&"_"&"Ingreso_Materiales_"&".png"
					JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Validar y Crear Factura").Object.doClick()
					
					t = 0
					Do
						varhab=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").GetROProperty("enabled")					
						wait 1
						t = t + 1
						If (t >= 180) Then
							JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorBotonEnviar_"&Num_Iter&".png", True
							imagenToWord "Error Boton Enviar_"&Num_Iter,RutaEvidencias() & "ErrorBotonEnviar_"&Num_Iter&".png"
							DataTable("s_Resultado", dtLocalSheet) = "Fallido"
							DataTable("s_Detalle", dtLocalSheet) = "No se habilitó el botón -Enviar- de Solicitar Orden de manera correcta"
							Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
							ExitActionIteration
						End If
					Loop While Not((JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (varhab="1"))
					
				
						If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist(1) or JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(1) Then
							If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist(0) Then
								varlog = JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text") 
							End If
							If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(0) Then
								varlog = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
							End If
							DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				       		DataTable("s_Detalle", dtLocalSheet) = varlog
				       		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
				     		If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").JavaButton("OK").Exist(1) Then
				        		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").JavaButton("OK").Click
							End If
							If 	JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist(1) Then
								JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
							End If
							wait 2
							If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Cancelar").Exist(1) Then
								JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Cancelar").Click
								wait 2
							End If
				     		If DataTable("s_Detalle", dtLocalSheet)<>"Por favor rellenar todas las identificaciones de equipos" Then
								If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist(2) Then
									JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
									ExitActionIteration
								End If	
				     		End  If
				    	End If
				End  If
				If tiempo>=20 Then
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorAsignarSeries_"&Num_Iter&".png", True
					imagenToWord "Error Asignar Series_"&Num_Iter,RutaEvidencias() & "ErrorAsignarSeries_"&Num_Iter&".png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)  
					DataTable("s_Resultado",dtLocalSheet) = "Fallido"
					DataTable("s_Detalle",dtLocalSheet) = "Luego de 20 intentos no se pudo realizar la Asignación de Series"
					ExitActionIteration
				else
					Reporter.ReportEvent micPass, "Exito", "Se realizo la Asignación de Series correctamente"
			End If
		Loop While Not varhab = "1"
		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").Exist(3) Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").Click
		End If
End Sub
Sub EmpujeOrden()
	If DataTable("e_Tipo_Data", dtLocalSheet) = "DATA LOGICA" Then
	
		wait 2
		
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select @@ hightlight id_;_29361700_;_script infofile_;_ZIP::ssf109.xml_;_
		t=0
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Exist)=False
			wait 1
			t = t + 1
			If (t >= 180) Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorGrupoOrdenes_"&Num_Iter&".png", True
				imagenToWord "Error Grupo Ordenes_"&Num_Iter,RutaEvidencias() & "ErrorGrupoOrdenes_"&Num_Iter&".png"
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la ventana -Grupo de órdenes- de manera correcta"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
		
		t=0
		Do
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Select "Tareas pendientes del equipo"
			wait 2
			t = t + 1
			If (t >= 180) Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorBotonFinalizarCompra_"&Num_Iter&".png", True
				imagenToWord "Error Botón Finalizar Compra y Activar_"&Num_Iter,RutaEvidencias() & "ErrorBotonFinalizarCompra_"&Num_Iter&".png"
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No salió de la ventana -Grupo de órdenes- de manera correcta"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Finalizar compra y activar").Exist
		wait 2
	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").SetFocus @@ hightlight id_;_14588149_;_script infofile_;_ZIP::ssf111.xml_;_
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click @@ hightlight id_;_7641522_;_script infofile_;_ZIP::ssf113.xml_;_
		
		tiempo=0
		Do 
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
				nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("-- Registros").GetROProperty("text")
				tiempo=tiempo+1
				wait 1
			End If
			
			If (tiempo >= 180) Then
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorCantRegistro_"&Num_Iter&".png", True
					imagenToWord "Error Cantidad de Registros_"&Num_Iter,RutaEvidencias() & "ErrorCantRegistro_"&Num_Iter&".png"
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar el Empuje de la Orden"
					Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
					wait 2
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
					wait 2
					ExitActionIteration
					wait 2
			End If
		
		Loop While Not(nroreg="1 Registros")
		wait 1
		
		
		tiempo=0
			Do
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
					wait 2
					tiempo = tiempo+1
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").Output CheckPoint("Equipo usuario:") @@ hightlight id_;_4113048_;_script infofile_;_ZIP::ssf1.xml_;_
					varValidaRespuestaCumplimiento = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").GetCellData (0,5)
					wait 1
				End If
					If (tiempo >= 180) Then
						JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorMensajeRespuesta_"&Num_Iter&".png", True
						imagenToWord "Error Mensaje de Respuesta de Cumplimiento_"&Num_Iter,RutaEvidencias() & "ErrorMensajeRespuesta_"&Num_Iter&".png"
						DataTable("s_Resultado",dtLocalSheet)="Fallido"
						DataTable("s_Detalle",dtLocalSheet)="La actividad 'Manejar Respuesta de Cumplimiento' no cargo"	
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
						ExitTestIteration
					End If 
		Loop While Not varValidaRespuestaCumplimiento = "Manejar Respuesta de Cumplimiento"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").SelectRow "#0"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Gestión manual").Click
		t=0
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaList("Estado de la gestión manual:").Exist) = False
			wait 1
			t = t + 1
			If (t >= 180) Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorBotonEnviar_"&Num_Iter&".png", True
				imagenToWord "Error Botón Enviar_"&Num_Iter,RutaEvidencias() & "ErrorBotonEnviar_"&Num_Iter&".png"
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la ventana -Grupo de órdenes- de manera correcta"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaList("Estado de la gestión manual:").Select "Cumplimiento Completo"
		Wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaList("Motivo de la gestión manual").Select "Manejo manual: Manejo Manual OSS"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").Click
		wait 1
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "EmpujeOK.png", True
		imagenToWord "Empuje OK",RutaEvidencias() & "EmpujeOK.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
	End If
	
End Sub
Sub OrdenCerrado()

		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
		wait 1 @@ hightlight id_;_27981779_;_script infofile_;_ZIP::ssf1.xml_;_
		t=0
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Exist) = False
			wait 1
			t = t + 1
			If (t >= 180) Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorCargaBuscarOrden_"&Num_Iter&".png", True
				imagenToWord "Error Carga Buscar Orden Cerrado_"&Num_Iter,RutaEvidencias() & "ErrorCargaBuscarOrden_"&Num_Iter&".png"
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la ventana -Buscar Órden- de manera correcta"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend

		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").SetFocus
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet) @@ hightlight id_;_19889480_;_script infofile_;_ZIP::ssf3.xml_;_
		wait 1 @@ hightlight id_;_22342896_;_script infofile_;_ZIP::ssf8.xml_;_
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaCheckBox("Solo órdenes pendientes").Set "OFF" @@ hightlight id_;_23592401_;_script infofile_;_ZIP::ssf22.xml_;_
		wait 1
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click @@ hightlight id_;_12440768_;_script infofile_;_ZIP::ssf4.xml_;_
		wait 8
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:") @@ hightlight id_;_32738597_;_script infofile_;_ZIP::ssf67.xml_;_
		Reporter.ReportEvent micPass,"Se valida el estado de la orden",  DataTable("s_ValEstadoOrden", dtLocalSheet)
		
		tiempo = 0
		Do
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:") @@ hightlight id_;_24634058_;_script infofile_;_ZIP::ssf1.xml_;_
			tiempo = tiempo + 1
				If (tiempo>=180) Then		
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorEstadoCerrado_"&Num_Iter&".png", True
					imagenToWord "Error Estado de Orden_"&Num_Iter,RutaEvidencias() & "ErrorEstadoCerrado_"&Num_Iter&".png"
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "La Orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" no culmino en estado Cerrado"
					Reporter.ReportEvent micFail,"Error al finalizar la orden","Es probable que la orden termine con tiempo excedido"
					ExitActionIteration
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
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
		
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist(3) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
			wait 1			
		End If
		
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(2) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
			wait 1
		End If
		
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Exist(3) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click	
			wait 1
		End If
		
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Exist(3) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click	
			wait 1
		End If
		If str_ValPenalidad ="SI" Then
			
		End If

End Sub
Sub cerrarVentanas
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Exist(3) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click	
			wait 1
		End If
'	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_3").JavaButton("Cerrar").Exist(3) Then
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_3").JavaButton("Cerrar").Click	
'			wait 1
'		End If	
'	If JavaWindow("Ejecutivo de interacción").JavaDialog("Guardar el formulario").JavaButton("Guardar").Exist(3) Then
'			JavaWindow("Ejecutivo de interacción").JavaDialog("Guardar el formulario").JavaButton("Guardar").Click	
'			wait 1
'		End If	
End Sub





