Dim IdServiciol, OrdenPendiente, WIC_Activa, Equipo, Agrega, Orden,str_idDispositivo, Motivo, Tipo, PrecioEq, Plan


IdServicio 			= DataTable("e_IdServicio", dtLocalSheet)
Equipo 				= DataTable("e_Equipo",dtLocalSheet)
str_idDispositivo 	= DataTable("idDispositivo", dtLocalSheet)
Motivo				= DataTable("Motivo", dtLocalSheet)
Financiamiento		= DataTable("e_Financiamiento", dtLocalSheet)
Plan				= DataTable("e_Plan", dtLocalSheet)
Cuotas              = DataTable("e_Cuotas", dtLocalSheet)

Call BusquedaIdServicio()
Call VerificacionOrdenPend()
Call FlujoWIC()
Call CambioSimplificado()

If ucase(Motivo) = "CAEQ_EQUIPO" OR ucase(Motivo) = "CAEQ_CAPL" Then
    If ucase(DataTable("e_MetodoEntrega", dtLocalSheet))  = "DELIVERY"Then
    	
    	Call InstruccionesEnvio()
    End If
	If ucase(Financiamiento) = "SI" Then
		Call NegociarPago()
	End If
	
	Call ResumenOrden() 
	Call EnviarPago()
	Call GestionLogistica()
	Call EmpujeOrden()
	Call OrdenCerrada()
End If

If ucase(Motivo) = "CAPL" Then
	Call ResumenOrden()
	'Call EnviarPago()
	'Call EmpujeOrden()
	Call OrdenCerrada()
End If
Sub BusquedaIdServicio()

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
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"PanelInteraccion.png", True
	imagenToWord "Visualización Panel de Interacción",RutaEvidencias() &Num_Iter&"_"&"PanelInteraccion.png"
	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTab("Acciones rápidas").Select "Suscripciones" @@ hightlight id_;_8326975_;_script infofile_;_ZIP::ssf1.xml_;_
	
		t=0
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Buscar ahora_2").Exist = False
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
	
	wait 3
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaEdit("TextFieldNative$1").Set IdServicio @@ hightlight id_;_21817730_;_script infofile_;_ZIP::ssf2.xml_;_
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Buscar ahora_2").Click
	wait 3
	filas=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Acciones rápidas").GetROProperty("rows") 
	For i = filas-1 To 0 Step -1
		varestado=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Acciones rápidas").GetCellData(i,"Estado")
		If varestado="Activo" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Acciones rápidas").SelectRow(i)
			wait 1
		End If
	Next
	wait 2
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"IdServicioBuscado.png", True
	imagenToWord "Visualización del Id de Servicio Buscado",RutaEvidencias() &Num_Iter&"_"&"IdServicioBuscado.png"
	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Ver Productos Asignados").Click
End Sub
Sub VerificacionOrdenPend()

		t=0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Calcular Penalidad").Exist) = False
			Wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtlocalSheet) = "Fallido"
				DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga del Panel de Interaccion"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorDetProdAsig.png", True
				imagenToWord "Error en la Carga de Detalle del Producto Asig",RutaEvidencias() &Num_Iter&"_"&"ErrorDetProdAsig.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
				Wait 2
				ExitActionIteration
			End If	
		Wend

'	OrdenPendiente = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaTab("Antigüedad de línea:")
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaTab("Antigüedad de línea:").Select "Órdenes pendientes [Ninguna]"
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaTab("Antigüedad de línea:").Select "Órdenes pendientes" @@ hightlight id_;_16401507_;_script infofile_;_ZIP::ssf7.xml_;_
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaTab("Antigüedad de línea:").Select "Configuración"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaTab("Antigüedad de línea:").Type micRight
	
	strNombre=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaTab("Antigüedad de línea:").GetROProperty("value")
	
	If strNombre="Conexiones" or strNombre="Conexiones [Ninguna]" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaTab("Antigüedad de línea:").Type micRight
	Else 
		If strNombre="Órdenes pendientes [Ninguna]" Then
			DataTable("s_Resultado", dtLocalSheet) = "Exitoso"
			DataTable("s_Detalle", dtLocalSheet) = "El número "&str_IDServicio&" no posee Ordenes pendientes"
			Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "SinOrdenPend.png", True
			imagenToWord "El Número no posee Orden Pendiente",RutaEvidencias() & "SinOrdenPend.png"
		Else 
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Cant_Ordenes_"&Num_Iter&".png", True
			imagenToWord "Error_Cant_Ordenes_"&Num_Iter,RutaEvidencias() & "Cant_Ordenes_"&Num_Iter&".png"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaTab("Antigüedad de línea:").Select "Órdenes pendientes"
			rowOrdenCnt=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaTable("Antigüedad de línea:").GetROProperty("rows")
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "El número "&str_IDServicio&" posee Órdenes pendientes"
			Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "OrdenPend.png", True
			imagenToWord "El Numero posee Orden Pendiente",RutaEvidencias() & "OrdenPend.png"
			ExitActionIteration
		End If
	End If	
	wait 1
	
	JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").JavaMenu("Cambiar express").Select
	
End Sub
Sub FlujoWIC()
	If ucase(DataTable("e_WIC_ValidaCli", dtLocalSheet)) = "SI" Then
		wait 1
		


RunAction "WIC", oneIteration
	End If
End Sub
Sub CambioSimplificado()

 @@ hightlight id_;_29628800_;_script infofile_;_ZIP::ssf8.xml_;_
	If ((ucase(Motivo) = "CAEQ_EQUIPO") or (ucase(Motivo) = "CAEQ_CAPL")) Then
	
			While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("Su búsqueda devolvió una").Exist = False
				wait 1
			Wend
			
			t=0
			While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Reemplazar dispositivo").Exist = False
				Wait 1
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtlocalSheet) = "Fallido"
					DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga del Cambio Simplificado"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCambioSimplificado.png", True
					imagenToWord "Error en la Carga del Cambio Simplificado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCambioSimplificado.png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
					Wait 2
					ExitActionIteration
				End If	
			Wend
			
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "AltaExpress.png", True
		imagenToWord "Alta Express",RutaEvidencias() & "AltaExpress.png"
		
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{PGDN}"
			
		If ucase(Motivo) = "CAEQ_CAPL" Then
			Call CambioPlan()
		End If	
	
		varRem=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Reemplazar dispositivo").GetROProperty("enabled")
	
		If varRem="0" Then
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").Set "ON"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").Set "OFF"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Reemplazar dispositivo").Click
		Else 
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Reemplazar dispositivo").Click
		End If
		
			t=0
			While JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaStaticText("Mostrando 1-6 de 20 equipos").Exist = False
				Wait 1
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtlocalSheet) = "Fallido"
					DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga del Cambio Simplificado"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCambioSimplificado.png", True
					imagenToWord "Error en la Carga del Cambio Simplificado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCambioSimplificado.png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
					Wait 2
					ExitActionIteration
				End If	
			Wend
				
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaEdit("TextFieldNative$1").Set Equipo
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaButton("Buscar").Click

		'Esperará el movil [Huawei P10 Negro] - 14/11/2019 
		
		

		If ucase(Equipo) = "HUAWEI P10 NEGRO" Then
			While JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaStaticText("Mostrando 1-6 de 10 equipos(st").Exist = false
				wait 1
			Wend
			wait 2 
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaCheckBox("Seleccionar_2").Set "ON"
		Else 
			While (JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").Exist) = False
				wait 1
			Wend
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").Set "ON"	
		
		End If
		
		'Mitigación de un componente que no esta habilitado - - 14/11/2019 
		If Not JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaButton("Agregar").GetROProperty("enabled") Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaButton("Cancelar").Click
			varRem=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Reemplazar dispositivo").GetROProperty("enabled")
			If varRem="0" Then
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").Set "ON"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").Set "OFF"
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Reemplazar dispositivo").Click
			Else 
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Reemplazar dispositivo").Click
			End If
			
				t=0
				While JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaStaticText("Mostrando 1-6 de 20 equipos").Exist = False
					Wait 1
					t = t + 1
					If (t >= 180) Then
						DataTable("s_Resultado", dtlocalSheet) = "Fallido"
						DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga del Cambio Simplificado"
						JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCambioSimplificado.png", True
						imagenToWord "Error en la Carga del Cambio Simplificado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCambioSimplificado.png"
						Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
						Wait 2
						ExitActionIteration
					End If	
				Wend
					
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaEdit("TextFieldNative$1").Set Equipo
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaButton("Buscar").Click
			If ucase(Equipo) = "HUAWEI P10 NEGRO" Then
				While JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaStaticText("Mostrando 1-6 de 10 equipos(st").Exist = false
					wait 1
				Wend
				wait 2 
				JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaCheckBox("Seleccionar_2").Set "ON"
			Else 
				While (JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").Exist) = False
					wait 1
				Wend
				JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").Set "ON"	
			
			End If
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Dispos.png", True
			imagenToWord "Dispositivo seleccionado",RutaEvidencias() & "Dispos.png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaButton("Agregar").Click
		Else 
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Dispos.png", True
		imagenToWord "Dispositivo seleccionado",RutaEvidencias() & "Dispos.png"
			
		End If
		
			t=0
			While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Cambiar forma de pago").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist)) = False
				Wait 1
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtlocalSheet) = "Fallido"
					DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga del Cambio Simplificado"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCambioSimplificado.png", True
					imagenToWord "Error en la Carga del Cambio Simplificado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCambioSimplificado.png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
					Wait 2
					ExitActionIteration
				End If	
			Wend
			
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
			varmens=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
			If varmens="No se cuenta con stock del equipo HUAWEI Y360 BLANCO. Por favor seleccionar uno distinto." Then
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaCheckBox("Seleccionar_2").Set
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Cambio Simplificado (Para").JavaButton("Agregar").Click
				wait 1
			End If
		End If
		wait 5
		
		If ucase(DataTable("e_MetodoEntrega", dtlocalSheet)) = "DELIVERY" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaList("Metodo de Entrega").Select "Delivery"
		Else 
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaList("Metodo de Entrega").Select "En tienda"
		End If
		
		
		
'		if ucase(Tipo) = "FINANCIADO" then
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Cambiar forma de pago").Set "ON"
'		Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
'		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "CambioFormaPago.png", True
'		imagenToWord "Cambio de la forma de Pago",RutaEvidencias() & "CambioFormaPago.png"
'		End If
		
		'---------------------------DIFERENCIACION DE PRECIOS---------------------------------
'		If UCASE(DataTable("e_DiferPrec", dtLocalSheet)) = "SI" Then
'			

'		End If
		'----------------------------------------------------------------------------------

		
			
        '--------------PLAN DE FINANCIAMIENTO--------------------------
        
        If ucase(Financiamiento) = "SI" Then
        	Select Case Cuotas
        		Case "12"
        			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaList("Plan de Financiamiento:").Select "MOVISTAR-12 cuotas"
				 	wait 1
				 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Calcular").Click
				 	Dim y
				 	y = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("2628,00(st)").GetROProperty ("text")
				 	While (y = "0,00" or y = "0.00")
				 		wait 1
				 		y = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("2628,00(st)").GetROProperty ("text")
				 	Wend
					Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Financiado12.png", True
					imagenToWord "Se realiza financiamiento: 12 cuotas",RutaEvidencias() & "Financiado12.png" 
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("2628,00(st)").CaptureBitmap RutaEvidencias() & "Monto12.png", True
					imagenToWord "Monto financiado:  ",RutaEvidencias() & "Monto12.png" 
        		
        		Case "18"
        			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaList("Plan de Financiamiento:").Select "MOVISTAR-18 cuotas"
				 	wait 1
				 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Calcular").Click
				 	Dim x
				 	x = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("2628,00(st)").GetROProperty ("text")
				 	While (x = "0,00" or x = "0.00")
				 		wait 1
				 		x = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("2628,00(st)").GetROProperty ("text")
				 	Wend
					Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Financiado18.png", True
					imagenToWord "Se realiza el Cambio express de tipo Financiado 18 cuotas",RutaEvidencias() & "Financiado18.png"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("2628,00(st)").CaptureBitmap RutaEvidencias() & "Monto18.png", True
					imagenToWord "Monto financiado:  ",RutaEvidencias() & "Monto18.png" 
        	End Select
        	
        ElseIf ucase(Financiamiento) = "NO" Then
        	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Cambiar forma de pago").Set "OFF"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaList("Plan de Financiamiento:").Select "Contado"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Calcular").Click
			wait 10
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Montocont.png", True
			imagenToWord "Calcular Monto",RutaEvidencias() & "Montocont.png"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("2628,00(st)").CaptureBitmap RutaEvidencias() & "Contado.png", True
			imagenToWord "Monto Calculado:  Contado       ",RutaEvidencias() & "Contado.png"
        End If
        	
	END IF
	
'		If ucase(Motivo) = "CAPL" Then
'			Call CambioPlan()
'		End If
		
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "finExpress.png", True
	imagenToWord "Alta Express Terminada  ",RutaEvidencias() & "finExpress.png" 
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Siguiente >").Click
	
	If Ucase(Motivo)="CAEQ_CAPL" OR ucase(Motivo) = "CAPL" Then
		Call DetalleCambioSimplificado()
	End If
	
End Sub @@ hightlight id_;_25229544_;_script infofile_;_ZIP::ssf5.xml_;_
Sub CambioPlan()
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaList("ComboBoxNative$1").Select "Todas las subcategorías"
		wait 2	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaEdit("TextFieldNative$1").Set Plan	
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Buscar").Click
		wait 7
			t=0
			While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").Exist = False
				Wait 1
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtlocalSheet) = "Fallido"
					DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga del Cambio Simplificado"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCambioSimplificado.png", True
					imagenToWord "Error en la Carga del Cambio Simplificado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCambioSimplificado.png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
					Wait 2
					ExitActionIteration
				End If	
			Wend
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").Set "ON"
		
'		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").GetROProperty("displayed") = 0 Then
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaCheckBox("Seleccionar").Set "ON"
'		End If
		
		wait 1
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "CambioPlan.png", True
		imagenToWord "Cambio de Plan",RutaEvidencias() & "CambioPlan.png"	
End Sub
Sub InstruccionesEnvio()
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Exist = False
		Wait 1
			
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la negociación de la Distribución de envío"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrNegociarEnvio.png", True
			imagenToWord "Error en la Carga de la negociación de la Distribución de envío",RutaEvidencias() &Num_Iter&"_"&"ErrNegociarEnvio.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("Delivery").Exist = False
		Wait 1
			
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la negociación de la Dirección de envío"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrNegociarEnvio.png", True
			imagenToWord "Error en la Carga de la negociación de la Dirección de envío",RutaEvidencias() &Num_Iter&"_"&"ErrNegociarEnvio.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
		End If	
	Wend
	wait 3
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Set "Direccion de prueba QA"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Número de teléfono del").Set "999999999"
	Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Envio.png", True
	imagenToWord "Negociar dirección de envío",RutaEvidencias() & "Envio.png"
	wait 3
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Siguiente >").Click

End Sub
Sub NegociarPago()
	
		t=0
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaButton("Límite de Compra").Exist = False
			Wait 1	
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtlocalSheet) = "Fallido"
				DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la negociación del Pago"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrNegociarPago.png", True
				imagenToWord "Error en la Carga de la negociación del Pago",RutaEvidencias() &Num_Iter&"_"&"ErrNegociarPago.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
				Wait 2
				ExitActionIteration
			End If	
		Wend
		wait 1
		
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").Exist Then
		If ucase(Financiamiento) = "SI" Then
		
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaCheckBox("Financiamiento Externo").Set "ON"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaCheckBox("Financiamiento Externo").Set "OFF"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaCheckBox("Financiamiento Externo").Set "ON"
'			
				t=0
				While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaEdit("Importe de Cuota Mayor:").Exist = False
					Wait 1
					t = t + 1
					If (t >= 180) Then
						DataTable("s_Resultado", dtlocalSheet) = "Fallido"
						DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la negociación del Pago"
						JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrNegociarPago.png", True
						imagenToWord "Error en la Carga de la negociación del Pago",RutaEvidencias() &Num_Iter&"_"&"ErrNegociarPago.png"
						Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
						Wait 2
						ExitActionIteration
					End If	
				Wend
			
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaEdit("Importe de Cuota Mayor:").Set "100"
			wait 3
			Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "FinancExt.png", True
			imagenToWord "Negociar Financiamiento Externo",RutaEvidencias() & "FinancExt.png"
		End If	
		
		var2=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaButton("Límite de Compra").GetROProperty ("enabled")
		
		If var2 = "1" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaButton("Límite de Compra").Click	
		End If
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaButton("Pago inmediato").Click
			
		While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist))= False
			wait 1
		Wend
		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
			varsap=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel_2").GetROProperty("text")
			DataTable("s_Resultado",dtLocalSheet)="Fallido"
			DataTable("s_Detalle",dtLocalSheet)=varsap
			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MensajeValidacción"&".png", True
			imagenToWord "Mensaje de Validacción", RutaEvidencias() &Num_Iter&"_"&"MensajeValidacción"&".png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
			wait 1
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"NegociarPago"&".png", True
			imagenToWord "Negociar Pago", RutaEvidencias() &Num_Iter&"_"&"NegociarPago"&".png"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaButton("Siguiente >").Click
				While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validar y Ver Contrato").Exist)=False
					wait 1
				Wend
		End If
	End If	

    Dim d
    d = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("BAR ID").GetROProperty("value")
    While d = ""
    	wait 1
    	d = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("BAR ID").GetROProperty("value")
    Wend
   
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Exist = true Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Select "Boleta"
		Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "PagoInm.png", True
		imagenToWord "Pago Inmediato",RutaEvidencias() & "PagoInm.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Enviar").Click
	
			t=0
			While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaButton("Siguiente >").Exist = False
				Wait 1
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtlocalSheet) = "Fallido"
					DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la negociación del Pago"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrNegociarPago.png", True
					imagenToWord "Error en la Carga de la negociación del Pago",RutaEvidencias() &Num_Iter&"_"&"ErrNegociarPago.png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
					Wait 2
					ExitActionIteration
				End If	
			Wend
		
		Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "VerPago.png", True
		imagenToWord "Visualización Del Pago realizado correctamente",RutaEvidencias() & "VerPago.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaButton("Siguiente >").Click
	End If
	
End Sub

Sub ResumenOrden()

		t=0
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validar y Ver Contrato").Exist = False
			Wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtlocalSheet) = "Fallido"
				DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la negociación del Pago"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrNegociarPago.png", True
				imagenToWord "Error en la Carga de la negociación del Pago",RutaEvidencias() &Num_Iter&"_"&"ErrNegociarPago.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
				Wait 2
				ExitActionIteration
			End If	
		Wend
		
	Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ResOrden.png", True
	imagenToWord "Resumen de la orden",RutaEvidencias() & "ResOrden.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validar y Ver Contrato").Click
	
	If ucase(DataTable("e_WIC_ContrCli",dtlocalSheet)) = "SI" Then
		wait 1
		RunAction "WIC2", oneIteration
	else
	    While JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist = False
	    	wait 1
	    Wend
	    wait 2
	    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Link.png", True
		imagenToWord "Resumen de la orden: Link de documentación",RutaEvidencias() & "Link.png"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Close
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").Exist = False
			wait 1
		Wend
		wait 2
		estado = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").GetROProperty("enabled")
		If estado <> "0" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").Set "ON"
		End If
	End If
	
		t=0
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validar y Ver Contrato").Exist = False
			Wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtlocalSheet) = "Fallido"
				DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la negociación del Pago"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrNegociarPago.png", True
				imagenToWord "Error en la Carga de la negociación del Pago",RutaEvidencias() &Num_Iter&"_"&"ErrNegociarPago.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
				Wait 2
				ExitActionIteration
			End If	
		Wend
	Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Orden.png", True
	imagenToWord "Se Enviará la orden",RutaEvidencias() & "Orden.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click @@ hightlight id_;_12430702_;_script infofile_;_ZIP::ssf2.xml_;_
	
		t=0
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 836795A").JavaEdit("TextAreaNative$1").Exist = False
			Wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtlocalSheet) = "Fallido"
				DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la pantalla Orden Enviada"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrOrdEnviada.png", True
				imagenToWord "Error en la Carga de la Orden Enviada",RutaEvidencias() &Num_Iter&"_"&"ErrOrdEnviada.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
				Wait 2
				ExitActionIteration
			End If	
		Wend
		
	Orden=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 836795A").GetROProperty("text")
	Orden = replace(Orden,"Orden ","")
	DataTable("s_Orden", dtlocalSheet) = Orden
	Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "OrdenEnviada.png", True
	imagenToWord "Orden Enviada",RutaEvidencias() & "OrdenEnviada.png"
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 836795A").Close @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf5.xml_;_
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Exist(3) Then
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click	
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Cerrar").Exist Then
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Cerrar").Click
	End If

End Sub
Sub EnviarPago()

	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select
	
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
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Set DataTable("s_Orden", dtlocalSheet)
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
		
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
		
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").SelectRow "#0"
	WAIT 1 @@ hightlight id_;_2640397_;_script infofile_;_ZIP::ssf15.xml_;_
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Gestión manual").Click @@ hightlight id_;_42744_;_script infofile_;_ZIP::ssf16.xml_;_
		t=0
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Enviar").Exist)=False
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
	JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Enviar").Click	
End Sub
Sub GestionLogistica()

	wait 1
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
	wait 1
	
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
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set Orden									
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
						If  Motivo="CAEQ_EQUIPO" Then
							JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#1","#4"
							Set shell = CreateObject("Wscript.Shell") 
							shell.SendKeys "{ENTER}"
							JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#1","#4",str_idDispositivo
							wait 2
						End If
					else
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#2","#4"
						Set shell = CreateObject("Wscript.Shell") 
						shell.SendKeys "{ENTER}"
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#2","#4",str_idSim
						wait 2
					End If
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Ingreso_Materiales_"&".png", True
					imagenToWord "Ingreso de Materiales", RutaEvidencias() &Num_Iter&"_"&"Ingreso_Materiales_"&".png"
					JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Validar y Crear Factura").Object.doClick()
					
					t = 0
					Do
						varhab=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").GetROProperty("enabled")					
						wait 3
						t = t + 1
						If (t >= 180) Then
							JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorBotonEnviar_"&Num_Iter&".png", True
							imagenToWord "Error Boton Enviar_"&Num_Iter,RutaEvidencias() & "ErrorBotonEnviar_"&Num_Iter&".png"
							DataTable("s_Resultado", dtLocalSheet) = "Fallido"
							DataTable("s_Detalle", dtLocalSheet) = "No se habilitó el botón -Enviar- de Solicitar Orden de manera correcta"
							Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
							ExitActionIteration
						End If
					Loop While Not((JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist) Or (varhab="1"))
				
						If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist(1) Then
							If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist(0) Then
								varlog = JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text") 
							End If
'							If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(0) Then
'								varlog = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
'							End If
							DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				       		DataTable("s_Detalle", dtLocalSheet) = varlog
				       		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
				     		If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").JavaButton("OK").Exist(1) Then
				        		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").JavaButton("OK").Click
							End If
'							If 	JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist(1) Then
'								JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
'							End If
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
		wait 2
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").Exist(3) Then
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Enviar_"&Num_Iter&".png", True
			imagenToWord "Gestión logística"&Num_Iter,RutaEvidencias() & "Enviar_"&Num_Iter&".png"
			Reporter.ReportEvent micpass, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet) 
			JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").Click
		End If
End Sub
Sub EmpujeOrden()
	If DataTable("e_Tipo_Data", dtLocalSheet) = "DATA LOGICA" Then
	
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select
		
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
	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").SetFocus
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Set DataTable("s_Orden", dtlocalSheet)
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
		
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
					'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").Output CheckPoint("Equipo usuario:")
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
			While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").Exist) = False
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
Sub OrdenCerrada()

	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
	wait 1
	
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
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Orden", dtLocalSheet)
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaCheckBox("Solo órdenes pendientes").Set "OFF"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
	wait 8
		
	DataTable("s_ValEstadoOrden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").GetCellData("#0","#4")
		
		tiempo = 0
		Do
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
			wait 2
			DataTable("s_ValEstadoOrden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").GetCellData("#0","#4")
			tiempo = tiempo + 1
			If (tiempo>=180) Then		
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorEstadoCerrado_"&Num_Iter&".png", True
				imagenToWord "Error Estado de Orden_"&Num_Iter,RutaEvidencias() & "ErrorEstadoCerrado_"&Num_Iter&".png"
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "La Orden:"&DataTable("s_Orden", dtLocalSheet)&" no culmino en estado Cerrado"
				Reporter.ReportEvent micFail,"Error al finalizar la orden","Es probable que la orden termine con tiempo excedido"
				ExitActionIteration
					
			End If
		Loop While Not DataTable("s_ValEstadoOrden", dtLocalSheet) = "Cerrado"
		DataTable("s_Resultado", dtLocalSheet)="Exito"
		DataTable("s_Detalle", dtLocalSheet)="La orden finalizó correctamente"
		Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
		wait 2
		
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Cerrada_"&".png", True
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
		

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Exist(3) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click	
		wait 1
	End If
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Cerrar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Cerrar").Click
	End If
		'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Cerrar").Click

'	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Exist(3) Then
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click	
'		wait 1
'	End If

End Sub
Sub DetalleCambioSimplificado()
		t=0
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaTable("(Nuevo)").Exist = False
			Wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtlocalSheet) = "Fallido"
				DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga del Cambio Simplificado"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCambioSimplificado.png", True
				imagenToWord "Error en la Carga del Cambio Simplificado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCambioSimplificado.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
				Wait 2
				ExitActionIteration
			End If	
		Wend
		
		t=0
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaStaticText("Totales(st)").Exist = False
			Wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtlocalSheet) = "Fallido"
				DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga del Cambio Simplificado"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCambioSimplificado.png", True
				imagenToWord "Error en la Carga del Cambio Simplificado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCambioSimplificado.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
				Wait 2
				ExitActionIteration
			End If	
		Wend
		
		t=0
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Siguiente >").Exist = False
			Wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtlocalSheet) = "Fallido"
				DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga del Cambio Simplificado"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCambioSimplificado.png", True
				imagenToWord "Error en la Carga del Cambio Simplificado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCambioSimplificado.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
				Wait 2
				ExitActionIteration
			End If	
		Wend
		
	wait 2
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "DetCambioPlan.png", True
	imagenToWord "Detalles de Cambio de Plan",RutaEvidencias() & "DetCambioPlan.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Cambio Simplificado (Para").JavaButton("Siguiente >").Click
End Sub







