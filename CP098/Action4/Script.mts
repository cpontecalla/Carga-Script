Dim IdServiciol, OrdenPendiente, WIC_Activa, Equipo, Agrega, Orden,str_idDispositivo, Motivo, DocCliente, Compromiso, MetodoEntrega, Financiamiento, Cuotas

Equipo 				= DataTable("e_Equipo",dtLocalSheet)
WIC_Activa 			= DataTable("e_WIC_Activa", dtLocalSheet)
WIC2				= DataTable("e_WIC2", dtLocalSheet)
str_idDispositivo 	= DataTable("idDispositivo", dtLocalSheet)
str_idSim			= DataTable("e_IdSIM", dtLocalSheet)
Plan				= DataTable("e_Plan", dtLocalSheet)
MetodoEntrega		= DataTable("e_MetodoEntrega", dtLocalSheet)
Compromiso			= DataTable("e_Compromiso", dtLocalSheet)
Financiamiento		= DataTable("e_Financiamiento", dtLocalSheet)
Cuotas				= DataTable("e_CantCuotas", dtLocalSheet)
TipoAlta = DataTable("e_TipoAlta", dtLocalSheet)
Call PanelInteraccion()
Call FlujoWIC()
Call AltaExpress()
If UCASE(Financiamiento) = "SI" Then
	Call NegociarPago()
End If
Call ResumenOrden()
Call EnviarOrden()
If UCASE(MetodoEntrega) <> "DELIVERY" Then
	Call EnviarPago()
	Call GestionLogistica()
	Call EmpujeOrden()
	Call OrdenCerrada()
End If

Call OrdenCerrada()

Sub PanelInteraccion()

		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Ver Detalles").Exist = False
			wait 1
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
			wait 1
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
	wait 6
	JavaWindow("Ejecutivo de interacción").JavaTable("Titulo").DoubleClickCell "#8","#0"

End Sub
Sub FlujoWIC()
	If ucase(WIC_Activa) = "SI" Then
		

RunAction "WIC1", oneIteration
	End If
End Sub
Sub SeleccionarPlan()

	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaCheckBox("Seleccionar").Exist = false
		wait 1
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga del Formulario Alta Express"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaAltaExpress.png", True
			imagenToWord "Error en la Carga Alta Express",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaAltaExpress.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
		End If	
	Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaEdit("TextFieldNative$1").Set Plan
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaButton("Buscar").Click
	wait 8
	while ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaCheckBox("Seleccionar").Exist)or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JOptionPane").Exist))= false
		wait 1 
	Wend
	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JOptionPane").Exist = True Then
		DataTable("s_Resultado", dtlocalSheet) = "Fallido"
		DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de plan express"
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"ErrorCargaPlanExpress.png", True
		imagenToWord "Error en la Carga del plan express",RutaEvidencias() &"ErrorCargaPlanExpress.png"
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
		Wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		ExitTestIteration
	End If
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaCheckBox("Seleccionar").Set "ON"
	wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SeleccionPlan.png", True
	imagenToWord "Seleccionar Plan",RutaEvidencias() &Num_Iter&"_"&"SeleccionPlan.png"
End Sub
Sub SeleccionarEquipo()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaButton("Agregar Equipo").Click
		
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaDialog("Alta Express (Para WALTHER").JavaStaticText("Mostrando 1-6 de 20 equipos").Exist = False
		Wait 1
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de alta express"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaAltaExpress.png", True
			imagenToWord "Error en la Carga del alta express",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaAltaExpress.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
		End If	
	Wend
		
	JavaWindow("Ejecutivo de interacción").JavaDialog("Alta Express (Para WALTHER").JavaEdit("TextFieldNative$1").Set Equipo
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaDialog("Alta Express (Para WALTHER").JavaButton("Buscar").Click
	
	

	If ucase(Equipo) = "HUAWEI P10 NEGRO" Then

	    While JavaWindow("Ejecutivo de interacción").JavaDialog("Alta Express (Para WALTHER").JavaStaticText("Mostrando 1-6 de 10 equipos(st").Exist = false
	    	wait 1
	    Wend
	    wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Alta Express (Para WALTHER").JavaCheckBox("Seleccionar_2").Set "ON"

	Else 
		t=0
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Alta Express (Para WALTHER").JavaCheckBox("Seleccionar").Exist = False
			Wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtlocalSheet) = "Fallido"
				DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de alta express"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaAltaExpress.png", True
				imagenToWord "Error en la Carga del Alta Express",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaAltaExpress.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
				Wait 2
				ExitActionIteration
			End If	
		Wend

		JavaWindow("Ejecutivo de interacción").JavaDialog("Alta Express (Para WALTHER").JavaCheckBox("Seleccionar").Set "ON"
	End If
	
	wait 2

	Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "EquipoSeleccionado.png", True
	imagenToWord "Modelo de telefono Buscado",RutaEvidencias() & "EquipoSeleccionado.png"
		
	
		Agrega = JavaWindow("Ejecutivo de interacción").JavaDialog("Alta Express (Para WALTHER").JavaButton("Agregar").GetROProperty("enabled")
		t=0
		While Agrega = 0
			Wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtlocalSheet) = "Fallido"
				DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de Alta Express"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaAltaExpres.png", True
				imagenToWord "Error en la Carga de alta express",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaAltaExpres.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
				Wait 2
				ExitActionIteration
			End If	
			Agrega = JavaWindow("Ejecutivo de interacción").JavaDialog("Alta Express (Para WALTHER").JavaButton("Agregar").GetROProperty("enabled")
		Wend
		
	JavaWindow("Ejecutivo de interacción").JavaDialog("Alta Express (Para WALTHER").JavaButton("Agregar").Click
End Sub


Sub AltaExpress()

	Select Case TipoAlta
		Case "Alta solo SIM"
		  Call SeleccionarPlan()
		Case "Alta equipo + SIM"
		  Call SeleccionarPlan()
		  Call SeleccionarEquipo()
	End Select

		t=0
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaCheckBox("Cambiar forma de pago").Exist = False
			Wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtlocalSheet) = "Fallido"
				DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de Alta Express"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaAltaExpress.png", True
				imagenToWord "Error en la Carga de Alta Express",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaAltaExpress.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
				Wait 2
				ExitActionIteration
			End If	
		Wend
	wait 5
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN 2}"
	wait 1
	If ucase(DataTable("e_Compromiso", dtlocalSheet)) <> empty Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaList("Compromiso").Select (DataTable("e_Compromiso", dtlocalSheet))
	End If
	
	If UCASE(MetodoEntrega) = "DELIVERY" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaList("Metodo de Entrega").Select "Delivery"
		wait 2
	End If
	
	If ucase(Financiamiento) = "SI" and TipoAlta = "Alta equipo + SIM" Then
		Select Case Cuotas
			Case "18"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaList("Plan de Financiamiento:").Select "MOVISTAR-18 cuotas"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaButton("Calcular").Click
				Dim g
				g = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaStaticText("2629,00(st)").GetROProperty("text")
				while (g="0.00" or g="0,00") 
					wait 1
					g = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaStaticText("2629,00(st)").GetROProperty("text")
				wend
				wait 1	
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "financiamiento18.png", True
				imagenToWord "Plan de Financiamiento: 18 cuotas",RutaEvidencias() & "financiamiento18.png"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaStaticText("2629,00(st)").CaptureBitmap RutaEvidencias() & "Monto18.png", True
				imagenToWord "Monto: 18 cuotas",RutaEvidencias() & "Monto18.png"
			Case "12"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaList("Plan de Financiamiento:").Select "MOVISTAR-12 cuotas"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaButton("Calcular").Click
				Dim m
				m = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaStaticText("2629,00(st)").GetROProperty("text")
				while (m="0.00" or m="0,00")
					wait 1
					m = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaStaticText("2629,00(st)").GetROProperty("text")
				wend
				wait 1
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "financiamiento12.png", True
				imagenToWord "Plan de Financiamiento: 12 cuotas",RutaEvidencias() & "financiamiento12.png"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaStaticText("2629,00(st)").CaptureBitmap RutaEvidencias() & "Monto12.png", True
				imagenToWord "Monto: 12 cuotas",RutaEvidencias() & "Monto12.png"
		End Select
	ElseIf  ucase(Financiamiento) = "NO" and TipoAlta = "Alta equipo + SIM" Then
	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaList("Plan de Financiamiento:").Select "Contado"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaButton("Calcular").Click
		Dim k
		k = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaStaticText("2629,00(st)").GetROProperty("text")
		while (k="0.00" or k="0,00")
			wait 1
			k = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaStaticText("2629,00(st)").GetROProperty("text")
		wend
		wait 1
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "fContado.png", True
		imagenToWord "Plan de Financiamiento: Contado",RutaEvidencias() & "fContado.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaStaticText("2629,00(st)").CaptureBitmap RutaEvidencias() & "MontoContado.png", True
		imagenToWord "Monto: Contado       ",RutaEvidencias() & "MontoContado.png"
	End If
	
	Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "DetAlta.png", True
	imagenToWord "Detalle Alta",RutaEvidencias() & "DetAlta.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Alta Express (Para WALTHER").JavaButton("Siguiente >").Click
		
	If ucase(MetodoEntrega) = "DELIVERY" Then
		Call Delivery()
	End If
		
End Sub


Sub Delivery()
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Dirección").Exist = false
		wait 1
	Wend
    Dim r
	r = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Dirección").GetROProperty("text")

	While r = ""
		wait 1
		r = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Dirección").GetROProperty("text")
	Wend
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Buscar detalles de contacto").Exist Then
	    wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Buscar detalles de contacto").Click
		t=0
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Buscar ahora").Exist = False
			Wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtlocalSheet) = "Fallido"
				DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de Alta Express"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorDelivery.png", True
				imagenToWord "Error en la Carga de Detalle Delivery",RutaEvidencias() &Num_Iter&"_"&"ErrorDelivery.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
				Wait 2
				ExitActionIteration
			End If	
		Wend
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaList("ComboBoxNative$1").Select "CE"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaEdit("TextFieldNative$1").Set DocCliente
		wait 1	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Buscar ahora").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaTable("SearchJTable").SelectRow "#0"
		wait 2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SelecContacto.png", True
		imagenToWord "Selección de Contacto",RutaEvidencias() &Num_Iter&"_"&"SelecContacto.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Click
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Exist = False
			wait 1
		Wend
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Lookup-Validated").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Lookup-Validated").Click
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_3").JavaButton("Buscar ahora").Exist = false
			wait 1
		Wend
		wait 1 
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_3").JavaTable("SearchJTable").Exist = false
			wait 1
		Wend
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_3").JavaButton("1 Registros").GetROProperty("text") <> "1 Registros" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_3").JavaTable("SearchJTable").SelectRow "#0"
		End If

		wait 3
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DireccionEnvio"&".png" , True
		imagenToWord "Dirección de Envio", RutaEvidencias() &Num_Iter&"_"&"DireccionEnvio"&".png"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_3").JavaButton("Seleccionar").Click
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Exist = false
			wait 1
		Wend
	End If
	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Set "PRUEBAS QA"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Número de teléfono del").Set "999999999"
		wait 2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SelecContacto.png", True
		imagenToWord "Selección de Contacto",RutaEvidencias() &Num_Iter&"_"&"SelecContacto.png"
		wait 2
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
	
		t=0
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaCheckBox("Financiamiento Externo").Exist = False
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
		
	If ucase(Financiamiento) = "SI" Then
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
			
		wait 5
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaEdit("Importe de Cuota Mayor:").Set "1"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaButton("Límite de Compra").Click
		wait 2
		Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "FinancExt.png", True
		imagenToWord "Negociar Financiamiento Externo",RutaEvidencias() & "FinancExt.png"
		
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaButton("Límite de Compra").GetROProperty("enabled") = 1 Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaButton("Límite de Compra").Click
	End If
	
	
	Agrega = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaButton("Pago inmediato").GetROProperty("enabled")
	t=0
		While Agrega = 0
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
			Agrega = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaButton("Pago inmediato").GetROProperty("enabled")
		Wend
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaButton("Pago inmediato").Click
	wait 5
	
	If (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist) Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
	End If
	wait 2
'	Dim Val
'	Val = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Numero RUC").GetROProperty("text")
'	If Val ="" Then
	'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Select "Boleta"
'	else
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Select "Factura"

'	End If

	'Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
	'JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "PagoInm.png", True
	'imagenToWord "Pago Inmediato",RutaEvidencias() & "PagoInm.png"
	'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Enviar").Click

		t=0
		While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaButton("Siguiente >").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JOptionPane").Exist) )= False
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
	
	If (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JOptionPane").Exist) = True Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click	
	End If
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("BAR ID").Exist = False
		wait 1
	Wend
	wait 1
	Dim y 
	y = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("BAR ID").GetROProperty("text")
	While y = ""
		wait 1
		y = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("BAR ID").GetROProperty("text")
	Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Select "Boleta"
	Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "VerPago.png", True
	imagenToWord "Visualización Del Pago realizado correctamente",RutaEvidencias() & "VerPago.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Enviar").Click

	wait 4
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "_Click_Pantalla_Siguiente_.png", True
	imagenToWord "-Negociar Pago Correcto-",RutaEvidencias() & "_Click_Pantalla_Siguiente_.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 833910A").JavaButton("Siguiente >").Click
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

	If ucase(WIC2) = "SI" Then
		
RunAction "WIC2", oneIteration

		Exit Sub
	End If
	
		t=0
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist = False
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
	wait 2
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ContratosOrden.png", True
	imagenToWord "Contratos de la orden",RutaEvidencias() & "ContratosOrden.png"
	JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Close
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").Set "ON"
	wait 2
End Sub
Sub EnviarOrden()
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
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 836795A").Close
	
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
		While((JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Enviar").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_3").JavaButton("Validar y Crear Factura").Exist))=False
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
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_3").JavaButton("Validar y Crear Factura").Exist = True Then
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_3").JavaButton("Cancelar").Click
		wait 2
		Exit Sub
	End If
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "EnviarPago.png", True
	imagenToWord "Enviar Pago",RutaEvidencias() & "EnviarPago.png"
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
				Select Case TipoAlta
					Case "Alta solo SIM"
					   JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#2","#4"
						Set shell = CreateObject("Wscript.Shell") 
						shell.SendKeys "{ENTER}"
						wait 2
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#2","#4",str_idSim
						wait 2
					Case "Alta equipo + SIM"
					  JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#1","#4"
						Set shell = CreateObject("Wscript.Shell") 
						shell.SendKeys "{ENTER}"
						wait 2
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#1","#4",str_idDispositivo
						wait 2
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#2","#4"
						Set shell = CreateObject("Wscript.Shell") 
						shell.SendKeys "{ENTER}"
						wait 2
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#2","#4",str_idSim
						wait 2
				End Select
							
					JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Ingreso_Materiales_"&".png", True
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
		wait 2
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").Exist(3) Then
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
			wait 3
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaList("Estado de la gestión manual:").Select "Cumplimiento Completo"
		Wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaList("Motivo de la gestión manual").Select "Manejo manual: Manejo Manual OSS"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").Click
		wait 1
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "EmpujeOK.png", True
		imagenToWord "Empuje OK",RutaEvidencias() & "EmpujeOK.png"
		if JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist= true then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
		End If
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

End Sub


