Dim var1, var2,var3, t, vargestion, varerror, nroreg, Num_Iter, filas, varValidaRespuestaCumplimiento, varasig, varasig2

Dim str_departamento
Dim str_provincia
Dim str_modeloCel
Dim str_tipoPlan
Dim str_tipoSIM
Dim str_valEstadoOrden
Dim str_tipoalta
Dim str_motivo_alta
Dim intStartTime, intStopTime
Dim str_idDispositivo
Dim tiempo
intStartTime = Timer
str_departamento	=	DataTable("e_Departamento", dtLocalSheet)
str_provincia		=	DataTable("e_Provincia", dtLocalSheet) 
str_modeloCel		=	DataTable("e_ModeloCelular", dtLocalSheet)
str_plan_comp		=	DataTable("e_TipoDePlan", dtLocalSheet)
str_tipoSIM			=	DataTable("e_TipoSIM", dtLocalSheet)
str_tipoalta		=	DataTable("e_TipodeAlta", dtLocalSheet)
str_motivo_alta		=	DataTable("e_MotivoAlta", dtLocalSheet)
str_Financiamiento	=	DataTable("e_Financiamiento", dtLocalSheet)
str_Cuotas			=  	DataTable("e_Cuotas", dtLocalSheet)
str_Permanencia		= 	DataTable("e_Permanencia", dtLocalSheet)
str_Periodo			=	DataTable("e_Periodo", dtLocalSheet)
str_MedioPago		=   DataTable("e_MetodoPago", dtLocalSheet)
str_TipoComprobante	=   DataTable("e_TipoPago", dtLocalSheet)
str_idSim			=	DataTable("e_ID_SIM", dtLocalSheet)
str_idDispositivo	=	DataTable("e_ID_Dispositivo", dtLocalSheet)
Num_Iter 			= 	Environment.Value("ActionIteration")

Call SeleccionarTipoAlta()
Call SeleccionarUbica()
Call SeleccionarEquipoMovil()
Call SeleccionarPlanTarifario()
Call ParametrosAlta()
Call Portabilidad()
Call TipoEnvio()
Call NegociarDistribucion()
Call NegociarPago()
Call GeneracionOrden()
'If DataTable("e_Ambiente", "Login [Login]")<>"PROD" Then
Call PagoManual()
'End If
Call GestionLogistica()
'If DataTable("e_Ambiente", "Login [Login]")<>"PROD" Then
'End If
Call OrdenCerrado()
Call DetalleActividadOrden()
Call ValidacionPortaflow()


Sub SeleccionarTipoAlta()
	wait 1
	Select Case DataTable("e_TipodeAlta", dtLocalsheet)
		Case "Alta Nueva Equipo + Linea"
			wait 5
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Panel_Interaccion_"&Num_Iter&".png", True
			imagenToWord "Panel de Interacción", RutaEvidencias() &"Panel_Interaccion_"&Num_Iter&".png"
			JavaWindow("Ejecutivo de interacción").JavaTable("Titulo").ActivateRow "#4"
			wait 1
		Case "Alta Nueva Solo Linea"
			wait 5
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Panel_Interaccion_"&Num_Iter&".png", True
			imagenToWord "Panel de Interacción", RutaEvidencias() &"Panel_Interaccion_"&Num_Iter&".png"
			JavaWindow("Ejecutivo de interacción").JavaTable("Titulo").ActivateRow "#5"
			wait 1
	End	Select
	
End Sub
Sub SeleccionarUbica()

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Autenticación del Cliente").Exist Then
		wait 2
		RunAction "WIC", oneIteration
	End If

			While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Ubicación").JavaList("Departamento:").Exist) = False
				wait 1	
			Wend
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Ubicación").JavaList("Departamento:").Select DataTable("e_Departamento", dtLocalSheet)
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Ubicación").JavaList("Provincia:").Select DataTable("e_Provincia", dtLocalSheet)
		wait 1
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "SeleccionarUbicacion_"&Num_Iter&".png", True
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Ubicación").JavaButton("Siguiente >").Click
			While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para POOL").JavaList("ComboBoxNative$1").Exist) Or(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaList("ComboBoxNative$1").Exist)) = False
				wait 1	
			Wend
End Sub
Sub SeleccionarEquipoMovil()

	If (DataTable("e_TipodeAlta", dtLocalSheet) = "Alta Nueva Equipo + Linea") Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para POOL").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 10000
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para POOL").JavaList("ComboBoxNative$1").Select "Celulares"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para POOL").JavaEdit("TextFieldNative$1").Set DataTable("e_ModeloCelular", dtLocalSheet)
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para POOL").JavaButton("Buscar").Click
		wait 1
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist(2) Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").close
		End If
		wait 1
		
		If DataTable("e_ModeloCelular", dtLocalSheet)<>"HUAWEI P10 NEGRO" Then
			tiempo=0
			Do
			tiempo=tiempo+1
				If tiempo>=60 Then
					DataTable("s_Resultado",dtLocalSheet) = "Fallido" 
					DataTable("s_Detalle", dtLocalSheet) = "El Equipo Móvil: "&DataTable("e_ModeloCelular", dtLocalSheet)&"no se encuentra"
					Reporter.ReportEvent micFail,DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
				Else
					Reporter.ReportEvent micPass, "Exito","Se encontro el equipo móvil buscado"
				End If
				wait 1
			Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para POOL").JavaButton("Agregar al carrito").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist))
			wait 1
		else 
			tiempo=0
			Do
			tiempo=tiempo+1
				If tiempo>=60 Then
					DataTable("s_Resultado",dtLocalSheet) = "Fallido" 
					DataTable("s_Detalle", dtLocalSheet) = "El Equipo Móvil: "&DataTable("e_ModeloCelular", dtLocalSheet)&"no se encuentra"
					Reporter.ReportEvent micFail,DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
				Else
					Reporter.ReportEvent micPass, "Exito","Se encontro el equipo móvil buscado"
				End If
				wait 1
			Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para POOL").JavaButton("Agregar al carrito_2").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist))
			wait 2
		End If		

		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
			varsap=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
			DataTable("s_Resultado",dtLocalSheet)="Fallido"
			DataTable("s_Detalle",dtLocalSheet)=varsap
			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Error SAP"&".png", True
			imagenToWord "Error SAP", RutaEvidencias() &Num_Iter&"_"&"Error SAP"&".png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para POOL").JavaButton("Cerrar").Click
			wait 1
			ExitActionIteration
		End If
		'End If

		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
			varsap=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
			DataTable("s_Resultado",dtLocalSheet)="Fallido"
			DataTable("s_Detalle",dtLocalSheet)=varsap
			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Error SAP"&".png", True
			imagenToWord "Error SAP", RutaEvidencias() &Num_Iter&"_"&"Error SAP"&".png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para POOL").JavaButton("Cerrar").Click
			wait 1
			ExitActionIteration
		End If

		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Equipo Móvil"&".png", True
		imagenToWord "Equipo Móvil", RutaEvidencias() &Num_Iter&"_"&"Equipo Móvil"&".png"
		wait 1
		
		If DataTable("e_ModeloCelular", dtLocalSheet)<>"HUAWEI P10 NEGRO" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para POOL").JavaButton("Agregar al carrito").Click
		else
			'MsgBox "Seleccionar el equipo movil"	
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para POOL").JavaButton("Agregar al carrito_2").Click
		End If
		
			While((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaButton("Buscar").Exist)) = False
				wait 1
			Wend 

		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
			varsap=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
			DataTable("s_Resultado",dtLocalSheet)="Fallido"
			DataTable("s_Detalle",dtLocalSheet)=varsap
			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Error_SAP_"&Num_Iter&".png", True
			imagenToWord "Error SAP", RutaEvidencias() &"Error_SAP_"&Num_Iter&".png"
		   	JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		   	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para LUDER").JavaButton("Cerrar").Click
			wait 1
			ExitTestIteration
		End If
End If
	
End Sub
Sub SeleccionarPlanTarifario()

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 70000
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaList("ComboBoxNative$1").Select "Planes Móviles"
	wait 3
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 40000
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaEdit("Equipo seleccionado:").Set DataTable("e_TipoDePlan", dtLocalSheet)
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaButton("Buscar").Click
	wait 1

		While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaCheckBox("Seleccionar").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist))   = False
			wait 1
		Wend
		
	wait 1
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
		wait 1
		varsap=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
		DataTable("s_Resultado",dtLocalSheet)="Fallido"
		DataTable("s_Detalle",dtLocalSheet)=varsap
		Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorPlanTarifario"&".png", True
		imagenToWord "ErrorPlanTarifario", RutaEvidencias() &Num_Iter&"_"&"ErrorPlanTarifario"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaButton("Cerrar").Click
		wait 1
		ExitActionIteration
	End If
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaCheckBox("Seleccionar").Set "ON"
	wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Plan Tarifario"&".png", True
	imagenToWord "Plan Tarifario", RutaEvidencias() &Num_Iter&"_"&"Plan Tarifario"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaButton("Siguiente >").Click
	
		tiempo = 0
			Do
			tiempo = tiempo + 1
				If tiempo>=120 Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No cargo la pantalla 'Actualizar Atributos'"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
				else
				Reporter.ReportEvent micPass,"OK","Continuar Flujo"
				End If
			Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Exist(1)

End Sub
Sub ParametrosAlta()
	
	wait 1
	Dim Iterator
	Count = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").GetROProperty ("items count")
	For Iterator = 0 To Count-1
	 	rs = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").GetItem (Iterator)
		If rs = str_tipoalta Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Select str_tipoalta
			Exit for
		ElseIf Iterator = Count-1 Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Select "Pedido de Cliente"
			Exit for
		End if	
	Next
		
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Texto del motivo:").Set str_motivo_alta
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Código de Centro Poblado").Set "1501010001"
	wait 1
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaCheckBox("Tiene cobertura").Set "OFF"
'	wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Atributos del Producto"&".png", True
	imagenToWord "Atributos del Producto", RutaEvidencias() &Num_Iter&"_"&"Atributos del Producto"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Siguiente >").Click

		While((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Exist) Or(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTab("JXTabbedPane").Exist))=False
			wait 1
		Wend

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Mensaje de Validación"&".png", True
		imagenToWord "Mensaje de Validación", RutaEvidencias() &Num_Iter&"_"&"Mensaje de Validación"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
	End If
	
End Sub
Sub Portabilidad()

	Call RecursosAlta()
	If DataTable("e_RolFamilia",dtLocalSheet)<>Empty Then
		Call SelecionRol()
	End If
	Call Financiamiento()
	Call Permanencia()
	
	wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Configuración del Producto"&".png", True
	imagenToWord "Configuración del Producto", RutaEvidencias() &Num_Iter&"_"&"Configuración del Producto"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
	wait 1

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Mensajesdevalidación"&".png", True
		imagenToWord "Mensajes de validación", RutaEvidencias() &Num_Iter&"_"&"Mensajesdevalidación"&".png"
		wait 1
		varpag=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		DataTable("s_Detalle", dtLocalSheet) = varpag
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
		varpag = Mid(varpag,1,47 )
			
		If  varpag="La regla rule with the following details failed" Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Click
		End If

		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
			wait 2
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
			wait 2
			ExitActionIteration
		End If
	End If
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Click

		tiempo = 0
		Do
		tiempo = tiempo + 1
					If tiempo>=80 Then
							DataTable("s_Resultado", dtLocalSheet) = "Fallido"
							DataTable("s_Detalle", dtLocalSheet) = "No cargo la pantalla 'Negociar dirección'"
							Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
							ExitActionIteration
					else
							Reporter.ReportEvent micPass,"OK","Continuar Flujo"
					End If
		wait 1
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("En tienda").Exist) Or(JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist))
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Mensajesdevalidación"&".png", True
		imagenToWord "Mensajes de validación", RutaEvidencias() &Num_Iter&"_"&"Mensajesdevalidación"&".png"
		wait 1
		varpag=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		DataTable("s_Detalle", dtLocalSheet) = varpag
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
		wait 2
		ExitActionIteration
	End If
End Sub
Sub RecursosAlta()
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Exist Then
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Set
		wait 2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Parametriza_Producto_"&Num_Iter&".png", True
		imagenToWord "Seleccionar Contacto", RutaEvidencias() &Num_Iter&"_"&"Parametriza_Producto_"&Num_Iter&".png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaButton("Siguiente >").Click
	End If

		Dim tiempo
		tiempo = 0
		Do
		tiempo = tiempo + 1
			If tiempo>=80 Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No cargo la pantalla 'Negociar Configuración'"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
			else
					Reporter.ReportEvent micPass,"OK","Continuar Flujo"
			End If
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTab("JXTabbedPane").Exist(2)
 
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("Buscar por nombre:").Set "Tipo de SIM"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Buscar").Click
	wait 1
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
	End If

	Select Case DataTable("e_TipodeAlta", dtLocalSheet)
		Case "Alta Nueva Equipo + Linea"
			'IngresoTipoSIM
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
		
		Case "Alta Nueva Solo Linea"
			'IngresoTipoSIM
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
			
			'IngresoGrupoSIM
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
			
			'IngresoIMEI
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
	End Select
	
End  Sub
Sub SelecionRol()

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Mostrar atributos:").Select "Todo"
	wait 3
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("Buscar por nombre:").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("Buscar por nombre:").Set "Rol"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Buscar").Click
	wait 3

	filas=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetROProperty("rows")
		For Iterator = filas-1 To 0 Step -1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow (Iterator)
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").PressKey "C", micCtrl
			JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").PressKey "V", micCtrl
			str_titulo = JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").GetROProperty("text")
			str_titulo = replace(str_titulo,"Nombre    Valor    Acción    Por única vez    Mensual     ","")
			str_titulo = replace(str_titulo,"0,00    0,00    ","")
			str_titulo = Left(str_titulo,3) 
				
				If str_titulo = "Rol" Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT" 
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", DataTable("e_RolFamilia", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Rol"&".png", True
					imagenToWord "Seleccion de Rol", RutaEvidencias() &Num_Iter&"_"&"Rol"&".png"
					Exit For
					wait 1
				End If
		Next
		wait 1
End Sub
Sub Financiamiento()

	IF ucase(str_Financiamiento) = "SI"  Then

	Select Case str_Cuotas
			Case 18
			 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "MOVISTAR-18 cuotas"
			     Set shell = CreateObject("Wscript.Shell")
			     shell.SendKeys "{RIGHT 100}"
			     JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar_2").Click
			     While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
					wait 1
					Wend
				     wait 1
		             JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Calcular").Click
		            While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
						wait 1
					Wend
			Case 12
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "MOVISTAR-12 cuotas"
			     Set shell = CreateObject("Wscript.Shell")
			     shell.SendKeys "{RIGHT 100}"
			     JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar_2").Click
			     While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
					wait 1
				Wend
			     wait 1
	             JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Calcular").Click
	          	While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
					wait 1
				Wend
		End Select
		
    	If JavaWindow("Ejecutivo de interacción").JavaDialog("JDialog").JavaButton("Aceptar").Exist Then
			varfin=JavaWindow("Ejecutivo de interacción").JavaDialog("JDialog").JavaObject("JPanel").GetROProperty("text")
			varfin="El Plan de Financiamiento seleccionado no está disponible para el cliente. Seleccione otro Plan de Financiamiento o continúe con el Plan de Financiamiento predeterminado."
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"FinancimientoNoDisponible"&".png", True
			imagenToWord "Financimiento No Disponible", RutaEvidencias() &Num_Iter&"_"&"FinancimientoNoDisponible"&".png"
			
			JavaWindow("Ejecutivo de interacción").JavaDialog("JDialog").JavaButton("Aceptar").Click
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar_2").Click
			While(JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Exist)=False
				wait 1
			Wend
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Click
			wait 1
			While(JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_3").JavaList("Motivo:").Exist)=False
				wait 1
			Wend
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_3").JavaList("Motivo:").Select "Pedido de Cliente"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_3").JavaButton("Aceptar").Click
			wait 5
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar_3").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar_3").Click
				wait 1
			End If
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
				wait 1
			End If
			ExitActionIteration
		End If
        wait 1
	     JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Financiamiento18cuotas.png", True
         imagenToWord "Financiamiento str_Cuotas cuotas",RutaEvidencias() & "Financiamiento18cuotas.png"
	     
	ElseIf ucase(str_Financiamiento) = "NO" Then 
	     JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "Contado"
	     Set shell = CreateObject("Wscript.Shell")
		 shell.SendKeys "{RIGHT 100}"
		 JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar_2").Click
		While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
			wait 1
		Wend
		 wait 1
         JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Calcular").Click
	   While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
			wait 1
		Wend
		 wait 1
		 JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "FinanciamientoContado.png", True
	     imagenToWord "Financiamiento contado",RutaEvidencias() & "FinanciamientoContado.png"
	End If
	
	
End Sub
Sub Permanencia()

	If str_Permanencia="SI" Then
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Mostrar atributos:").Select "Todo"
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("Buscar por nombre:").Set "Período de compromiso del Equipo"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Buscar").Click
		wait 3
		filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:_2").GetROProperty("rows")
		For Iterator = filas-1 To 0 step -1
		    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:_2").SelectRow ("#"&Iterator)
			j = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:_2").GetCellData("#"&Iterator, "#1")
		    h = Instr(1,j,"Período de compromiso del Equipo")
		    If h <> 0 Then
		    	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:_2").DoubleClickCell "#"&Iterator, "#1", "LEFT" 
		    	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:_2").SetCellData "#"&Iterator, "#1", str_Periodo
		    	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "PeriodoPermanencia.png", True
				imagenToWord "Periodo Permanencia",RutaEvidencias() & "PeriodoPermanencia.png"
		    	Exit for 
		    End If
		Next 
		wait 1
	End If
End  Sub
Sub TipoEnvio()
	
	Select Case DataTable("e_MetodoEntrega", dtLocalsheet)
		Case "Delivery"'
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("Delivery").Set "ON"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png" , True
				imagenToWord "Canal de Venta", RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png"
				wait 1
					While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Lookup-Validated").Exist) = False
							wait 1
					Wend
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Lookup-Validated").Click
				wait 1
					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Exist) = False
						wait 1
					Wend
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaTable("SearchJTable").SelectRow "#0"
				wait 1
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DireccionEnvio"&".png" , True
				imagenToWord "Dirección de Envio", RutaEvidencias() &Num_Iter&"_"&"DireccionEnvio"&".png"
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Click
				wait 1
					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Exist) = False
						wait 1
					Wend
				wait 2	
				Set shell = CreateObject("Wscript.Shell") 
	            shell.SendKeys "{PGDN}"
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Set "PRUEBAS QA"
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Número de teléfono del").Set "987654321"
				wait 2
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"EntregaDelivery"&".png", True
				imagenToWord "Entrega Delivery", RutaEvidencias() &Num_Iter&"_"&"EntregaDelivery"&".png"
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Siguiente >").Click

'				wait 2
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("Delivery").Set "ON"
'				wait 2
'				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "MetodoEntrega_Delivery_"&Num_Iter&".png", True
'					While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Lookup-Validated").Exist) = False
'							wait 1
'					Wend
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Set "QA"
'				wait 2
'				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png", True
'				imagenToWord "Canal de Venta", RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png"
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Siguiente >").Click
'				wait 2
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("ID del Acuerdo de Facturación:").WaitProperty "editable", 1, 10000
			
'					While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").Exist)=False
'						wait 1
'					Wend
'
'					Do While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Lookup-Validated").Exist) = False
'							wait 1	
'							Dim c
'							c=c+1
'								If (c=8) Then exit do		
'					Loop 
'					wait 3
'			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Lookup-Validated").Exist(5) Then
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Lookup-Validated").Click
'				wait 2
'				
'					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").Exist)= False
'						wait 1
'					Wend
'					
'				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").Exist Then
'					fila=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaTable("SearchJTable").GetROProperty("rows")
'					fila= CInt(fila)
'					If fila>0 Then
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaTable("SearchJTable").SelectRow "#0"
'						wait 2
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaButton("Seleccionar").Click
'						wait 3
'					End If
'				End If
'				
'			End If
'			wait 2
'			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Negociar Distribución"&".png", True
'			imagenToWord "Negociar Distribución", RutaEvidencias() &Num_Iter&"_"&"Negociar Distribución"&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Siguiente >").Click
'				
		Case "En Tienda"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png", True
				imagenToWord "Canal de Venta", RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Siguiente >").Click
				wait 1
				Call Porta()
				wait 1
				
		End Select
		
		
		Do While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Lookup-Validated").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaButton("Pago inmediato").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist)) = False
				wait 1	
				Dim d
				d=d+1
				If (d=8) Then exit do		
		Loop 
		wait 1
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(1) Then
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No se puede seleccionr método de entrega"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
			ExitActionIteration
		End If
		wait 1
			
End Sub
Sub Porta()

	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Formulario de portabilidad").JavaEdit("Nombre del cliente:").Exist = False
		wait 1
	Wend
	Dim t
	t = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Formulario de portabilidad").JavaEdit("Nombre del cliente:").GetROProperty("text")
	While t = ""
		wait 1
		t = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Formulario de portabilidad").JavaEdit("Nombre del cliente:").GetROProperty("text")
	Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Formulario de portabilidad").JavaList("Tipo de Cliente:").Select "Uso Interno"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Formulario de portabilidad").JavaList("Modalidad ID en TEF:").Select "Postpago"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Formulario de portabilidad").JavaList("Tipo de documento:").Select "C"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Formulario de portabilidad").JavaList("Tipo de servicio:").Select "Móvil"
	wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"FormularioPorta.png" , True
	imagenToWord "Formulario de portabilidad", RutaEvidencias() &"FormularioPorta.png"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Formulario de portabilidad").JavaButton("Siguiente >").Click
	
End Sub
Sub NegociarDistribucion()
	
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").Exist Then
		
			While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").Exist = False
				wait 1
			Wend
			Dim f
			f = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").GetROProperty("text")
			While f = ""
				wait 1
				f = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").GetROProperty("text")
			Wend
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaRadioButton("Nuevo").Set "ON"
			wait 2
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Negociar Distribución"&".png", True
			imagenToWord "Negociar Distribución", RutaEvidencias() &Num_Iter&"_"&"Negociar Distribución"&".png"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Siguiente >").Click
			
		End If
			
			tiempo = 0
			Do
			tiempo = tiempo + 1
				If tiempo>=120 Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No cargo la pantalla 'Negociar Pago'"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
				else
					Reporter.ReportEvent micPass,"OK","Continuar Flujo"
				End If
			wait 1
			Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaButton("Pago inmediato").Exist(2)) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist) Or(JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist))   
			
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(1) Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Mensajesdevalidación"&".png", True
				imagenToWord "Mensajes de validación", RutaEvidencias() &Num_Iter&"_"&"Mensajesdevalidación"&".png"
				wait 1
				varpag=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = varpag
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Cerrar").Click
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
				wait 2
				ExitActionIteration
			End If
	
End Sub
Sub NegociarPago()
	
	If str_Financiamiento = "SI" Then

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
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaEdit("Importe de Cuota Inicial:").Set 100
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaEdit("Importe de Cuota Mayor:").Set 144
			wait 1
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "SolicitudFinanciamiento.png", True
			imagenToWord "Solicitud Financiamiento",RutaEvidencias() & "SolicitudFinanciamiento.png"
		Else  
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
				varmsj=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
				
				If varmsj="Por favor, llene todos los campos antes de clacular el Límite de Crédito del Cliente" Then
					JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
				End If
				If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
				End If
			End If
				getenable=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaCheckBox("Financiamiento Externo").GetROProperty("enabled")
				If getenable=1 Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaCheckBox("Financiamiento Externo").Set "OFF"
				End If
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
		Exit Sub
	End If
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Select "Factura"
	wait 1
'	var8 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Numero RUC").GetROProperty("text")
'	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Numero RUC").GetROProperty("text") = "" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Select "Boleta"
'	End If
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Select str_MedioPago
	If str_MedioPago = "Pago a la Factura" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Cantidad de cuotas:").Select "1"
		wait 1
	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Calcular").Click
	End If
	
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
	
		t = 0 
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist) = False
			Wait 1
			
			t = t + 1
			If (t >= 180) Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorRadiobuttonEnTienda_"&Num_Iter&".png", True
				imagenToWord "Error RadioButton En Tienda_"&Num_Iter,RutaEvidencias() & "ErrorRadiobuttonEnTienda_"&Num_Iter&".png"
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No se habilitó la opción -En tienda- de la siguiente pantalla"
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
				RunAction "WIC2", oneIteration
				Exit Do
			End If
		Wend
		
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

	'Bucle que espera el envío de la orden
	t = 0
	While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist)) = False
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
	
	'Control de Mensajde de Validación
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
		varlog = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable_3").GetCellData(0,1)
			DataTable("s_Resultado", dtLocalSheet) = "Mensaje de Validación"
			DataTable("s_Detalle", dtLocalSheet) = varlog
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "MensajeValidacion.png", True
			imagenToWord "Mensaje de Validación",RutaEvidencias() & "MensajeValidacion.png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Aceptar").Click
	End If
	text = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").GetROProperty("text")
    DataTable("s_Nro_Orden", dtLocalSheet) = RTRIM(LTRIM((replace(text,"Orden",""))))
    WAIT 1
    
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").CaptureBitmap RutaEvidencias() & "Orden_Generada_.png", True
    imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_.png"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Click
    
	DataTable("s_Resultado", dtLocalSheet) = "Éxito"
	DataTable("s_Detalle", dtLocalSheet) = "Se envió la orden "&str_NroOrden&" correctamente"
		
	If DataTable("e_MetodoEntrega", dtLocalSheet)="Delivery" Then
		ExitActionIteration
	End If	
			
			
End Sub
Sub PagoManual()

	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select
	wait 1
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Exist) = False
			wait  1
		Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Select "Tareas pendientes del equipo"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").SetFocus
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden",dtLocalSheet)
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
	wait 1
	
	tiempo=0
		Do 
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
				nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("-- Registros").GetROProperty("attached text")
				tiempo=tiempo+1
				wait 1
			End If
			If (tiempo >= 30) Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar el Pago de la Orden"
					Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
					ExitActionIteration
			End If
		Loop While Not(nroreg="1 Registros")
		wait 3
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").SelectRow "#0"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Gestión manual").Click
		
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Enviar").Exist) = False
			wait  1
		Wend
		
	tiempo=0
	Do
		var = JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Enviar").GetROProperty("enabled")
		tiempo=tiempo+1
			If (tiempo >= 8) Then
				DataTable("s_Resultado", dtLocalSheet) = "Exito"
	  			DataTable("s_Detalle", dtLocalSheet) = "La orden: "&DataTable("s_Nro_Orden", dtLocalSheet)&" ya fue pagado"
				Reporter.ReportEvent micPass, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Cancelar").Click
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
				wait 1
				Exit Do
			End If
		Loop While Not (var <> "0")
		wait 1
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Enviar").Exist Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Pago Manual"&".png", True
		imagenToWord "Pago Manual", RutaEvidencias() &Num_Iter&"_"&"Pago Manual"&".png"
		var3 = JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Enviar").GetROProperty("enabled")
		If var3 = "0" Then
			 JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Cancelar").Click 
			 else
			 JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Enviar").Click
			 Reporter.ReportEvent micDone, "Pago Correcto", "El número de orden : "&DataTable("s_Nro_Orden", dtLocalSheet)&" fué correctamente pagado"
		End If
		
	End If
End Sub
Sub GestionLogistica()
	
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Exist)=False
			wait 1
		Wend
	
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
			End If
			If (tiempo >= 30) Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar el Empuje de la Orden"
					Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
					ExitActionIteration
			End If
		Loop While Not (nroreg="1 Registros")
		wait 1
	
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
					While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").Exist) = False
						wait 1
					Wend
					
						If DataTable("e_TipodeAlta", dtLocalSheet)="Alta Nueva Equipo + Linea" Then
'								vardisp=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").GetCellData (1,1)
'								If vardisp="Dispositivo" Then
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#1","#4"
									Set shell = CreateObject("Wscript.Shell") 
									shell.SendKeys "{ENTER}"
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#1","#4",DataTable("e_ID_Dispositivo", dtLocalSheet)
									wait 2
'								Else 
'								 JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#1","#4"
'									Set shell = CreateObject("Wscript.Shell") 
'									shell.SendKeys "{ENTER}"
'									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#1","#4",DataTable("e_ID_SIM", dtLocalSheet)
'									wait 2
'								End If
					
'								varsim=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").GetCellData (2,1)
'								If varsim="Tarjeta SIM" Then
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#2","#4"
									Set shell = CreateObject("Wscript.Shell") 
									shell.SendKeys "{ENTER}"
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#2","#4",DataTable("e_ID_SIM", dtLocalSheet)
									wait 2
									
'								Else
'									
'									 JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#2","#4"
'									Set shell = CreateObject("Wscript.Shell") 
'									shell.SendKeys "{ENTER}"
'									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#2","#4",DataTable("e_ID_Dispositivo", dtLocalSheet)
'								End If
						else
'								varsim=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").GetCellData(1,4)
'								If varsim<>str_idSim Then
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#1","#4"
									Set shell = CreateObject("Wscript.Shell") 
									shell.SendKeys "{ENTER}"
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#1","#4",DataTable("e_ID_SIM", dtLocalSheet)
									wait 2
'								End If
						End If
						
					JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Ingreso_Materiales_"&".png", True
					imagenToWord "Ingreso de Materiales", RutaEvidencias() &Num_Iter&"_"&"Ingreso_Materiales_"&".png"
					JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Validar y Crear Factura").Object.doClick()
					wait 1
					tiempo = 0
					Do
						tiempo=tiempo+1
							varhab=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").GetROProperty("enabled")					
							wait 3
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
				       		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Error Logística"&".png" , True
							imagenToWord "Error Logística", RutaEvidencias() &Num_Iter&"_"&"Error Logística"&".png"
				     		
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
									If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist Then
									JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
								End If
								If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist(2) Then
									JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
									ExitActionIteration
								End If
				     		End  If
				    	End If
				End  If
				If tiempo>=20 Then
					Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)  
					DataTable("s_Resultado",dtLocalSheet) = "Fallido"
					DataTable("s_Detalle",dtLocalSheet) = "Luego de 20 intentos no se pudo realizar la Asignación de Series"
					ExitActionIteration
				else
					Reporter.ReportEvent micPass, "Exito", "Se realizo la Asignación de Series correctamente"
			End If
		Loop While Not varhab = "1"
		wait 2

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").Click
	End If
	
End Sub
Sub OrdenCerrado()
	
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Exist) = False
			wait 1
		Wend	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaCheckBox("Solo órdenes pendientes").Set "OFF"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
	wait 8
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:")
	Reporter.ReportEvent micPass,"Se valida el estado de la orden", DataTable("s_ValEstadoOrden", dtLocalSheet)
	
		tiempo = 0
		Do
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
		tiempo = tiempo +1
		wait 3
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:")
			If tiempo>=50 Then		
				DataTable("s_Resultado",dtLocalSheet) = "Fallido" 
				DataTable("s_Detalle", dtLocalSheet) = "La Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)&" no culmino en estado Cerrado"
				Reporter.ReportEvent micFail,DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				If DataTable("s_ValEstadoOrden", dtLocalSheet) = "Enviado" Then
					Exit Do
					wait 1
				End If	
				'ExitActionIteration
			else
				Reporter.ReportEvent micPass, "Se valida el estado de la orden", DataTable("s_ValEstadoOrden", dtLocalSheet)
			End If
		wait 1
		Loop While Not DataTable("s_ValEstadoOrden", dtLocalSheet) = "Enviado"
		DataTable("s_Resultado", dtLocalSheet) = "Éxito"
		DataTable("s_Detalle", dtLocalSheet) = "La orden culminó correctamente"
		Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Cerrada_"&".png", True
		imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"Orden_Cerrada_"&".png"

End Sub
Sub DetalleActividadOrden()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaList("Ver por:").Select "Acciones de orden"
	wait 3
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").DoubleClickCell 0, "#8", "LEFT"
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{ENTER}"
	wait 1

		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 713979A").JavaEdit("Fecha de vencimiento:").Exist)=False
			wait 1
		Wend
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 713979A").JavaTab("Nombre del cliente:").Select "Actividad"
	
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 713979A").JavaTable("SearchJTable").Exist)=False
			wait 1	
		Wend
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_1"&".png", True
	imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_1"&".png"
	
	shell.SendKeys "{PGDN}"
	wait 1
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_2"&".png", True
	imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_2"&".png"
	
	dim Iterator , filas	
	filas=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 713979A").JavaTable("SearchJTable").GetROProperty("rows")
	For Iterator = filas-1 to 0 step -1
		varselec=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 713979A").JavaTable("SearchJTable").GetCellData(Iterator,0)
		If varselec="Manejar número de portabilidad" Then
			DataTable("s_Resultado",dtLocalSheet)="Exitoso"
			DataTable("s_Detalle",dtLocalSheet)="La orden "&DataTable("s_Nro_Orden",dtLocalSheet)&" culmino en estado Cerrado, exitoso en la Actividad "&varselec&""
			Reporter.ReportEvent micPass, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet) 	   			    	
		     Exit for 	    
		     Else 
		     	DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)="La orden "&DataTable("s_Nro_Orden",dtLocalSheet)&" no culmino en estado Cerrado, falló en la Actividad "&varselec&""
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 713979A").JavaButton("Cancelar").Click
		
				wait 2
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
					wait 2
				End If
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(2) Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
				End If
					ExitActionIteration
				Exit for
		End If	
	Next
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 713979A").JavaButton("Cancelar").Click

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
		wait 2
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
	End If
	
End Sub
Sub ValidacionPortaflow()
		RunAction "Consulta Previa", oneIteration
End Sub






