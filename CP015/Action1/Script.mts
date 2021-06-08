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
str_idSim			=	DataTable("e_ID_SIM", dtLocalSheet)
str_idDispositivo	=	DataTable("e_ID_Dispositivo", dtLocalSheet)
Num_Iter 			= 	Environment.Value("ActionIteration")
 @@ hightlight id_;_31334378_;_script infofile_;_ZIP::ssf8.xml_;_
Call SeleccionarTipoAlta()
Call FlujoWIC()
Call SeleccionarUbica()
Call SeleccionarEquipoMovil()
Call SeleccionarPlanTarifario()
Call ParametrosAlta()
Call RecursosAlta()
Call TipoEnvio()
Call Financiamiento()
Call GeneracionOrden()
If DataTable("e_Ambiente", "Login")<>"PROD" Then
	Call PagoManual()
End If
Call GestionLogistica()
If DataTable("e_Ambiente", "Login")<>"PROD" Then
	Call EmpujeOrden()
End If
Call OrdenCerrado()
Call DetalleActividadOrden()

Sub SeleccionarTipoAlta()
	Select Case DataTable("e_TipodeAlta", dtLocalsheet)
		Case "Alta Nueva Equipo + Linea"
			wait 3
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Panel_Interaccion_"&Num_Iter&".png", True
			imagenToWord "Panel de Interacción", RutaEvidencias() &"Panel_Interaccion_"&Num_Iter&".png"
			JavaWindow("Ejecutivo de interacción").JavaTable("Titulo").ActivateRow "#4"
			wait 1
			
		Case "Alta Nueva Solo Linea"
			wait 3
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Panel_Interaccion_"&Num_Iter&".png", True
			imagenToWord "Panel de Interacción", RutaEvidencias() &"Panel_Interaccion_"&Num_Iter&".png"
			JavaWindow("Ejecutivo de interacción").JavaTable("Titulo").ActivateRow "#5" @@ hightlight id_;_27080509_;_script infofile_;_ZIP::ssf1.xml_;_
			wait 1
	End	Select
	
End Sub
Sub FlujoWIC()
	While ((JavaWindow("Ejecutivo de interacción").JavaDialog("Autenticación del Cliente").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Ubicación").JavaList("Departamento:").Exist)) = False
		wait 1
	Wend
	wait 1
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Autenticación del Cliente").Exist Then
		RunAction "WIC", oneIteration
	End If

End Sub
Sub SeleccionarUbica()

			While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Ubicación").JavaList("Departamento:").Exist) = False
				wait 1	
			Wend
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Ubicación").JavaList("Departamento:").Select DataTable("e_Departamento", dtLocalSheet)
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Ubicación").JavaList("Provincia:").Select DataTable("e_Provincia", dtLocalSheet)
		wait 1
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "SeleccionarUbicacion_"&Num_Iter&".png", True
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Ubicación").JavaButton("Siguiente >").Click
			While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaList("ComboBoxNative$1").Exist) Or(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaList("ComboBoxNative$1").Exist)) = False
				wait 1	
			Wend
End Sub
Sub SeleccionarEquipoMovil()

	If (DataTable("e_TipodeAlta", dtLocalSheet) = "Alta Nueva Equipo + Linea") Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 30000
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaList("ComboBoxNative$1").Select "Celulares"
		wait 5
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaEdit("TextFieldNative$1").Set DataTable("e_ModeloCelular", dtLocalSheet)
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaButton("Buscar").Click
		wait 5
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").close
		End If
		wait 2
		
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
			Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaButton("Agregar al carrito").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist))
			wait 2
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
			Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaButton("Agregar al carrito_2").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist))
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
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para LUDER").JavaButton("Cerrar").Click
			wait 2
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
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para LUDER").JavaButton("Cerrar").Click
			wait 2
			ExitActionIteration
		End If

		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Equipo Móvil"&".png", True
		imagenToWord "Equipo Móvil", RutaEvidencias() &Num_Iter&"_"&"Equipo Móvil"&".png"
		wait 1
		
		If DataTable("e_ModeloCelular", dtLocalSheet)<>"HUAWEI P10 NEGRO" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaButton("Agregar al carrito").Click
		else
			'MsgBox "Seleccionar el equipo movil"	
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaButton("Agregar al carrito_2").Click
		End If
		
			While((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaButton("Buscar").Exist)) = False
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
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 20000
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaList("ComboBoxNative$1").Select DataTable("e_SubCategoria", dtLocalSheet)
	wait 5
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 70000
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaEdit("Equipo seleccionado:").Set DataTable("e_TipoDePlan", dtLocalSheet)
	wait 5
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaButton("Buscar").Click

		While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaCheckBox("Seleccionar").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist))   = False
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
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaButton("Cerrar").Click
		wait 1
		ExitActionIteration
	End If
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaCheckBox("Seleccionar").Set "ON" @@ hightlight id_;_22129898_;_script infofile_;_ZIP::ssf2.xml_;_
	wait 2
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Plan Tarifario"&".png", True
	imagenToWord "Plan Tarifario", RutaEvidencias() &Num_Iter&"_"&"Plan Tarifario"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaButton("Siguiente >").Click
	
		tiempo = 0
			Do
			tiempo = tiempo + 1
				If tiempo>=160 Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No cargo la pantalla 'Actualizar Atributos'"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
				else
				Reporter.ReportEvent micPass,"OK","Continuar Flujo"
				End If
			wait 1
			Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Exist

End Sub
Sub ParametrosAlta()
	

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
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaCheckBox("Tiene cobertura").Set "OFF" @@ hightlight id_;_13164893_;_script infofile_;_ZIP::ssf3.xml_;_
	wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Atributos del Producto"&".png", True
	imagenToWord "Atributos del Producto", RutaEvidencias() &Num_Iter&"_"&"Atributos del Producto"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Siguiente >").Click

		While((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Exist) Or(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTab("JXTabbedPane").Exist))=False
			wait 1
		Wend

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Mensaje de Validación"&".png", True
		imagenToWord "Mensaje de Validación", RutaEvidencias() &Num_Iter&"_"&"Mensaje de Validación"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
	End If
	
End Sub
Sub RecursosAlta()
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Set @@ hightlight id_;_2822126_;_script infofile_;_ZIP::ssf4.xml_;_
		wait 2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Parametriza_Producto_"&Num_Iter&".png", True
		imagenToWord "Seleccionar Contacto", RutaEvidencias() & "Parametriza_Producto_"&Num_Iter&".png"
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
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTab("JXTabbedPane").Exist @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf19.xml_;_
 
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTab("JXTabbedPane").Select "Asignación de número"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("TextFieldNative$1").Type "6%%%%%%%%"	
	'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("TextFieldNative$1").Type "920955%%%"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Proponer números").Click
	wait 2
	
		tiempo=0
			Do
				tiempo=tiempo+1
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Asignar número").Exist Then
					varasig=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Asignar número").GetROProperty("enabled")
				End If
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Distribuir Número").Exist Then
					varasig2=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Distribuir Número").GetROProperty("enabled")
				End If
				wait 2
		Loop  While Not ((varasig="1") Or (varasig2="1") Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist))
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist Then
		DataTable("s_Resultado",dtLocalSheet) = "Fallido" 
		DataTable("s_Detalle", dtLocalSheet) = "No hay ningún número devuelto"
	    Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
	    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"No hay números disponibles"&".png", True
		imagenToWord "No hay ningún número disponible", RutaEvidencias() &Num_Iter&"_"&"No hay números disponibles"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
		wait 1
		ExitTest
	End If
	'wait 3
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Distribuir Número").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Distribuir Número").Click
		wait 1
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Asignar número").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Asignar número").Click
	End If
	wait 1
'	If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist Then
'		JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
'		wait 1
'	End If
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("SearchJTable").Output CheckPoint("SearchJTable") @@ hightlight id_;_23150386_;_script infofile_;_ZIP::ssf16.xml_;_
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Número Asignado"&".png", True
	imagenToWord "Número Asignado", RutaEvidencias() &Num_Iter&"_"&"Número Asignado"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTab("JXTabbedPane").Select "Configuración"			
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("Buscar por nombre:").Set "Tipo de SIM"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Buscar").Click
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
	wait 8
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
			    	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData "#"&Iterator, "#1",DataTable("e_TipoSIM", dtLocalSheet)
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
			    	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData "#"&Iterator, "#1",DataTable("e_TipoSIM", dtLocalSheet) 
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


'	filas=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetROProperty("rows")
'	For Iterator = filas-1 To 0 Step -1
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow "#"&Iterator
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").PressKey "C",micCtrl
'			JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").PressKey "V",micCtrl
'			str_titulo=JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").GetROProperty("text")
'			str_titulo = Replace(str_titulo,"Nombre    Valor    Por única vez    Mensual     ","")
'			If str_titulo="Grupo de SIM    NA            " Then
'				str_titulo = Left(str_titulo,12)
'				else
'				str_titulo = Left(str_titulo,11)
'			End If
'			wait 1
'					Select Case DataTable("e_TipodeAlta", dtLocalSheet)
'						Case "Alta Nueva Equipo + Linea"
'							If str_titulo="Tipo de SIM" Then
'								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT"
'								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", DataTable("e_TipoSIM", dtLocalSheet) 
'								JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Mostrar_Atributos_"&Num_Iter&".png", True
'								wait 2
'								Exit For
'							End  If
'							
'						Case "Alta Nueva Solo Linea"
'							If str_titulo="Tipo de SIM" Then
'								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT"
'								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", DataTable("e_TipoSIM", dtLocalSheet) 
'								wait 2
'							End  If
'							If str_titulo="Grupo de SIM" Then
'								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT"
'								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", "Estandar"
'								wait 2
'							End  If
'							If str_titulo="Número IMEI" Then
'								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT"
'								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", "811111111111111"
'								wait 2
'								Exit For
'							End If
'						End Select	
'					wait 1
'		Next
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Configuración del Producto"&".png", True
	imagenToWord "Configuración del Producto", RutaEvidencias() &Num_Iter&"_"&"Configuración del Producto"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
	wait 8
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Mensajesdevalidación"&".png", True
		imagenToWord "Mensajes de validación", RutaEvidencias() &Num_Iter&"_"&"Mensajesdevalidación"&".png"
		wait 1
		varpag=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
		If varpag="<html>El tipo de SIM no puede ser NA -&#8203; se debe seleccionar un valor relevante. (Detectado en Tarjeta SIM).</html>" Then
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
			wait 1
			'IngresoTipoSIM
			filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetROProperty("rows")
			For Iterator = filas-1 To 0 step -1
			    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow ("#"&Iterator)
				j = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetCellData("#"&Iterator, "#1")
			    h = Instr(1,j,"Tipo de SIM")
			    If h <> 0 Then
			        JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell "#"&Iterator, "#1", "LEFT" 
			    	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData "#"&Iterator, "#1",DataTable("e_TipoSIM", dtLocalSheet)
			    	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "IngresoTipoSIM.png", True
					imagenToWord "Ingresamos Tipo SIM",RutaEvidencias() & "IngresoTipoSIM.png"
			    	Exit for 
			    End If
			Next
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
			wait 8			
		End If
	End If

	If DataTable("e_RolFamilia",dtLocalSheet)<>Empty Then
		SelecionRol()
	End If
	
	If DataTable("e_TipoServAdicional",dtLocalSheet)<>Empty Then
		ServiciosAdicionales()
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
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
	
	If DataTable("e_Financiamiento_Externo",dtLocalSheet)="SI" Then
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select DataTable("e_Plan_Financiamiento",dtLocalSheet)
        wait 1
        JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Calcular").Click
        wait 8
        JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Financiamiento"&".png", True
		imagenToWord "Monto calculado", RutaEvidencias() &Num_Iter&"_"&"Financiamiento"&".png"
		
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
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Mensajesdevalidación"&".png", True
		imagenToWord "Mensajes de validación", RutaEvidencias() &Num_Iter&"_"&"Mensajesdevalidación"&".png"
		wait 1
		varpag=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		DataTable("s_Detalle", dtLocalSheet) = varpag
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
		wait 1
		ExitActionIteration
	End If

End Sub
Sub SelecionRol()

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Mostrar atributos:").Select "Todo"
	wait 3
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("Buscar por nombre:").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("Buscar por nombre:").Set "Rol"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Buscar").Click
	wait 2

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
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
		wait 8
End Sub
Sub ServiciosAdicionales()
	
	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").GetROProperty("items count")
	For Iterator = 0 To filas-1
		nodeName=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").GetItem(Iterator)
		If nodeName = DataTable("e_TipoServAdicional", dtLocalSheet) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select(nodeName)
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Expand(nodeName)
			filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").GetROProperty("items count")
		 End If
	Next 
	
	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").GetROProperty("items count")
	For Iterator = 0 To filas-1	 
		 nodeName2nd=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").GetItem(Iterator)
		 If nodeName2nd=DataTable("e_ServAdicional", dtLocalSheet) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select(nodeName2nd)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "IngresoTipoSIM.png", True
			imagenToWord "Ingresamos Tipo SIM",RutaEvidencias() & "IngresoTipoSIM.png"
			Exit For
		End If
 	Next 
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
	wait 8
	
End Sub
Sub TipoEnvio()
	
	Select Case DataTable("e_MetodoEntrega", dtLocalsheet)
		Case "Delivery"
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
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Set "QA"
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Número de teléfono del").Set "1234"
				wait 1
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
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Siguiente >").Click @@ hightlight id_;_11418863_;_script infofile_;_ZIP::ssf5.xml_;_
'				wait 2
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("ID del Acuerdo de Facturación:").WaitProperty "editable", 1, 10000 @@ hightlight id_;_9978055_;_script infofile_;_ZIP::ssf28.xml_;_
			
					While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").Exist)=False
						wait 1
					Wend

					Do While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Lookup-Validated").Exist) = False
							wait 1	
							Dim c
							c=c+1
								If (c=8) Then exit do		
					Loop 
					wait 3
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Lookup-Validated").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Lookup-Validated").Click
				wait 2
				
					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").Exist)= False
						wait 1
					Wend
					
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").Exist Then
					fila=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaTable("SearchJTable").GetROProperty("rows")
					fila= CInt(fila)
					If fila>0 Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaTable("SearchJTable").SelectRow "#0"
						wait 2
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaButton("Seleccionar").Click
						wait 3
					End If
				End If
				
			End If
			wait 2
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Negociar Distribución"&".png", True
			imagenToWord "Negociar Distribución", RutaEvidencias() &Num_Iter&"_"&"Negociar Distribución"&".png"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Siguiente >").Click
'				
		Case "En Tienda"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png", True
				imagenToWord "Canal de Venta", RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Siguiente >").Click
				
					Do While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Lookup-Validated").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Pago inmediato").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist)) = False
							wait 1	
							Dim d
							d=d+1
								If (d=8) Then exit do		
					Loop 
					wait 3
					
				If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
						DataTable("s_Resultado", dtLocalSheet) = "Fallido"
						DataTable("s_Detalle", dtLocalSheet) = "No se puede seleccionr método de entrega"
						Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
						JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
						ExitActionIteration
				End If
				
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Lookup-Validated").WaitProperty "enabled", True, 20000
					wait 3
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaRadioButton("Nuevo").Set
					wait 2

	'				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Lookup-Validated").Exist Then
	'					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Lookup-Validated").Click
	'					wait 2
	'						While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").Exist)= False
	'							wait 1
	'						Wend
	'					wait 5
	'					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaButton("Seleccionar").Exist Then
	'						fila=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaTable("SearchJTable").GetROProperty("rows")
	'						fila= CInt(fila)
	'						If fila>0 Then
	'							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaTable("SearchJTable").SelectRow "#0"
	'							wait 2
	'							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaButton("Seleccionar").Click @@ hightlight id_;_20437236_;_script infofile_;_ZIP::ssf10.xml_;_
								wait 3
	'						End If
	'					End If
	'					
						JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Negociar Distribución"&".png", True
						imagenToWord "Negociar Distribución", RutaEvidencias() &Num_Iter&"_"&"Negociar Distribución"&".png"
						wait 3 
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Siguiente >").Click
	'				End If
						
				End If
				
		End Select
		
		wait 1
'		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Siguiente >").Exist Then
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Siguiente >").Click
'			wait 2
'		End If
		
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
			Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Pago inmediato").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist) Or(JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist))   
			
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
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
Sub Financiamiento()
	wait 3
	While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Pago inmediato").Exist)Or(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Límite de Compra").Exist)Or(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist))=False
		wait 1
	Wend
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Pago inmediato").Exist Then
		If DataTable("e_Financiamiento_Externo",dtLocalSheet)="SI" Then
			wait 2
			'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaCheckBox("Financiamiento Externo").Set "ON"
			
			While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaEdit("Importe de Cuota Mayor:").Exist)=False
				wait 1
			Wend
			wait 2
			'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaEdit("Importe de Cuota Inicial:").Set "0"
			'wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaEdit("Importe de Cuota Mayor:").Set "1"
			wait 2
			'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaList("Plan de Financiamiento:").Select "MOVISTAR-12 cuotas"
			'wait 2
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SolcFinanciamiento"&".png", True
			imagenToWord "Solicitud de Financiamiento", RutaEvidencias() &Num_Iter&"_"&"SolcFinanciamiento"&".png"
		End If
		wait 2
		var2=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Límite de Compra").GetROProperty ("enabled")
		If var2 = "1" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Límite de Compra").Click	
		End If
		
		
			tiempo=0
			Do
			tiempo=tiempo+1
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Pago inmediato").Exist Then
					varasig=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Pago inmediato").GetROProperty("enabled")
				End If
				wait 2
		Loop  While Not (varasig="1")
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Pago inmediato").Click
		
			While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist))=False
				wait 1
			Wend
	End If
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
			varsap=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
			DataTable("s_Resultado",dtLocalSheet)="Fallido"
			DataTable("s_Detalle",dtLocalSheet)=varsap
			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MensajeValidacción"&".png", True
			imagenToWord "Mensaje de Validacción", RutaEvidencias() &Num_Iter&"_"&"MensajeValidacción"&".png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"NegociarPago"&".png", True
			imagenToWord "Negociar Pago", RutaEvidencias() &Num_Iter&"_"&"NegociarPago"&".png"
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Siguiente >").Click
			
				While((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist)Or(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist))=False
					wait 1
				Wend
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Cerrar").Click
'			wait 2
'			JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
'			wait 2
'			ExitActionIteration
	End If

	wait 5
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Select DataTable("e_TipoPago",dtLocalSheet)
		wait 2
		'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Select DataTable("e_MetodoPago",dtLocalSheet)
		wait 2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Pago Inmediato"&".png", True
		imagenToWord "Pago Inmediato", RutaEvidencias() &Num_Iter&"_"&"Pago Inmediato"&".png"
		
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Enviar").Click
			While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").Exist)Or(JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist))=False
				wait 1
			Wend
		wait 1
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(1) Then
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Validacion"&".png", True
			imagenToWord "Mensaje de Validación", RutaEvidencias() &Num_Iter&"_"&"Validacion"&".png"
			wait 1
			varsap=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
			If varsap="El RUC es obligatorio para la Factura" Then
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Select "Boleta"
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Enviar").Click
				wait 3
			End If
		End If
			
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"NegociarPago"&".png", True
		imagenToWord "Negociar Pago", RutaEvidencias() &Num_Iter&"_"&"NegociarPago"&".png"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Siguiente >").Click
		
			While((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist)Or(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist))=False
				wait 1
			Wend
	
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
			varpag=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = varpag
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorPago"&".png", True
			imagenToWord "Error Pago", RutaEvidencias() &Num_Iter&"_"&"ErrorPago"&".png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Cerrar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
			wait 4
			ExitActionIteration
		End If
	End If
End Sub
Sub GeneracionOrden()
	

	Dim tiempo
	tiempo = 0
	Do
		'While((JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist)) = False
		While((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist)) = False
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Click
			If DataTable("e_WIC_ContrCli",dtLocalSheet)="SI" Then
					RunAction "WIC2", oneIteration
				Exit Do
			End If
			wait 3
		Wend
		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist Then
			wait 3
			var1 = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("attached text")
	   	 	JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
	   	 	wait 2
		End  If
		
'		If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist(2) Then
'			JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
'			wait 2
'		End If
		wait 1
			
			If tiempo>=180 Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No se a cargado el contrato correctamente"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			else
				Reporter.ReportEvent micPass,"Contrato Exitoso","Se a cargado el contrato correctamente"
			End If
	wait 2
	Loop While Not (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist or (var1 = "Contratos no Generados") or (var1 = "0"))
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Generación de Orden"&".png" , True
	imagenToWord "Generación de Orden", RutaEvidencias() &Num_Iter&"_"&"Generación de Orden"&".png"
	

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "GenerarContrato_"&Num_Iter&".png", True
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
	Wait 1
	
	'Click en "Enviar orden"
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Exist Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Envio de Orden"&".png" , True
		imagenToWord "Envio de Orden", RutaEvidencias() &Num_Iter&"_"&"Envio de Orden"&".png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
	End If
	
	'Bucle que espera el envío de la orden
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").Exist = False
		wait 1
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
		varvend=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
		If varvend="Por favor valide el Código de Ventas." Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click	
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validar").Click
			wait 5
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").JavaEdit("TextFieldNative$1").Set"4%%%%%%%"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").JavaTable("SearchJTable").SelectRow "#0"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").JavaButton("Seleccionar").Click

			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
			wait 2
		End If
	End  If
	Wend
	
	'Captura de la orden generada
	
	wait 2
	DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").GetROProperty("text")
	flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
	DataTable("s_Nro_Orden", dtLocalSheet) = replace (DataTable("s_Nro_Orden", dtLocalSheet),"Orden ","")
	Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png", True
	imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Click
	wait 1
	

			
	If DataTable("e_MetodoEntrega", dtLocalSheet)="Delivery" Then
		ExitActionIteration
	End If	
			
			
End Sub
Sub PagoManual()
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select @@ hightlight id_;_24061018_;_script infofile_;_ZIP::ssf2.xml_;_
	wait 1
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Exist) = False
			wait  1
		Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Select "Tareas pendientes del equipo" @@ hightlight id_;_29532040_;_script infofile_;_ZIP::ssf3.xml_;_
	wait 3
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
	wait 2
	
	tiempo=0
		Do 
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
				nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("1 Registros").GetROProperty("attached text")
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
		wait 1
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").SelectRow "#0"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Gestión manual").Click
		
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").Exist) = False
			wait  1
		Wend
		
	tiempo=0
	Do
		var = JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").GetROProperty("enabled")
		tiempo=tiempo+1
			If (tiempo >= 8) Then
				DataTable("s_Resultado", dtLocalSheet) = "Exito"
	  			DataTable("s_Detalle", dtLocalSheet) = "La orden: "&DataTable("s_Nro_Orden", dtLocalSheet)&" ya fue pagado"
				Reporter.ReportEvent micPass, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Cancelar").Click
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
				wait 2
				Exit Do
			End If
		Loop While Not (var <> "0")
		wait 1
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").Exist Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Pago Manual"&".png", True
		imagenToWord "Pago Manual", RutaEvidencias() &Num_Iter&"_"&"Pago Manual"&".png"
		var3 = JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").GetROProperty("enabled")
		If var3 = "0" Then
			 JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Cancelar").Click 
			 else
			 JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").Click
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
		wait 2
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").SelectRow "#0"
	
	tiempo=0
			Do
				If (DataTable("s_Detalle", dtLocalSheet)="Por favor rellenar todas las identificaciones de equipos") or (DataTable("s_Detalle", dtLocalSheet)<>"Por favor rellenar todas las identificaciones de equipos") Then
					If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Cancelar").Exist Then
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Cancelar").Click
						wait 2
					End If
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Gestionar logística").Click
					tiempo=tiempo+1
					wait 1
					While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").Exist) = False
						wait 1
					Wend
					Do
						vardisp=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").GetCellData (1,1)
						wait 3
						If vardisp = "Tarjeta SIM" Then
							Exit do 
						End If
					Loop While Not vardisp ="Dispositivo"
						If str_tipoalta="Alta Nueva Equipo + Linea" Then
								vardisp=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").GetCellData (1,4)
								If vardisp<>str_idDispositivo Then
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#1","#4"
									Set shell = CreateObject("Wscript.Shell") 
									shell.SendKeys "{ENTER}"
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#1","#4",DataTable("e_ID_Dispositivo", dtLocalSheet)
									wait 2
								End If
					
								varsim=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").GetCellData (2,4)
								If varsim<>str_idSim Then
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#2","#4"
									Set shell = CreateObject("Wscript.Shell") 
									shell.SendKeys "{ENTER}"
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#2","#4",DataTable("e_ID_SIM", dtLocalSheet)
									wait 2
								End If
						else
								varsim=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").GetCellData(1,4)
								If varsim<>str_idSim Then
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#1","#4"
									Set shell = CreateObject("Wscript.Shell") 
									shell.SendKeys "{ENTER}"
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#1","#4",DataTable("e_ID_SIM", dtLocalSheet)
									wait 2
								End If
						End If
						
					JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Ingreso_Materiales_"&".png", True
					imagenToWord "Ingreso de Materiales", RutaEvidencias() &Num_Iter&"_"&"Ingreso_Materiales_"&".png"
					JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Validar y Crear Factura").Object.doClick()
					
					tiempo = 0
					Do
						tiempo=tiempo+1
							varhab=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").GetROProperty("enabled")					
							wait 3
					Loop While Not((JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (varhab="1"))
					
				
						If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist(1) or JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
							If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist Then
								varlog = JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text") 
							End If
							If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
								varlog = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
							End If
							DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				       		DataTable("s_Detalle", dtLocalSheet) = varlog
				       		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
				       		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Error Logística"&".png" , True
							imagenToWord "Error Logística", RutaEvidencias() &Num_Iter&"_"&"Error Logística"&".png"
				     		
				     		If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").JavaButton("OK").Exist Then
				        		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").JavaButton("OK").Click
							End If
							If 	JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist Then
								JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
							End If
							wait 2
							If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Cancelar").Exist Then
								JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Cancelar").Click
								wait 2
							End If
				     		If DataTable("s_Detalle", dtLocalSheet)<>"Por favor rellenar todas las identificaciones de equipos" Then
									If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist Then
									JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
								End If
								If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
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

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").Click
	End If
	
End Sub
Sub EmpujeOrden()
	
		If DataTable("e_Tipo_De_DATA_Sim", dtLocalSheet) = "DATA LOGICA" Then

			JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
			JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select @@ hightlight id_;_19748072_;_script infofile_;_ZIP::ssf61.xml_;_
				While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Exist) = False
					wait 1
				Wend
			
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Select "Tareas pendientes del equipo" @@ hightlight id_;_25130440_;_script infofile_;_ZIP::ssf62.xml_;_
			wait 3
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").SetFocus
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
			wait 2
			
				tiempo=0
				Do 
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
						nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("1 Registros").GetROProperty("text")
						tiempo=tiempo+1
						wait 1
					End If
					
					If (tiempo >= 120) Then
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
			
'				tiempo=0
'				Do
'					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
'						wait 2
'						tiempo = tiempo+1
'						
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").Output CheckPoint("Equipo usuario:")
'						varValidaRespuestaCumplimiento = Environment("s_ValidaManejarRespuestaCumplimiento")
'						wait 1
'					End If
'						If (tiempo >= 120) Then
'							DataTable("s_Resultado",dtLocalSheet)="Fallido"
'							DataTable("s_Detalle",dtLocalSheet)="La actividad 'Manejar Respuesta de Cumplimiento' no cargo"	
'							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
'							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
'							ExitTestIteration
'						End If 
'				Loop While Not varValidaRespuestaCumplimiento = "Manejar Respuesta de Cumplimiento"
		
			wait 5
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").SelectRow "#0"
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Gestión manual").Click
				While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaList("Estado de la gestión manual:").Exist) = False
					wait  1
				Wend
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaList("Estado de la gestión manual:").Select "Cumplimiento Completo"
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaList("Motivo de la gestión manual").Select "Manejo manual: Manejo Manual OSS"
			wait 2
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Empuje Orden"&".png" , True
			imagenToWord "Empuje Orden", RutaEvidencias() &Num_Iter&"_"&"Empuje Orden"&".png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").Click
			wait 3
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
			wait 2
		End If
	
End Sub
Sub OrdenCerrado()
	
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select @@ hightlight id_;_17809817_;_script infofile_;_ZIP::ssf6.xml_;_
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
		Loop While Not DataTable("s_ValEstadoOrden", dtLocalSheet) = "Cerrado"
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

		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 776642A").JavaEdit("Fecha de vencimiento:").Exist)=False
			wait 1
		Wend
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 776642A").JavaTab("Nombre del cliente:").Select "Actividad"
	
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 776642A").JavaTable("SearchJTable").Exist)=False
			wait 1	
		Wend
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_1"&".png", True
	imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_1"&".png"
	
	shell.SendKeys "{PGDN}"
	wait 1
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_2"&".png", True
	imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_2"&".png"
	
	filas=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 776642A").JavaTable("SearchJTable").GetROProperty("rows")
	For Iterator = 0 To filas-1
		varselec=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 776642A").JavaTable("SearchJTable").GetCellData(Iterator,0)
	Next
	
	If varselec<>"Actualizar Descuento" Then
	 	DataTable("s_Resultado",dtLocalSheet)="Fallido"
		DataTable("s_Detalle",dtLocalSheet)="La orden "&DataTable("s_Nro_Orden",dtLocalSheet)&" culmino en la Actividad "&varselec&""
		Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 776642A").JavaButton("Cancelar").Click

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
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 776642A").JavaButton("Cancelar").Click

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
		wait 2
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
	End If
	
End Sub



