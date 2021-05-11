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
'Call TipoEnvio()
'Call Financiamiento()
'Call GeneracionOrden()
'If DataTable("e_Ambiente", "Login [Login]")<>"PROD" Then
'	Call PagoManual()
'End If
'Call GestionLogistica()
'If DataTable("e_Ambiente", "Login [Login]")<>"PROD" Then
'	Call EmpujeOrden()
'End If
'Call OrdenCerrado()
'Call DetalleActividadOrden()

Sub SeleccionarTipoAlta()

wait 10
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
			JavaWindow("Ejecutivo de interacción").JavaTable("Titulo").ActivateRow "#5" @@ hightlight id_;_27080509_;_script infofile_;_ZIP::ssf1.xml_;_
			wait 1
	End	Select
	
End Sub
Sub FlujoWIC()

	If DataTable("e_WIC_ValidaCli", dtLocalsheet)="SI" Then
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
		wait 3
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist(2) Then
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
		wait 2
		
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
			wait 2
			ExitTestIteration
		End If
End If
	
End Sub
Sub SeleccionarPlanTarifario()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 70000
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaList("ComboBoxNative$1").Select "Planes Móviles"
	wait 5
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 40000
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
	wait 5
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaCheckBox("Seleccionar").Set "ON" @@ hightlight id_;_22129898_;_script infofile_;_ZIP::ssf2.xml_;_
	wait 2
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Plan Tarifario"&".png", True
	imagenToWord "Plan Tarifario", RutaEvidencias() &Num_Iter&"_"&"Plan Tarifario"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaButton("Siguiente >").Click
	
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
	'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Select str_tipoalta
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

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Mensaje de Validación"&".png", True
		imagenToWord "Mensaje de Validación", RutaEvidencias() &Num_Iter&"_"&"Mensaje de Validación"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
	End If
	
End Sub
Sub RecursosAlta()
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Exist(4) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Set @@ hightlight id_;_2822126_;_script infofile_;_ZIP::ssf4.xml_;_
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
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTab("JXTabbedPane").Exist(2) @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf19.xml_;_
 
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTab("JXTabbedPane").Select "Asignación de número"
	wait 3
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("TextFieldNative$1").Type "6%%%%%%%%"	
	'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("TextFieldNative$1").Type "92095%%%%"
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
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist(2) Then
		DataTable("s_Resultado",dtLocalSheet) = "Fallido" 
		DataTable("s_Detalle", dtLocalSheet) = "No hay ningún número devuelto"
	    Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
	    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"No hay números disponibles"&".png", True
		imagenToWord "No hay ningún número disponible", RutaEvidencias() &Num_Iter&"_"&"No hay números disponibles"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
		wait 2
		ExitTestIteration
	End If
	'wait 3
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Distribuir Número").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Distribuir Número").Click
		wait 1
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Asignar número").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Asignar número").Click
	End If
	
	wait 3
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
		wait 1
	End If
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

	filas=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetROProperty("rows")
	For Iterator = filas-1 To 0 Step -1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow "#"&Iterator
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").PressKey "C",micCtrl
			JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").PressKey "V",micCtrl
			str_titulo=JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").GetROProperty("text")
			str_titulo = Replace(str_titulo,"Nombre    Valor    Por única vez    Mensual     ","")
			If str_titulo="Grupo de SIM    NA            " Then
				str_titulo = Left(str_titulo,12)
				else
				str_titulo = Left(str_titulo,11)
			End If
			wait 1
					Select Case DataTable("e_TipodeAlta", dtLocalSheet)
						Case "Alta Nueva Equipo + Linea"
							If str_titulo="Tipo de SIM" Then
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT"
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", DataTable("e_TipoSIM", dtLocalSheet) 
								JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Mostrar_Atributos_"&Num_Iter&".png", True
								wait 2
								Exit For
							End  If
							
						Case "Alta Nueva Solo Linea"
							If str_titulo="Tipo de SIM" Then
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT"
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", DataTable("e_TipoSIM", dtLocalSheet) 
								wait 2
							End  If
							If str_titulo="Grupo de SIM" Then
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT"
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", "Estandar"
								wait 2
							End  If
							If str_titulo="Número IMEI" Then
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT"
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", "811111111111111"
								wait 2
								Exit For
							End If
						End Select	
					wait 1
		Next
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Configuración del Producto"&".png", True
	imagenToWord "Configuración del Producto", RutaEvidencias() &Num_Iter&"_"&"Configuración del Producto"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
	wait 8
	
	If DataTable("e_RolFamilia",dtLocalSheet)<>Empty Then
		call SelecionRol()
	End If
	
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
	
	If DataTable("e_Financiamiento_Externo",dtLocalSheet)="SI" Then
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select DataTable("e_Plan_Financiamiento",dtLocalSheet)
        wait 1
        JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Calcular").Click
        wait 8
        JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Financiamiento"&".png", True
		imagenToWord "Monto calculado", RutaEvidencias() &Num_Iter&"_"&"Financiamiento"&".png"
	Else
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Calcular").Click
        wait 10
        JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Contado"&".png", True
		imagenToWord "Monto calculado", RutaEvidencias() &Num_Iter&"_"&"Contado"&".png"
		
	End If	
	
	If  DataTable("e_CR3988",dtLocalSheet)="SI" Then
		
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
		DataTable("s_Resultado", dtLocalSheet) = "éxitoso"
		DataTable("s_Detalle", dtLocalSheet) = varpag
		Reporter.ReportEvent micPass, DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		wait 2
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cancelar oferta").Click
		
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar una opción").Exist = False 
			wait 1
		Wend
		
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar una opción").JavaButton("Sí").Click
		wait 1
		
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración").Exist = False
		  wait 1
		Wend
		wait 1
		
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración").JavaButton("Aceptar").Click
		wait 1
		 While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").Exist = False
		 	wait 1
		 Wend
		 
		 wait 1
		 
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"ordenCancelada.png", True
		imagenToWord "Orden Cancelada", RutaEvidencias() &"ordenCancelada.png"
		wait 2
		ExitActionIteration
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("En tienda").Exist=True Then
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		DataTable("s_Detalle", dtLocalSheet) = "Flujo err´0neo"
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
		
		ExitActionIteration
	End If

End Sub















	
