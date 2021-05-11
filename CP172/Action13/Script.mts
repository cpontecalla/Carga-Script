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
'If DataTable("e_Ambiente", "Login [Login]")<>"PROD" Then
	'Call PagoManual()
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
	
	'CARGOS AGREGADOS
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Expand "#0;Móvil;Servicios Adicionales"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Servicios Adicionales;BO Cargo Mensual Facturacion Detallada"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Servicios Adicionales;BO Cargo Mensual Servicio Integral de Emergencia"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Expand "#0;Móvil;Datos"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Expand "#0;Móvil;Datos;Paquete de Datos"
    wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;YouTube Ilim x 30 dias x S/25"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;YouTube Ilim x 7 dias x S/10"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;YouTube Ilim x 1 dia x S/5"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;YouTube Ilim x 15 dias x S/15"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Paquete 4G Ilim x 6 meses x S/30"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Paquete Compartible 1 GB"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Paq internet Ilimitado 4G x 30 dias x S/30"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Paq internet Ilimitado 4G x 15 dias x S/30"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Movistar Play Ilim x 1 dia x S/ 5"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Instagram Ilim x 7 dias x S/ 5"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Instagram Ilim x 1 dias x S/ 1"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Instagram Ilim x 15 dias x S/ 10"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Internet Negocios 4000"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Internet Mail"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Internet Total"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Paquete Total 30 Días 1GB"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Datos Negocios 1GB"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;800MB x 15 dias x S/15"    
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Full Internet 250MB(*)"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Paquete 2Gb Internet por 30 dias"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Paquete 1Gb Internet por 15 dias"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Paquete 500Mb Internet por 15 dias"    
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Paquete 400Mb Internet por 7 dias"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
    wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Datos;Paquete de Datos;Paquete 150Mb Internet por 7 dias"        
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click    
	wait 1
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Mostrar atributos:").Select "Todo"
	wait 1
	
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"Resumen2.png", True
	imagenToWord "Resumen de Facturación", RutaEvidencias() &"Resumen2.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
	wait 8
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Calcular").Click

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
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		DataTable("s_Detalle", dtLocalSheet) = varpag
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
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

	End If

End Sub
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
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Siguiente >").Click @@ hightlight id_;_11418863_;_script infofile_;_ZIP::ssf5.xml_;_
'				wait 2
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("ID del Acuerdo de Facturación:").WaitProperty "editable", 1, 10000 @@ hightlight id_;_9978055_;_script infofile_;_ZIP::ssf28.xml_;_
			
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
				
					Do While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Lookup-Validated").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Pago inmediato").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist)) = False
							wait 1	
							Dim d
							d=d+1
								If (d=8) Then exit do		
					Loop 
					wait 3
					
				If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(1) Then
						DataTable("s_Resultado", dtLocalSheet) = "Fallido"
						DataTable("s_Detalle", dtLocalSheet) = "No se puede seleccionr método de entrega"
						Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
						JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
						ExitActionIteration
				End If
				
				
					While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").Exist = False
						wait 1
					Wend
					Dim text
					text=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").GetROProperty("text")
					While text=""
						text=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").GetROProperty("text")
					Wend
					
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Lookup-Validated").WaitProperty "enabled", True, 20000
					wait 3
					
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaRadioButton("Nuevo").Set "ON"
					wait 2
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaList("Mostrar:").Select "Acciones de orden activas "

					
					wait 10

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
						wait 6 
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Siguiente >").Click
	'				End If
						
				
				
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
			Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Pago inmediato").Exist(2)) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist) Or(JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist))   
			
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
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Nombre a Facturar BAR").Exist=False
		wait 1
	Wend
	Dim nom
	nom=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Nombre a Facturar BAR").GetROProperty("text")
	While nom=""
		wait 1
		nom=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Nombre a Facturar BAR").GetROProperty("text")
	Wend
		
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Select DataTable("e_TipoPago",dtLocalSheet)
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Select DataTable("e_MetodoPago",dtLocalSheet)
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
	
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(1) Then
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
	 
	 If JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden").JavaButton("Aceptar").Exist = True Then
	 		DataTable("s_Resultado", dtLocalSheet) = "éxitoso"
			DataTable("s_Detalle", dtLocalSheet) = "exceso máximo permitido"
			Reporter.ReportEvent micPass, DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
			
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"ErrorEx.png", True
			imagenToWord "Error de exceso máximo permitido", RutaEvidencias() &"ErrorEx.png"

			JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden").JavaButton("Aceptar").Click
			wait 1
			
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Cancelar oferta").Click
			wait 1
			While JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar una opción").JavaButton("Sí").Exist = False
				wait 1	
			Wend
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar una opción").JavaButton("Sí").Click
			
			While JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden_2").JavaButton("Aceptar").Exist = False
				wait 1
			Wend
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden_2").JavaButton("Aceptar").Click
			wait 1
			While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").Exist= False
				wait 1
			Wend
			  
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"Cancel.png", True
			imagenToWord "Orden Cancelada", RutaEvidencias() &"Cancel.png"
			 ExitActionIteration         
			End If
		
End Sub		
