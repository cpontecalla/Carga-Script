Dim str_IDServicio
Dim str_Plan
Dim str_Motivo
Dim str_NroOrden
Dim str_TipoData
Dim str_ValOrden
Dim varlog, varlog2, varlog3, varlog4, varlog5, varlog6, t, varhab
Dim nroReg, varValidaRespuestaCumplimiento

varlog3 		= "<html>Si el plan origen es Prepago y destino Postpago,&#8203; el cambio de plan se realizará de inmediato y el cobro se iniciará al final del ciclo<br>Si el plan origen es Postpago,&#8203; el cambio de plan se realizará en la cíclica siguiente. (Detectado en Plan).</html>"
varlog4 		= "<html>Cambio de plan no está permitido cuando hay pendiente de Plan de"
varlog5 		= "<html>Ten en cuenta que los cambios y cobros serán efectivos en el inicio del siguiente ciclo por única vez. (Detectado en Plan).</html>"
varlog6 		= "<html>Ten en cuenta que los cambios se activaran al final de ciclo,&#8203; y al cliente se le cobrara a partir del próximo ciclo,&#8203; pero no para ofertas Prepagas. (Detectado en Plan).</html>"
Num_Iter 	    = Environment.Value("ActionIteration") 

str_IDServicio  = DataTable("e_ID_Servicio", dtLocalSheet)
str_Tipo_Categ  = DataTable("e_Tipo_Subcategoria", dtLocalSheet)
str_Categ_Plan	= DataTable("e_Tipos_Categoria_Plan", dtLocalSheet)
str_Plan    	= DataTable("e_Plan", dtLocalSheet)
str_Motivo		= DataTable("e_Motivo", dtLocalSheet)
str_NroOrden    = DataTable("s_Nro_Orden", dtLocalSheet)
str_TipoData    = DataTable("e_TipoData", dtLocalSheet)
str_ValOrden    = DataTable("s_ValEstadoOrden", dtLocalSheet)

Call PanelInteraccion()
Call ProductoAsignado()
Call DetallesProducto()
Call FlujoWIC()
Call IngresodePlan()
Call ActualizarAtributos()
Call ParametrizaProducto()
Call NegociarConfiguracion()
'Call NegociarDistribucion()
'Call Financiamiento()
Call GeneraContrato()
'If DataTable("e_Ambiente", "Login")<>"PROD" Then
	'Call EmpujeOrden()
'End If
Call ValidaOrden()
Call DetalleActividadOrden()

Sub PanelInteraccion()
wait 10		
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Ver Detalles").Exist) = False
			wait 1	
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Panel de Interacción- de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png", True
				imagenToWord "No cargó la pantalla -Panel de Interacción- de manera correcta",RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
			wait 1
		Wend
		
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Exist) = False
			wait 1	
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Panel de Interacción- de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png", True
				imagenToWord "No cargó la pantalla -Panel de Interacción- de manera correcta",RutaEvidencias() &Num_Iter&"_"&"ErrorPanelInteraccion.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
			wait 1
		Wend
		
End Sub
Sub	ProductoAsignado()
    wait 10
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Productos Asignados").Select @@ hightlight id_;_28643574_;_script infofile_;_ZIP::ssf1.xml_;_
	wait 1
	
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaEdit("TextFieldNative$1_2").Exist) = False
			wait 1	
		Wend
	
	
		
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaCheckBox("Incluir órdenes pendientes").Exist) = False
			wait 1	
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó el control -Incluir órdenes pendientes- de manera correcta"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
		wait 1
	
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaCheckBox("Incluir órdenes pendientes").Set "ON" @@ hightlight id_;_21562001_;_script infofile_;_ZIP::ssf3.xml_;_
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaEdit("TextFieldNative$1_2").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaEdit("TextFieldNative$1_2").Set str_IDServicio
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Click @@ hightlight id_;_31932278_;_script infofile_;_ZIP::ssf4.xml_;_
	wait 2

		tiempo=0
		Do 
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Buscar ahora").Click
				nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("1 Registros").GetROProperty("label")
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

		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaTable("Ver por:").Exist) = False
			wait 1	
		Wend
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"IDServicioEncontrado.png", True
	imagenToWord "Id de Servicio: "&str_IDServicio&" encontrado", RutaEvidencias() &Num_Iter&"_"&"IDServicioEncontrado.png"	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaTable("Ver por:").DoubleClickCell 0, "#2", "LEFT"
	wait 4
	
End Sub
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
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Calcular Penalidad").Exist
		Wait 2
	
	JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").Select
	If JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").JavaMenu("Reemplazar Paquetes").GetROProperty("enabled") = "1" Then
		JavaWindow("Ejecutivo de interacción").JavaMenu("Acciones").JavaMenu("Pedidos").JavaMenu("Reemplazar Paquetes").Select
		Else 
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
	    DataTable("s_Detalle", dtLocalSheet) = "No se puede cambiar de plan al número: "&str_IDServicio&", ya que la opción Reemplazar Paquetes esta deshabilitada"
	    Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
	    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OpcionSuspenderDeshabilitada.png", True
		imagenToWord "No se puede cambiar de plan al número: "&str_IDServicio&" ya que la opción Reemplazar Paquetes esta deshabilitada", RutaEvidencias() &Num_Iter&"_"&"OpcionSuspenderDeshabilitada.png"
	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaButton("Cerrar").Click
		ExitActionIteration
	End If
	
		tiempo = 0
		Do
			tiempo = tiempo + 1 
			If tiempo >= 180 Then
			    DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			    DataTable("s_Detalle", dtLocalSheet) =  "Error al cargar la pantalla, no cargo la pantalla Detalles del producto"
			    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorActualizarAtributos.png", True
				imagenToWord "Error al cargar, no cargó la pantalla Actualizar Atributos", RutaEvidencias() &Num_Iter&"_"&"ErrorActualizarAtributos.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) ,  DataTable("s_Detalle", dtLocalSheet)
				else
				Reporter.ReportEvent micPass, "Exito", "Cargo correctamente la pantalla Detalles del producto"
			End If
			wait 2
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").Exist) or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaEdit("TextFieldNative$1").Exist))
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").Exist Then
		varlog2 = JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").JavaTable("Las siguientes acciones").GetCellData(0,0)
	  	DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		DataTable("s_Detalle", dtLocalSheet) = varlog2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MensajeError.png", True
		imagenToWord "Mensaje de Error",RutaEvidencias() &Num_Iter&"_"&"MensajeError.png"
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").JavaButton("Rechazar solicitud de").Click
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
		wait 2
		ExitActionIteration
	End If
	
End Sub
Sub FlujoWIC()

	If DataTable("e_WIC_ValidaCli",dtLocalSheet)="SI" Then
		

RunAction "WIC", oneIteration
	End If
	'	    	'Se da Clic en el botón continuar
'			If UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar").UIAObject("Consulta previa").Exist(3) And UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar").UIAButton("Continuar").Exist(3) Then
'				UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar").UIAButton("Continuar").Click
'			else
'				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
'				DataTable("s_Detalle", dtLocalSheet) = "No se habilitó el botón CONTINUAR"	
'				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
'				UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").Close
'				Wait 5
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
'				Wait 5
'				ExitActionIteration	
'			End If
'			Wait 3
'			
'				'Se espera la carga de la 2da pantalla WIC
'				t = 0
'				While (UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar").UIAObject("Scoring cliente").Exist) = False
'					t = t + 1
'					If t >= 60 Then
'						DataTable("s_Resultado", dtLocalSheet) = "Fallido"
'						DataTable("s_Detalle", dtLocalSheet) = "La siguiente pantalla no cargó y no se pudo continuar con el flujo WIC luego de 60 intentos"
'						Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
'						UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").Close
'						Wait 5
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
'						Wait 5
'						ExitActionIteration
'					End If
'				Wend
'		
'			'Se ejecuta flujo en 2da Pantalla WIC
'			If UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar").UIAObject("Scoring cliente").Exist Then
'				UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAComboBox("Tipo de sustento:").UIAButton("Abrir").Click
'				Wait 1
'				UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAComboBox("Tipo de sustento:").Select "Boleta de pago - Fijo"
'				Wait 3
'				UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAComboBox("Sustento:").UIAButton("Abrir").Click
'				Wait 1
'				UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAComboBox("Sustento:").Select "ingreso neto último mes"
'				Wait 3
'				UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAEdit("Valor de sustento:").SetValue "5000"
'				Wait 3
'				UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAButton("Calcular").Click
'			Else
'				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
'				DataTable("s_Detalle", dtLocalSheet) = "No se habilitó el botón CALCULAR"
'				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
'				UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").Close
'				Wait 5
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
'				Wait 5
'				ExitActionIteration
'			End If
'		
'			'Control de Mensaje de Error	
'			If JavaDialog("Error Message").Exist(5) Then
'				var1 = JavaDialog("Error Message").JavaObject("JPanel").GetROProperty("text")
'				JavaDialog("Error Message").JavaButton("Cancelar").Click   
'				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
'				DataTable("s_Detalle", dtLocalSheet) = var1
'				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
'				Wait 5
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
'				Wait 5
'				ExitActionIteration
'			End If	
'	
'				'Se espera la carga de la 3ra pantalla WIC
'				t = 0
'				While (UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAObject("Score").Exist Or UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAWindow("Error Message").Exist) = False
'					t = t + 1
'					If t >= 60 Then
'						DataTable("s_Resultado", dtLocalSheet) = "Fallido"
'						DataTable("s_Detalle", dtLocalSheet) = "La siguiente pantalla no cargó y no se pudo continuar con el flujo WIC luego de 60 intentos"
'						Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
'						UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").Close
'						Wait 5
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
'						Wait 5
'						ExitActionIteration
'					End If
'				Wend
'		
'			'Se ejecuta flujo en 3ra Pantalla WIC
'			If (UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAObject("Score").Exist) And (UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar").UIAObject("APROBAR").Exist) Then
'				If UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAButton("Continuar").Exist(3) Then
'					If UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAButton("Imprimir").Exist(3) Then
'						UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAButton("Continuar").Click
'					Else
'						DataTable("s_Resultado", dtLocalSheet) = "Fallido"
'						DataTable("s_Detalle", dtLocalSheet) = "No se habilitó el botón CONTINUAR o IMPRIMIR"	
'						Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
'						UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").Close
'						Wait 5
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
'						Wait 5
'						ExitActionIteration			
'					End If
'				End If
'			End If
'		
'				'Se espera la carga de la 4ta pantalla WIC
'				t = 0
'				While (UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIACheckBox("No conoce centro poblado").Exist) = False
'					t = t + 1
'					If t >= 60 Then
'						DataTable("s_Resultado", dtLocalSheet) = "Fallido"
'						DataTable("s_Detalle", dtLocalSheet) = "La siguiente pantalla no cargó y no se pudo continuar con el flujo WIC luego de 60 intentos"
'						Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
'						UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").Close
'						Wait 5
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
'						Wait 5
'						ExitActionIteration
'					End If
'				Wend
'		
'			'Se ejecuta flujo en 4ta Pantalla WIC
'			If UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIACheckBox("No conoce centro poblado").Exist Then
'				UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").UIAObject("Movistar_2").UIAButton("Continuar_2").Click
'			Else
'				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
'				DataTable("s_Detalle", dtLocalSheet) = "No se habilitó el botón CONTINUAR"	
'				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
'				UIAWindow("Ejecutivo de interacción").UIAWindow("Autenticación del Cliente").Close
'				Wait 5
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
'				Wait 5
'				ExitActionIteration		
'			End If
'		End If
'	End If
End Sub
Sub IngresodePlan()

		While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaEdit("TextFieldNative$1").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Detalles del producto").Exist)) = False
			wait 1	
		Wend
	wait 8
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaList("ComboBoxNative$1").Select str_Tipo_Categ @@ hightlight id_;_25126001_;_script infofile_;_ZIP::ssf7.xml_;_
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaList("ComboBoxNative$1").Select str_Categ_Plan @@ hightlight id_;_29115386_;_script infofile_;_ZIP::ssf8.xml_;_
	wait 2	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaEdit("TextFieldNative$1").Set str_Plan @@ hightlight id_;_29919409_;_script infofile_;_ZIP::ssf7.xml_;_
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaButton("Buscar").Click @@ hightlight id_;_8794597_;_script infofile_;_ZIP::ssf8.xml_;_
	wait 2
	
		tiempo = 0
		Do
			tiempo = tiempo + 1
			If (tiempo >= 180) Then
				Reporter.ReportEvent micFail, "Error de Búsqueda", "No se encontró el Plan Ingresado"
				DataTable("s_Resultado",dtLocalSheet) = "Fallido"
				DataTable("s_Detalle",dtLocalSheet)   = "Error de Búsqueda no se encontró el plan: "&str_Plan
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"PlanNoEncontrado.png", True
				imagenToWord "Plan No Encontrado",RutaEvidencias() &Num_Iter&"_"&"PlanNoEncontrado.png"
				ExitActionIteration
			Else
				Reporter.ReportEvent micPass, "Exito", "Se ha cargado el plan correctamente"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"PlanEncontrado.png", True
				imagenToWord "Plan Encontrado",RutaEvidencias() &Num_Iter&"_"&"PlanEncontrado.png"			
			End If
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaCheckBox("Seleccionar").Exist
		wait 1
		
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaCheckBox("Seleccionar").Set "ON" @@ hightlight id_;_11194932_;_script infofile_;_ZIP::ssf9.xml_;_
	wait 2

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaButton("Siguiente >").Click @@ hightlight id_;_8669108_;_script infofile_;_ZIP::ssf10.xml_;_
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"PlanSeleccionado.png", True
	imagenToWord "Plan Seleccionado",RutaEvidencias() &Num_Iter&"_"&"PlanSeleccionado.png"
	
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaStaticText("Totales(st)").Exist) = False
			wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Nuevo Plan- de manera correcta"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
		wait 1
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para DORA").JavaButton("Siguiente >").Click @@ hightlight id_;_9477847_;_script infofile_;_ZIP::ssf11.xml_;_
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"PantallaSeleccionarPlan.png", True
	imagenToWord "Pantalla Seleccionar Plan",RutaEvidencias() &Num_Iter&"_"&"PantallaSeleccionarPlan.png"
	
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de_2").JavaList("Motivo:").Exist) = False
			wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Actualizar atributos- de manera correcta"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorPantallaActualizarAtributos.png", True
				imagenToWord "Error Pantalla Actualizar Atributos",RutaEvidencias() &Num_Iter&"_"&"ErrorPantallaActualizarAtributos.png"
				ExitActionIteration
			End If
		Wend
End Sub
Sub ActualizarAtributos()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de_2").JavaList("Motivo:").Select str_Motivo
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de_2").JavaEdit("Texto del motivo:").Set "Cambio de Plan" @@ hightlight id_;_1371164_;_script infofile_;_ZIP::ssf14.xml_;_
	wait 2
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"PantallaActualizarAtributos.png", True
	imagenToWord "Pantalla Actualizar Atributos",RutaEvidencias() &Num_Iter&"_"&"PantallaActualizarAtributos.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Siguiente >").Click @@ hightlight id_;_15762962_;_script infofile_;_ZIP::ssf17.xml_;_

		t = 0
		While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Exist)) = False
			wait 1
			t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Selección de Contacto- de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorSeleccionContacto.png", True
				imagenToWord "Error de Selección de Contacto.png",RutaEvidencias() &Num_Iter&"_"&"ErrorSeleccionContacto.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
		wait 1
	
End Sub
Sub	ParametrizaProducto()
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Set
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaButton("Siguiente >").Click
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"PantallaSeleccionContacto.png", True
		imagenToWord "Pantalla Selección de Contacto",RutaEvidencias() &Num_Iter&"_"&"PantallaSeleccionContacto.png"
	End If
	
		t = 0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Exist) = False
			wait 1
				t = t + 1
			If (t >= 180) Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Negociar Configuración del Producto Móvil- de manera correcta"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorPantallaNegociarConfiguracionProductoMovil.png", True
				imagenToWord "Error Pantalla Negociar Configuracion Producto Movil",RutaEvidencias() &Num_Iter&"_"&"ErrorPantallaNegociarConfiguracionProductoMovil.png"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
	wait 1
End Sub
Sub NegociarConfiguracion()

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click @@ hightlight id_;_27691117_;_script infofile_;_ZIP::ssf21.xml_;_
	wait 7
	
		t=0
		Do
		t=t+1
			wait 2
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cancelar").Exist Then
				wait 1
				varlog = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable_2").GetCellData(0,1)
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MensajeValidaciónNuevo.png", True
				imagenToWord "Mensaje de Validación",RutaEvidencias() &Num_Iter&"_"&"MensajeValidaciónNuevo.png"
				wait 1	
					Select Case varlog
						Case varlog3
							wait 2
							JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Validacion.png", True
							imagenToWord "Pantalla Negociar Configuración",RutaEvidencias() &Num_Iter&"_"&"Validacion.png"
							wait 1
							JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Aceptar").Click
							wait 2
							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Click
						Case varlog4
							'(Left((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable_2").GetCellData(0,1)),70) = varlog4)
							JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCambioPlanNoPermitido.png", True
							imagenToWord "Error Cambio Plan No Permitido",RutaEvidencias() &Num_Iter&"_"&"ErrorCambioPlanNoPermitido.png"
							wait 1
							JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
							wait 3
							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
							wait 3
							JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
							wait 3
							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click
							DataTable("s_Resultado", dtLocalSheet) = "Fallido"
							DataTable("s_Detalle", dtLocalSheet) = varlog2
							Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
							ExitActionIteration
						Case varlog5
							wait 2
							JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Validacion.png", True
							imagenToWord "Pantalla Negociar Configuración",RutaEvidencias() &Num_Iter&"_"&"Validacion.png"
							wait 1
							JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Aceptar").Click
							wait 2
							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Click
						Case varlog6
						wait 2
							JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Validacion.png", True
							imagenToWord "Pantalla Negociar Configuración",RutaEvidencias() &Num_Iter&"_"&"Validacion.png"
							wait 1
							JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Aceptar").Click
							wait 2
							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Click
					End Select
			End If
		Loop While Not (t=15)
		
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Click
			wait 1
		End If

		tiempo = 0
		Do
			tiempo = tiempo + 1
			If (tiempo >= 180) Then
				Reporter.ReportEvent micFail, "Error", "No cargó la pantalla que contiene el ID del Acuerdo Financiero"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorAcuerdoFacturacion.png", True
				imagenToWord "Error Acuerdo de Facturacion",RutaEvidencias() &Num_Iter&"_"&"ErrorAcuerdoFacturacion.png"
			else
				Reporter.ReportEvent micPass, "Éxito en carga de ID del Acuerdo Financiero", "Se ha cargado correctamente los datos del ID del Acuerdo Financiero"	
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"AcuerdoFacturacionActual.png", True
				imagenToWord "Acuerdo Facturación Actual",RutaEvidencias() &Num_Iter&"_"&"AcuerdoFacturacionActual.png"
			End If
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("ID del Acuerdo de Facturación:").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaButton("Pago inmediato").Exist))
		
		
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("ID del Acuerdo de Facturación:").Exist = True Then
			Call NegociarDistribucion()
			'Call Financiamiento()
		ElseIf JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaButton("Pago inmediato").Exist=True Then
			
			Call Financiamiento()
			
		End If
End Sub
Sub NegociarDistribucion()
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("Nombre y Dirección de").Exist=False
			wait 1
		Wend
		wait 2
		Dim nom
		nom=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("Nombre y Dirección de").GetROProperty("text")
		
		While nom=""
			wait 1
			nom=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("Nombre y Dirección de").GetROProperty("text")
		Wend

	wait 5
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("ID del Acuerdo de Facturación:").WaitProperty "editable", 1, 10000
	
			Do While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaButton("Lookup-Validated").Exist) = False
				wait 1
					c=c+1 
					If (c=30) Then exit Do 
			Loop
		wait 1

	t = 0
	While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("Nombre y Dirección de").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist)) = False
		wait 1	
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No cargó la pantalla -Negociar Distribución- de manera correcta"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorPantallaNegociarConfiguracion.png", True
			imagenToWord "No cargó la pantalla -Negociar Distribución- de manera correcta",RutaEvidencias() &Num_Iter&"_"&"ErrorPantallaNegociarConfiguracion.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
	Wend
	wait 4
	
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaRadioButton("Nuevo").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaRadioButton("Nuevo").Set
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaList("Mostrar:").Select "Acciones de orden activas "
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaButton("Siguiente >").Click @@ hightlight id_;_15762962_;_script infofile_;_ZIP::ssf26.xml_;_
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"AcuerdoFacturacionNuevo.png", True
		imagenToWord "Acuerdo Facturación Nuevo",RutaEvidencias() &Num_Iter&"_"&"AcuerdoFacturacionNuevo.png"
	End If
	
		tiempo = 0
		Do
			tiempo = tiempo + 1
			If (tiempo >= 180) Then
				Reporter.ReportEvent micFail, "Error de Resumen de Contrato", "No se ha cargado el resumen correctamente"
			Else
				Reporter.ReportEvent micPass, "Resumen de Contrato Exitoso", "Se ha cargado el resumen correctamente"	
			End If
		wait 1
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaButton("Pago inmediato").Exist))
		 
		 If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaButton("Pago inmediato").Exist=True Then
		 	
			Call Financiamiento()
		 End If
		
			
End Sub
Sub Financiamiento()
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaEdit("ID del cliente:").Exist=False
		wait 1
	Wend
	wait 3
	Dim textID
	textID=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaEdit("ID del cliente:").GetROProperty("text")
	While textID =""
    	wait 1
    	textID=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaEdit("ID del cliente:").GetROProperty("text")
    Wend
    	wait 1
    Dim finEx
	 finEx=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaCheckBox("Financiamiento Externo").GetROProperty("enabled") 

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaCheckBox("Financiamiento Externo").Exist =True and finEx="1" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaCheckBox("Financiamiento Externo").Set "OFF"
	End If
   

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaButton("Pago inmediato").Exist Then
		wait 2
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaButton("Límite de Compra").Exist Then
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Negociar Pago"&".png" , True
			imagenToWord "Negociar Pago", RutaEvidencias() &Num_Iter&"_"&"Negociar Pago"&".png"
			varfinan=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaButton("Límite de Compra").GetROProperty("enabled")
			wait 1
			If varfinan = "1" Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaButton("Límite de Compra").Click	
			End If
			wait 2
		End If
		
			tiempo=0
				Do 
				tiempo=tiempo+1
				varpagoinm=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaButton("Pago inmediato").GetROProperty("enabled")
				wait 2	
			Loop While Not (varpagoinm="1")
			wait 2
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaButton("Pago inmediato").Click
		wait 2
			
			While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Exist) = False
				wait 1
			Wend
			Dim Iterator, Count
				Count = 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").GetROProperty ("items count")
			
				For Iterator = 0 To Count-1
				 	rs = 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").GetItem (Iterator)
				 	
					If rs = DataTable("e_MedioPago", dtLocalSheet) Then
							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Select DataTable("e_MedioPago", dtLocalSheet)
'						    wait 1
'								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Cantidad de cuotas:").Select DataTable("e_Cant_Cuota" , dtLocalSheet)
'								wait 1
'								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Calcular").Click
								JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Financiamiento"&".png" , True
								imagenToWord "Financiamiento", RutaEvidencias() &Num_Iter&"_"&"Financiamiento"&".png"
						
						Exit for
					ElseIf Iterator = Count-1 Then
							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Select "Externo"
						Exit for
					End if	
					Next
				wait 3
				
				
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DetallePago"&".png" , True
		imagenToWord "Detalle Pago", RutaEvidencias() &Num_Iter&"_"&"DetallePago"&".png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Enviar").Click
		
		While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist))=False
			wait 1
		Wend
		
		wait 1
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(2) Then
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Validacion"&".png", True
			imagenToWord "Mensaje de Validación", RutaEvidencias() &Num_Iter&"_"&"Validacion"&".png"
			wait 1
			varsap=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel_2").GetROProperty("text")

			If varsap="El RUC es obligatorio para la Factura" Then
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Select "Boleta"
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Enviar").Click
				wait 3
			End If
		End If
				
		wait 5
			While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaButton("Siguiente >").Exist) = False
				wait 1
			Wend
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 846329A").JavaButton("Siguiente >").Click
		
		While((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist) or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist)) = False
			wait 1
		Wend
		
		
		
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
				var1 = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
				var1 = replace(var1,"<html>","")
				var1 = replace(var1,"</html>","")
				DataTable("s_Resultado",dtLocalSheet) = "Fallido"
				DataTable("s_Detalle",dtLocalSheet)=var1
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 1221256A").JavaButton("Cerrar").Click
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
				wait 5
				ExitActionIteration
			End If
	End If
End  Sub
Sub GeneraContrato() @@ hightlight id_;_29139057_;_script infofile_;_ZIP::ssf28.xml_;_
	
	wait 1
	varhab=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").GetROProperty("enabled")
	If  varhab<>"0" Then
	
		tiempo = 0
		Do
			While((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist)) = False
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Click
				If DataTable("e_WIC_ContrCli",dtLocalSheet)="SI" Then
					

RunAction "WIC2", oneIteration


					Exit do
				End If
				wait 3
			Wend 
	
	'		'Click "Validade y Ver Contrato"
	'		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist(2) Then
	'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Click
	'			If DataTable("e_WIC_Activa",dtLocalSheet)="SI" Then
	'				RunAction "WICGenContrato [WICGenContrato]", oneIteration
	'				Exit do
	'			End If
	'			Wait 5
	'		End If	
		
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(2) Then
				wait 3
				var1= JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
				wait 1
			End If
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist(2) Then
				JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
				wait 2
			End If
			If (tiempo >= 180) Then
				DataTable("s_Resultado",dtLocalSheet) = "Fallido" 
				DataTable("s_Detalle", dtLocalSheet) = "No se ha cargado el contrato correctamente"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Resultado",dtLocalSheet)
				ExitActionIteration
			Else
				Reporter.ReportEvent micPass, "Contrato Exitoso", "Se ha cargado el contrato correctamente"	
			End If
			wait 2
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist(2)) or (var1 = "Contratos no Generados") or (var1 = "0"))
		
		If DataTable("e_WIC_ContrCli",dtLocalSheet)<>"SI" Then
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist(2) Then
				JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Close
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").Set "ON"
				Wait 3
			End If
		End If
	End  If
	
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
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Exist(2) Then
		'Damos clic en el boton "Enviar Orden"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
		Wait 8
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist=True Then
			wait 3
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Aceptar").Click
			
		End If
		
	End If
	
	'Bucle que espera el envío de la orden
	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").Exist) = False
		wait 1
	Wend
	
	'Se captura la orden generada
'	Select Case DataTable("e_Ambiente", "Login [Login]")
'		Case "UAT8"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),8)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Click
'			wait 2
'		Case "UAT4"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Click
'			wait 2
'		Case "UAT6"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Click
'			wait 2		
'		Case "UAT10"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Click
'			wait 2	
'		Case "UAT13"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Click
'			wait 2
'		Case "PROD"
'			DataTable("s_Nro_Orden", dtLocalSheet) = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),10)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").CaptureBitmap RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png", True
'			imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_"&Num_Iter&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Click
'			wait 2				
'	End Select

    Dim text
    text = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").GetROProperty("text")
    DataTable("s_Nro_Orden", dtLocalSheet) = RTRIM(LTRIM((replace(text,"Orden",""))))
    WAIT 1
    
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").CaptureBitmap RutaEvidencias() & "Orden_Generada_.png", True
    imagenToWord "Orden_Generada_"&Num_Iter,RutaEvidencias() & "Orden_Generada_.png"
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Click
    
End Sub
Sub EmpujeOrden() @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf2.xml_;_

	If (str_TipoData = "DATA LOGICA") Or (DataTable("s_ValEstadoOrden", dtLocalSheet) = "Enviado") Then
	
		tiempo=0
		Do 
			tiempo=tiempo+1
			vardepo=JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").GetROProperty("enabled")
			wait 2
		Loop While Not (vardepo="1")
	
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select @@ hightlight id_;_19748072_;_script infofile_;_ZIP::ssf61.xml_;_
		Wait 1 @@ hightlight id_;_19748072_;_script infofile_;_ZIP::ssf61.xml_;_
		
			While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Exist) = False
				Wait 1
			Wend
				
			Do
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Select "Tareas pendientes del equipo" @@ hightlight id_;_30837205_;_script infofile_;_ZIP::ssf3.xml_;_
				Wait 2
			Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Finalizar compra y activar").Exist
			Wait 3
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Exist=False
			wait 1
		Wend	
		wait 10

		'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").SetFocus
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
		Wait 3
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
		Wait 2

			tiempo=0
			Do 
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
					nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("1 Registros").GetROProperty("text")
					tiempo=tiempo+1
					wait 1
				End If
				If (tiempo >= 180) Then
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
		
'			tiempo=0
'			Do
'				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
'					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
'					wait 2
'					tiempo = tiempo+1
'					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Output CheckPoint("Equipo usuario:") @@ hightlight id_;_22747135_;_script infofile_;_ZIP::ssf1.xml_;_
'					varValidaRespuestaCumplimiento = Environment("s_ValidaManejarRespuestaCumplimiento")
'					wait 1
'				End If
'					If (tiempo >= 180) Then
'						DataTable("s_Resultado",dtLocalSheet)="Fallido"
'						DataTable("s_Detalle",dtLocalSheet)="La actividad 'Manejar Respuesta de Cumplimiento' no cargo"	
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
'						wait 2
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
'						wait 2
'						ExitTestIteration
'					End If 
'			Loop While Not varValidaRespuestaCumplimiento = "Manejar Respuesta de Cumplimiento"


		Wait 5
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").SelectRow "#0"

		Wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Gestión manual").Click
		
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaList("Estado de la gestión manual:").Exist) = False
				
				Wait  1
		Wend
			Wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaList("Estado de la gestión manual:").Select "Cumplimiento Completo"
		Wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaList("Motivo de la gestión manual").Select "Manejo manual: Manejo Manual OSS"

		Wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").Click

			While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").Exist) = False
				Wait 1
			Wend
	End If
		
End Sub
Sub ValidaOrden()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
	
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Exist) = False
			wait 1
		Wend
		wait 10
	'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaCheckBox("Solo órdenes pendientes").Set "OFF"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
	wait 2
	
		tiempo=0
		Do 
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
				nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("0 Registros").GetROProperty("attached text")
				tiempo=tiempo+1
				wait 1
			End If
			If (tiempo >= 180) Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar la Gestion Logistica"
					Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
					ExitActionIteration
			End If
		Loop While Not (nroreg="1 Registros")
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:") @@ hightlight id_;_26172125_;_script infofile_;_ZIP::ssf85.xml_;_
	Reporter.ReportEvent micPass,"Se valida el estado de la orden", DataTable("s_ValEstadoOrden", dtLocalSheet)
	
		tiempo = 0
		Do
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
			tiempo = tiempo +1
			wait 3
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:")
				If (tiempo >= 180) Then	
						DataTable("s_Resultado",dtLocalSheet) = "Fallido" 
						DataTable("s_Detalle", dtLocalSheet) = "La Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)&" no culmino en estado Programado"
						Reporter.ReportEvent micFail,DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
						JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OrdenFallida.png", True
						imagenToWord "Orden Fallida",RutaEvidencias() &Num_Iter&"_"&"OrdenFallida.png"
						If DataTable("s_ValEstadoOrden", dtLocalSheet) = "Enviado" Then
							Exit Do
							wait 1
						End If	
				else
					Reporter.ReportEvent micPass,"Correcto", "Se valida el estado de la orden: "&DataTable("s_Nro_Orden",dtLocalSheet)
				End If
		Loop While Not ((DataTable("s_ValEstadoOrden", dtLocalSheet) = "Programado") Or (DataTable("s_ValEstadoOrden", dtLocalSheet) = "Cerrado")Or (DataTable("s_ValEstadoOrden", dtLocalSheet) = "Enviado"))
	


		If (DataTable("s_ValEstadoOrden", dtLocalSheet) = "Programado") Or (DataTable("s_ValEstadoOrden", dtLocalSheet) = "Cerrado")=True Then
			DataTable("s_Resultado", dtLocalSheet) = "Éxito"
			DataTable("s_Detalle", dtLocalSheet) = "La orden culminó correctamente en el estado"&DataTable("s_ValEstadoOrden", dtLocalSheet)
			Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OrdenExitosa.png", True
			imagenToWord "Orden Exitosa",RutaEvidencias() &Num_Iter&"_"&"OrdenExitosa.png"
			wait 2
	
			ElseIf (DataTable("s_ValEstadoOrden", dtLocalSheet) = "Enviado") Then
			 Call EmpujeOrden() 
			 wait 2
				JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
				JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
				
					While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Exist) = False
						wait 1
					Wend
					wait 3
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").SetFocus
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaCheckBox("Solo órdenes pendientes").Set "OFF"
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
				wait 2

				tiempo=0
						Do 
							If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Exist Then
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
								nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("0 Registros").GetROProperty("attached text")
								tiempo=tiempo+1
								wait 1
							End If
							If (tiempo >= 180) Then
									JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
									DataTable("s_Resultado", dtLocalSheet) = "Fallido"
									DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar la Gestion Logistica"
									Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
									JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
									ExitActionIteration
							End If
						Loop While Not (nroreg="1 Registros")
					
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:") @@ hightlight id_;_26172125_;_script infofile_;_ZIP::ssf85.xml_;_
					Reporter.ReportEvent micPass,"Se valida el estado de la orden", DataTable("s_ValEstadoOrden", dtLocalSheet)
					
						tiempo = 0
						Do
							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
							tiempo = tiempo +1
							wait 3
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:")
								
								
								If (tiempo >= 180) Then	
										DataTable("s_Resultado",dtLocalSheet) = "Fallido" 
										DataTable("s_Detalle", dtLocalSheet) = "La Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)&" no culmino en estado Cerrado"
										Reporter.ReportEvent micFail,DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
										JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OrdenFallida.png", True
										imagenToWord "Orden Fallida",RutaEvidencias() &Num_Iter&"_"&"OrdenFallida.png"
'										If DataTable("s_ValEstadoOrden", dtLocalSheet) = "Enviado" Then
'											Exit Do
'											wait 1
'										End If	
								else
									Reporter.ReportEvent micPass,"Correcto", "Se valida el estado de la orden: "&DataTable("s_Nro_Orden",dtLocalSheet)
								End If
						Loop While Not ((DataTable("s_ValEstadoOrden", dtLocalSheet) = "Cerrado"))
						JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OrdenExitosaa.png", True
						imagenToWord "Orden Exitosa",RutaEvidencias() &Num_Iter&"_"&"OrdenExitosaa.png"
			 
			
		End If
		
End Sub
Sub DetalleActividadOrden()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaList("Ver por:").Select "Acciones de orden"
	wait 3
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").DoubleClickCell 0, "#8", "LEFT"
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{ENTER}"
	wait 1

	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 794525A").JavaEdit("Fecha de vencimiento:").Exist)=False
		wait 1
	Wend
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 794525A").JavaTab("Nombre del cliente:").Select "Actividad"
	
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 794525A").JavaTable("SearchJTable").Exist)=False
		wait 1	
	Wend
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_1"&".png", True
	imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_1"&".png"
	
	shell.SendKeys "{PGDN}"
	wait 1
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_2"&".png", True
	imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_2"&".png"
	
	
	dim Iterator , filas	
	filas=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 794525A").JavaTable("SearchJTable").GetROProperty("rows")
	For Iterator = filas-1 To 0 step -1	    
		varselec=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 794525A").JavaTable("SearchJTable").GetCellData(Iterator,0)	
		
		If varselec="Cerrar Acción de Orden" or varselec="Actualizar Descuento" or varselec="Finalizar compra en la Negociación" Then
			DataTable("s_Resultado",dtLocalSheet)="Exitoso"
			DataTable("s_Detalle",dtLocalSheet)="La orden "&DataTable("s_Nro_Orden",dtLocalSheet)&" culmino en estado Cerrado, exitoso en la Actividad "&varselec&""
			Reporter.ReportEvent micPass, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet) 	   			    	
		     Exit for 	    
		     Else 
		     	DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)="La orden "&DataTable("s_Nro_Orden",dtLocalSheet)&" no culmino en estado Cerrado, falló en la Actividad "&varselec&""
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
		
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 794525A").JavaButton("Cancelar").Click

		
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
wait 1

'	
'	filas=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 794525A").JavaTable("SearchJTable").GetROProperty("rows")
'	For Iterator = 0 To filas-1
'		varselec=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 794525A").JavaTable("SearchJTable").GetCellData(Iterator,0)
'	Next
'	
'	If varselec<>"Cerrar Acción de Orden" Then
'	 	DataTable("s_Resultado",dtLocalSheet)="Fallido"
'		DataTable("s_Detalle",dtLocalSheet)="La orden "&DataTable("s_Nro_Orden",dtLocalSheet)&" culmino en la actividad "&varselec&""
'		Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
'		
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 794525A").JavaButton("Cancelar").Click
'
'		wait 2
'		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
'			wait 2
'		End If
'		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(2) Then
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
'		End If
'		ExitActionIteration
'		wait 1
'	End If
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 794525A").JavaButton("Cancelar").Click

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
		wait 2
	End If
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
	End If
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Exist(3) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto_2").JavaButton("Cerrar").Click	
		wait 2
	End If
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Exist(3) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Productos asignados").JavaButton("Cerrar").Click
		wait 2
	End If
	
End Sub




