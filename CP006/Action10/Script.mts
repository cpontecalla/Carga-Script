
Option  Explicit

Dim Valor, Fila, Valor2, i, Filas, varsap, tiempo, var, var1, flag, vartext, vardisp, varsim, shell, varlog, varhab, Num_Iter, varfinan, varfila, varmsgval, varId, Count, rs
Dim varValidaRespuestaCumplimiento, nroreg, varpagoinm, varnuevo, varprob, varasig, varasig2, varacuer, vardepo, varlogis, Iterator, varselec, str_titulo, varmsg, varvend
Dim str_tipoalta
Dim str_motivoalta
Dim str_tipoSIM
Dim str_metodo_entrega
Dim str_Financiamiento
Dim str_Cuotas
Dim str_mediopago
Dim str_cant_cuota
Dim str_tipodata
Dim str_idDispositivo
Dim str_idSim
Dim varimporte, var6, var8
Dim c,j,h,nodeName,nodeName2nd,nodeName3nd,varpag, vartexto, textCodVen

Num_Iter 		 	= Environment.Value("ActionIteration") 
str_tipoalta     	= DataTable("e_Tipo_Alta", dtLocalSheet)
str_motivoalta   	= DataTable("e_Motivo_Alta", dtLocalSheet)
str_tipoSIM       	= DataTable("e_TipoSIM", dtLocalSheet)
str_Financiamiento	= DataTable("e_Financiamiento", dtLocalSheet)
str_Cuotas			= DataTable("e_Cuotas", dtLocalSheet)
str_metodo_entrega	= DataTable("e_Metodo_Entrega", dtLocalSheet)
str_mediopago		= DataTable("e_MedioPago", dtLocalSheet)
str_cant_cuota    	= DataTable("e_Cant_Cuota", dtLocalSheet)
str_tipodata      	= DataTable("e_Tipo_De_DATA_Sim", dtLocalSheet)
str_idDispositivo 	= DataTable("e_ID_Dispositivo", dtLocalSheet)
str_idSim         	= DataTable("e_ID_SIM", dtLocalSheet)
Num_Iter		  	= Environment.Value("ActionIteration")

Call EncontrarAcuerdoComercial()
Call ParametrosAlta()
Call RecursosAlta()
Call TipoEnvio()
Call PagoInmediato()
Call GeneracionOrden()
Call PagoManual()
Call GestionLogistica()
Call EmpujeOrden()
Call OrdenCerrado()
Call DetalleActividadOrden()

Sub EncontrarAcuerdoComercial()
	wait 1
		Do
		tiempo=0
		Do
			tiempo=tiempo+1
			varacuer = JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Visión General del Cliente").GetROProperty("enabled")
			wait 1
		Loop While Not (varacuer="1")
	
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Visión General del Cliente").Select
	wait 1
	

		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("BAR ID").Exist) = False
			wait 1
		Wend
		
	Valor = ""
		filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("BAR ID").GetROProperty("rows")
		
		For Iterator = 0 To filas-1 step 1 
			
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("BAR ID").ClickCell Iterator , 4
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("BAR ID").DoubleClickCell "#"&Iterator ,"#4", "LEFT"
'				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("BAR ID").PressKey "C",micCtrl
				JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").PressKey "V",micCtrl
				'MsgBox Valor
				Valor = JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").GetROProperty("value")
				'MsgBox Valor
				Valor = Left(Valor,21)
				Valor2 = DataTable("e_Grupo_Negocio","A_Comercial")
		If Valor = Valor2 Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("BAR ID").ActivateRow "#"&Iterator
			'MsgBox Valor2
			Reporter.ReportNote "Se encontro el valor"
        	JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").Set ""
   			Exit For
    	End If
		If Iterator = filas -1 Then
			Reporter.ReportNote "No se encontró el valor, se repetira la acción"
        	JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").Set ""
        	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaButton("Cerrar").Click
    	Exit For
		End If
    	JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").Set ""
		Next
	Loop While Not Valor = Valor2
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"AcuerdoComercialSeleccionado"&".png" , True
	imagenToWord "Acuerdo Comercial Seleccionado", RutaEvidencias() &Num_Iter&"_"&"AcuerdoComercialSeleccionado"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaButton("Alta Individual").Click
	wait 2
End Sub	
Sub ParametrosAlta()
	
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Proporcionar suscriptor").JavaButton("<html>Dar de alta</html>").Exist )= False
			wait 1
		Wend
	If str_tipoalta<>"Alta Nueva solo SIM" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Proporcionar suscriptor").JavaTable("Mensual").SelectRow "#0"
		'MsgBox "Selecciona Equio Móvil"
		wait 1
	End If
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Proporcionar suscriptor").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Proporciopnar_Suscriptor_"&".png" , True
	imagenToWord "Proporcionar Suscriptor", RutaEvidencias() &Num_Iter&"_"&"Proporciopnar_Suscriptor_"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Proporcionar suscriptor").JavaButton("<html>Dar de alta</html>").Click
	'MsgBox "Selecciona Dar de alta"
	
		tiempo=0
		Do
		tiempo=tiempo+1
			If tiempo>=180 Then
				DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)="No cargo la ventana Actualizar Atributo no cargo"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaActualizarAtributos"&".png" , True
				imagenToWord "Error Carga Actualizar Atributos", RutaEvidencias() &Num_Iter&"_"&"ErrorCargaActualizarAtributos"&".png"
				ExitActionIteration
			else
				Reporter.ReportEvent micPass, "Exito","La ventana 'Actualizar Atributo' correctamente"
			End If
			wait 1
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist))
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
		varsap=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		  	DataTable("s_Detalle", dtLocalSheet) = varsap
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Error_SAP_"&".png", True
			imagenToWord "Error SAP", RutaEvidencias() &Num_Iter&"_"&"Error_SAP_"&".png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
			SystemUtil.CloseProcessByName "javaw.exe"
			SystemUtil.CloseProcessByName "jp2launcher.exe"
			ExitTest
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Problema"&".png", True
		imagenToWord "No Hay Números Disponible", RutaEvidencias() &Num_Iter&"_"&"Problema"&".png"
		var=JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaObject("JPanel").GetROProperty("text")
		var= replace(var,"<html>","")
		var= replace(var,"</html>","")
		var= replace(var,"<br>","")
		var= replace(var,"?&#8203","")
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		DataTable("s_Detalle", dtLocalSheet) = var
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
		wait 1
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Proporcionar suscriptor").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Proporcionar suscriptor").JavaButton("Cerrar").Click
			wait 2
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").Close
			wait 2
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Click
			wait 2
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
				wait 1
			End If
			JavaWindow("Ejecutivo de interacción").JavaMenu("Archivo").JavaMenu("Salida").Select
		End If
		ExitTest
	End If
	
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
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Texto del motivo:").Set str_motivoalta
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Código de Centro Poblado").Set "0101010001"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaCheckBox("No tiene cobertura").Set "ON"
	wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Parametros_Alta_"&".png", True
	imagenToWord "Parametros de la Alta", RutaEvidencias() &Num_Iter&"_"&"Parametros_Alta_"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Siguiente >").Click
	
		tiempo=0
		Do
		tiempo=tiempo+1
			If tiempo>=180 Then
				DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)="La ventana 'Parametrización del Producto' no cargo"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaParametrizaelProducto"&".png" , True
				imagenToWord "Error Carga Parametriza el Producto", RutaEvidencias() &Num_Iter&"_"&"ErrorCargaParametrizaelProducto"&".png"
				ExitActionIteration
			else
				Reporter.ReportEvent micPass, "Exito", "La ventana 'Parametrización del Producto cargo correctamente'"
			End If
			wait 1
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Exist(1)
		
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Set
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaButton("Siguiente >").Click
	End If
		
		tiempo=0
		Do
		tiempo=tiempo+1
			If tiempo>=180 Then
				DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)="La ventana 'Negociar Configuración del Producto Móvil no cargo'"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Proporciopnar_Suscriptor_"&".png" , True
				imagenToWord "Proporcionar Suscriptor", RutaEvidencias() &Num_Iter&"_"&"Proporciopnar_Suscriptor_"&".png"
				ExitActionIteration
			else
				Reporter.ReportEvent micPass, "Exito", "La ventana 'Negociar Configuración del Producto Móvil' cargo correctamente"
			End If
			wait 1
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Buscar").Exist(1)
End Sub
Sub RecursosAlta()

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTab("JXTabbedPane").Select "Asignación de número"
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("TextFieldNative$1").Exist) = False
			wait 1
		Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("TextFieldNative$1").Set "6%%%%%%%%"
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("TextFieldNative$1").Set "920%%%%%%"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Proponer números").Click
	wait 2
	
		tiempo=0
			Do
				tiempo=tiempo+1
'				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Asignar número").Exist Then
'					varasig=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Asignar número").GetROProperty("enabled")
'				End If
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Distribuir Número").Exist Then
					varasig2=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Distribuir Número").GetROProperty("enabled")
				End If
				wait 1
		Loop  While Not ((varasig2="1") Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist))
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"NoHayNúmeroDisponible_"&".png", True
		imagenToWord "No Hay Números Disponible", RutaEvidencias() &Num_Iter&"_"&"NoHayNúmeroDisponible_"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
	  	DataTable("s_Detalle", dtLocalSheet) = "No hay ningún número devuelto desde Resource Manager System."
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
		wait 2
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").Close
			wait 2
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaMenu("Archivo").JavaMenu("Salida.").Select
		End If
		ExitTest
	End If
	wait 1
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Distribuir Número").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Distribuir Número").Click
		wait 1
	End If
'	
'	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Asignar número").Exist Then
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Asignar número").Click
'		wait 1
'	End If
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Número_Telefónico_"&".png", True
	imagenToWord "Número Telefónico", RutaEvidencias() &Num_Iter&"_"&"Número_Telefónico_"&".png"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("SearchJTable").Output CheckPoint("SearchJTable")
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTab("JXTabbedPane").Select "Configuración"			
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("Buscar por nombre:").Set "Tipo de SIM"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Buscar").Click
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
	While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
		wait 1
	Wend
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		wait 1
	End If
	
	Select Case DataTable("e_Tipo_Alta", dtLocalSheet)
		Case "Alta Nueva Equipo + SIM"
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
		Case "Alta Nueva solo SIM"
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
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
	While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
		wait 1
	Wend
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Mensajesdevalidación"&".png", True
		imagenToWord "Mensajes de validación", RutaEvidencias() &Num_Iter&"_"&"Mensajesdevalidación"&".png"
		wait 1
		varpag=	JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
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
			 While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
					wait 1
				Wend
		End If
	End If
	
	If DataTable("e_TipoServAdicional",dtLocalSheet)<>Empty Then
		Call ServiciosAdicionales()
		
	End If
	Call Financiamiento()
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"NPC"&".png", True
	imagenToWord "Negociar Configuración del Producto", RutaEvidencias() &Num_Iter&"_"&"NPC"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Click
	
		tiempo = 0
		Do
		tiempo = tiempo + 1
			If tiempo>=180 Then
		 	 	DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		  	 	DataTable("s_Detalle", dtLocalSheet) = "No se cargo la pantalla correctamente"
		   		 Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
		   		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaMetdosEntrega"&".png" , True
				imagenToWord "Error Carga Metdos Entrega", RutaEvidencias() &Num_Iter&"_"&"ErrorCargaMetdosEntrega"&".png"
			 	ExitActionIteration
			else
				Reporter.ReportEvent micPass,"OK","Continuar Flujo"
			End If
			wait 1
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("En tienda").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist))
		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist Then
			varprob=JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaObject("JPanel").GetROProperty("text")
			varprob= replace(varprob,"<html>","")
			varprob= replace(varprob,"</html>","")
			varprob= replace(varprob,"<br>","")
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = varprob
			Reporter.ReportEvent micWarning, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Error"&".png" , True
			imagenToWord "Error", RutaEvidencias() &Num_Iter&"_"&"Error"&".png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").Close
			wait 2
			ExitActionIteration
		End If
		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
			varmsgval = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = varmsgval
			Reporter.ReportEvent micFail , DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Error"&".png" , True
			imagenToWord "Error", RutaEvidencias() &Num_Iter&"_"&"Error"&".png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
			wait 4
			ExitActionIteration
		End If 
		
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
			If DataTable("e_ServAdicional2", dtLocalSheet)<>Empty Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Expand(nodeName2nd)
			End If
			
			Exit For
		End If
 	Next 
 	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").GetROProperty("items count")
	For Iterator = 0 To filas-1	 
		 nodeName3nd=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").GetItem(Iterator)
		 If nodeName3nd=DataTable("e_ServAdicional2", dtLocalSheet) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select(nodeName3nd)
			Exit For
		End If
 	Next 
 	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ServiciosAdicionales.png", True
	imagenToWord "Servicios Adicional seleccionado",RutaEvidencias() & "ServiciosAdicionales.png"
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").Click
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
	While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
		wait 1
	Wend
	
End Sub
Sub Financiamiento()
	
	IF ucase(str_Financiamiento) = "SI"  Then
	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Mostrar atributos:").Select "Obligatorio"
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("Buscar por nombre:").Set ""
		wait 1
		Select Case str_Cuotas
			Case 18
			 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "Corporativo-18 cuotas"
			     Set shell = CreateObject("Wscript.Shell")
			     shell.SendKeys "{RIGHT 100}"
			     JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
			     While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
					wait 1
					Wend
				     wait 1
		             JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Calcular").Click
		            While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
						wait 1
					Wend
			Case 12
			 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "Corporativo-12 cuotas"
			     Set shell = CreateObject("Wscript.Shell")
			     shell.SendKeys "{RIGHT 100}"
			     JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
			     While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
					wait 1
					Wend
				     wait 1
		             JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Calcular").Click
		            While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
						wait 1
					Wend
			Case 6
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "Corporativo-6 cuotas"
			     Set shell = CreateObject("Wscript.Shell")
			     shell.SendKeys "{RIGHT 100}"
			     JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
			     While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
					wait 1
				Wend
			     wait 1
	             JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Calcular").Click
	          	While JavaWindow("Ejecutivo de interacción").JavaObject("StatusBar$4").Exist = true
					wait 1
				Wend
				
			Case 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "Diferido-1 cuota"
			     Set shell = CreateObject("Wscript.Shell")
			     shell.SendKeys "{RIGHT 100}"
			     JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
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
	     JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Financiamiento.png", True
         imagenToWord "Plan de financiamiento",RutaEvidencias() & "Financiamiento.png"
	     
	ElseIf ucase(str_Financiamiento) = "NO" Then 
		wait 1
	     JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaList("Plan de Financiamiento:").Select "Contado"
	     Set shell = CreateObject("Wscript.Shell")
		 shell.SendKeys "{RIGHT 100}"
		 JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
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
	     imagenToWord "Plan al Contado",RutaEvidencias() & "FinanciamientoContado.png"
	End If
End Sub
Sub TipoEnvio()
	
	Select Case str_metodo_entrega
		Case "En tienda"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("En tienda").Set "ON"
				wait 1
		Case "Delivery"
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("Delivery").Set "ON"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png" , True
				imagenToWord "Canal de Venta", RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png"
				
					While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Lookup-Validated").Exist) = False
						wait 1
					Wend
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Lookup-notValidated_2").Click
				wait 1
					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Exist) = False
						wait 1
					Wend
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaTable("SearchJTable").SelectRow "#0"
				wait 1
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DireccionEnvio"&".png" , True
				imagenToWord "Dirección de Envio", RutaEvidencias() &Num_Iter&"_"&"DireccionEnvio"&".png"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Click
					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Exist) = False
						wait 1
					Wend
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Buscar detalles de contacto").Click
				wait 1
					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Exist) = False
						wait 1
					Wend
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaList("ComboBoxNative$1").Select DataTable ("e_docuDelivery",dtLocalSheet)
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaEdit("TextFieldNative$1").Set DataTable ("e_numDocuDelivery",dtLocalSheet)
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Buscar ahora").Click
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaTable("SearchJTable").SelectRow "#0"
				wait 1
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DireccionReceptor"&".png" , True
				imagenToWord "Direccion de Receptor", RutaEvidencias() &Num_Iter&"_"&"DireccionReceptor"&".png"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Click
				wait 1
					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Exist) = False
						wait 1
					Wend
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Set "QA"
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Número de teléfono del").Set "999999999"
				wait 1
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"EntregaDelivery"&".png", True
				imagenToWord "Entrega Delivery", RutaEvidencias() &Num_Iter&"_"&"EntregaDelivery"&".png"
		Case "Recojo en otra tienda"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("Recojo en otra tienda").Set "ON"
			wait 1
	End Select
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png" , True
	imagenToWord "Canal de Venta", RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Siguiente >").Click
	wait 1
	
		While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaButton("Pago inmediato").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist)) = False
			wait 1
		Wend
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
	    var1 = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
	    var1 = replace(var1,"<html>","")
		var1 = replace(var1,"</html>","")
	   ' vardispo2 = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(1,1)
	    Reporter.ReportEvent micFail, "Fallido", var1
	    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MensajeValidacion"&".png" , True
		imagenToWord "Mensaje de Validación", RutaEvidencias() &Num_Iter&"_"&"MensajeValidacion"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
		ExitActionIteration
	End If 

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
		varmsg=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		DataTable("s_Detalle", dtLocalSheet) = varmsg
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Error"&".png" , True
		imagenToWord "Error", RutaEvidencias() &Num_Iter&"_"&"Error"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
		wait 2
		ExitActionIteration
	End If

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
	    vardispo1 = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
	    vardispo2 = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(1,1)
	    Reporter.ReportEvent micFail, "Fallido", vardispo2
	    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Error"&".png" , True
		imagenToWord "Error", RutaEvidencias() &Num_Iter&"_"&"Error"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		ExitActionIteration
	End If 
	
End  Sub
Sub PagoInmediato()
	
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").Exist=False
		wait 1
	Wend
	
	tiempo=0
	While(vartexto<>"") = False
		wait 1
		tiempo=tiempo+1
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").Exist Then
			vartexto=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").GetROProperty("text")
		End If
		If(tiempo>=160) Then
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No carga el Nombre y Dirección de Facturación"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Cerrar").Click
			While(JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Exist)=False
				wait 1
			Wend
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Click
			While(JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Distribución").JavaList("Motivo:").Exist ) = False
				wait 1
			Wend
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Distribución").JavaList("Motivo:").Select "Pedido de Cliente"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Distribución").JavaButton("Aceptar").Click
			wait 1
			While(JavaDialog("Cierre del formulario").JavaButton("Descartar").Exist) Or(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Exist)=False
				wait 1
			Wend
			If JavaDialog("Cierre del formulario").JavaButton("Descartar").Exist Then
				JavaDialog("Cierre del formulario").JavaButton("Descartar").Click
				wait 1
			End If
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Click
				wait 1
			End If
			ExitActionIteration
		End If
	Wend
	
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaRadioButton("Nuevo").Set
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaList("Mostrar:").Select "Acciones de orden activas "
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ValidaciónAcuerdoFacturación.png", True
	imagenToWord "Validación Acuerdo Facturación",RutaEvidencias() & "ValidaciónAcuerdoFacturación.png"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Siguiente >").Click	
	wait 1
	
		tiempo = 0
		Do
		tiempo = tiempo + 1
		If tiempo>=180 Then
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No cargo la pantalla Neogicar Pago"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaNegociarPago"&".png" , True
			imagenToWord "Error Carga Negociar Pago", RutaEvidencias() &Num_Iter&"_"&"ErrorCargaNegociarPago"&".png"
			ExitActionIteration
		else
		Reporter.ReportEvent micPass,"OK","Continuar Flujo"
		End If
		wait 1
	Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaButton("Pago inmediato").Exist
	wait 1
	
	If str_Financiamiento = "SI" Then
		tiempo=0
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaEdit("Importe de Cuota Mayor:").Exist) = False
			wait 1
			tiempo = tiempo + 1
			If (tiempo >= 180) Then
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorCargaFinanciamiento_"&Num_Iter&".png", True
				imagenToWord "Error Carga Financiamiento_"&Num_Iter,RutaEvidencias() & "ErrorCargaFinanciamiento_"&Num_Iter&".png"
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargó la ventana -Pago Inmediato Financimiento- de manera correcta"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			End If
		Wend
		wait 1
		varimporte=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaEdit("Importe de Cuota Inicial:").GetROProperty("enabled")
		If varimporte="1" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaEdit("Importe de Cuota Inicial:").Set 100
		End If
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaEdit("Importe de Cuota Mayor:").Set 1
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
	
	tiempo=0
	While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Exist)  or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaStaticText("'Seleccione la casilla").Exist))= False
		wait 1
		tiempo = tiempo + 1
		If (tiempo >= 180) Then
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
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Select "Factura"
	wait 1
	
	var8 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Numero RUC").GetROProperty("text")
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Numero RUC").GetROProperty("text") = "" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Tipo de documento:").Select "Boleta"
	End If
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Select str_MedioPago
	If str_MedioPago = "Pago a la Factura" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Cantidad de cuotas:").Select str_Cuotas
		wait 1
	    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Calcular").Click
	End If
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ProcesarPagoInmediato.png", True
	imagenToWord "Procesar Pago Inmediato",RutaEvidencias() & "ProcesarPagoInmediato.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Enviar").Click
	
	tiempo=0
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 599306A").JavaButton("Pago inmediato").Exist) = true
		wait 1
		tiempo = tiempo + 1
		If (tiempo >= 180) Then
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
	
	tiempo = 0 
	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist) = False
		Wait 1
		
		tiempo = tiempo + 1
		If (tiempo >= 180) Then
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

	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist = False
		wait 1
	Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Click
	wait 2
	
	While JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist = False
	 	wait 1
	Wend
	 If JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist= True Then
 	 	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "LinkDocu.png", True
		imagenToWord "LinkDeDocumentación_"&Num_Iter,RutaEvidencias() & "LinkDocu.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Close
		wait 2
		Dim var
		var = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").GetROProperty("enabled")
		If var = "1"  Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").Set "ON"
		End If
	End If

	tiempo = 0
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
		
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
		wait 1
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaStaticText("Por favor valide el Código").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
			wait 1
			Call CodigoVendedor()
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
		End If
	End If

	While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist)) = False
		wait 1
	Wend
	wait 1
	DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").GetROProperty("text")
	flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
	DataTable("s_Nro_Orden", dtLocalSheet) = replace (DataTable("s_Nro_Orden", dtLocalSheet),"Orden ","")
	Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png", True
	imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2272771A").JavaButton("Cerrar").Click
	wait 1
	
	If str_metodo_entrega="Delivery" Then
		ExitActionIteration
	End If
	
End Sub
Sub CodigoVendedor()

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validar").Click
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").JavaEdit("TextFieldNative$1").Set "41523813"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").JavaButton("Buscar ahora").Click
	
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").JavaTable("SearchJTable").Exist = False
		wait 1
	Wend
	wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").JavaTable("SearchJTable").SelectRow "#0"
    wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").JavaButton("Seleccionar").Click
	wait 1
	
End Sub
Sub PagoManual()

	If (str_mediopago<>"Externo") Then
		wait 1
			tiempo=0
			Do 
				tiempo=tiempo+1
				vardepo=JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").GetROProperty("enabled")
				wait 2
			Loop While Not (vardepo="1")
		
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select
		wait 1
		
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Exist) = False
			wait  1
		Wend
		wait 1
		
		Do
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Select "Tareas pendientes del equipo"
			wait 2
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Finalizar compra y activar").Exist
		wait 2
	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").SetFocus
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
		wait 2
		
		tiempo=0
		Do 
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
				nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("-- Registros").GetROProperty("attached text")
				tiempo=tiempo+1
				wait 1
			End If
			If (tiempo >= 120) Then
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
		
		var = JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").GetROProperty("enabled")
			tiempo=0
				While(var <> "0") = False	
					wait 1
					tiempo=tiempo+1
					If (tiempo >= 180) Then
							JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").Close
							DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				  			DataTable("s_Detalle", dtLocalSheet) = "La orden: "&DataTable("s_Nro_Orden", dtLocalSheet)&" ya fue pagado"
							Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
							JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").Close
							ExitActionIteration
					End If
				Wend
		
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").Click
		Reporter.ReportEvent micDone, "Pago Correcto", "El número de orden : "&DataTable("s_Nro_Orden", dtLocalSheet)&" fué correctamente pagado"
		
	End  If

End Sub
Sub GestionLogistica()
		
		tiempo=0
			Do
				varlogis=JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").GetROProperty("enabled")
				wait 2
			Loop While Not (varlogis="1")

		wait 1
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
		wait 1
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Exist)=False
			wait 1
		Wend
		wait 2
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
			If (tiempo >= 120) Then
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
					
						If str_tipoalta="Alta Nueva Equipo + SIM" Then
								vardisp=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").GetCellData (1,1)
								Do
									vardisp=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").GetCellData (1,1)
									wait 3
									If vardisp = "Tarjeta SIM" Then
										Exit do 
									End If
								Loop While Not vardisp ="Dispositivo"
					
								If vardisp="Dispositivo" Then
								
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#1","#4"
									Set shell = CreateObject("Wscript.Shell") 
									shell.SendKeys "{ENTER}"
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#1","#4",DataTable("e_ID_Dispositivo", dtLocalSheet)
									wait 2
								Else 
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#1","#4"
									Set shell = CreateObject("Wscript.Shell") 
									shell.SendKeys "{ENTER}"
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#1","#4",DataTable("e_ID_SIM", dtLocalSheet)
									wait 2
									
								End If
					
								varsim=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").GetCellData (2,1)
								
								If varsim="Tarjeta SIM" Then
								
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#2","#4"
									Set shell = CreateObject("Wscript.Shell") 
									shell.SendKeys "{ENTER}"
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#2","#4",DataTable("e_ID_SIM", dtLocalSheet)
									wait 2
								Else 
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#2","#4"
									Set shell = CreateObject("Wscript.Shell") 
									shell.SendKeys "{ENTER}"
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#2","#4",DataTable("e_ID_Dispositivo", dtLocalSheet)
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
					wait 1
					JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Validar y Crear Factura").Object.doClick()
					 wait 1
					
					tiempo = 0
					Do
						tiempo=tiempo+1
							varhab=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").GetROProperty("enabled")					
							wait 1
					Loop While Not((JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (varhab="1"))
				
						'If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist(1) or Window("Ejecutivo de interacción").Window("Buscar: Orden > Solicitar").Window("Mensaje").Exist(1) or JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(1) Then
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
				     			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MensajeValidacion"&".png", True
								imagenToWord "Mensaje de Validación", RutaEvidencias() &Num_Iter&"_"&"MensajeValidacion"&".png"
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
								If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist(1) Then
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
Sub EmpujeOrden()
	
		If DataTable("e_Tipo_De_DATA_Sim", dtLocalSheet) = "DATA LOGICA" Then
		
			tiempo=0
			Do 
				tiempo=tiempo+1
				vardepo=JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").GetROProperty("enabled")
				wait 2
			Loop While Not (vardepo="1")
		
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
			JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select
			wait 1
			While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Exist) = False
				wait 1
			Wend
			
			
			Do
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Select "Tareas pendientes del equipo"
				wait 2
			Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Finalizar compra y activar").Exist
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
					nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("-- Registros").GetROProperty("text")
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
			wait 1
'			
'			tiempo=0
'			Do
'				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
'					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
'					wait 2
'					tiempo = tiempo+1
'					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").Output CheckPoint("Equipo usuario:")
'					varValidaRespuestaCumplimiento = Environment("s_ValidaManejarRespuestaCumplimiento")
'					wait 1
'				End If
'					If (tiempo >= 120) Then
'						DataTable("s_Resultado",dtLocalSheet)="Fallido"
'						DataTable("s_Detalle",dtLocalSheet)="La actividad 'Manejar Respuesta de Cumplimiento' no cargo"	
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
'						ExitTestIteration
'					End If 
'			Loop While Not varValidaRespuestaCumplimiento = "Manejar Respuesta de Cumplimiento"
'			wait 2
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
			JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").Click
			While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").Exist) = False
				wait  1
			Wend
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
		End If
End Sub
Sub OrdenCerrado()

		tiempo=0
			Do
				tiempo=tiempo+1
				varlogis=JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").GetROProperty("enabled")
				wait 2
			Loop While Not (varlogis="1")
			
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
		wait 1
		
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Exist) = False
			wait 1
		Wend

		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").SetFocus
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaCheckBox("Solo órdenes pendientes").Set "OFF"
		wait 1
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
		wait 8
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:")
		Reporter.ReportEvent micPass,"Se valida el estado de la orden",  DataTable("s_ValEstadoOrden", dtLocalSheet)
		
		tiempo = 0
		Do
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:")
			tiempo = tiempo + 1
				If (tiempo>=150) Then		
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
	
	filas=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 713979A").JavaTable("SearchJTable").GetROProperty("rows")
	For Iterator = 0 To filas-1
		varselec=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 713979A").JavaTable("SearchJTable").GetCellData(Iterator,0)
	Next
	
	If varselec<>"Cerrar Acción de Orden" Then
	 	DataTable("s_Resultado",dtLocalSheet)="Fallido"
		DataTable("s_Detalle",dtLocalSheet)="La orden "&DataTable("s_Nro_Orden",dtLocalSheet)&" culmino en la Actividad "&varselec&""
		Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 713979A").JavaButton("Cancelar").Click

		wait 2
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
			wait 2
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(1) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
		End If
		ExitActionIteration
		wait 1
	End If
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 713979A").JavaButton("Cancelar").Click

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
		wait 2
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
	End If
	
End Sub






