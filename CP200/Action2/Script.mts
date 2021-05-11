Dim IdServiciol, OrdenPendiente, WIC_Activa, Equipo, Agrega, Orden,str_idDispositivo, Motivo, Tipo, PrecioEq, Plan


IdServicio 			= DataTable("e_IdServicio", dtLocalSheet)


Call BusquedaIdServicio()
Call ValidacionCampo()

Sub BusquedaIdServicio()
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaEdit("Nombre completo:").Exist = False
		wait 1
	Wend
	Dim nombre
	nombre = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaEdit("Nombre completo:").GetROProperty("text")
	While nombre = ""
		wait 1
		nombre = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaEdit("Nombre completo:").GetROProperty("text")
	Wend
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"PanelInteraccion.png", True
	imagenToWord "Visualización Panel de Interacción",RutaEvidencias() &"PanelInteraccion.png"
	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTab("Acciones rápidas").Select "Suscripciones" @@ hightlight id_;_8326975_;_script infofile_;_ZIP::ssf1.xml_;_
	wait 2
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaTable("Acciones rápidas").Exist = false
		wait 1
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
Sub ValidacionCampo()

	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaTab("Antigüedad de línea:").Exist = False
		wait 1
	Wend @@ hightlight id_;_16401507_;_script infofile_;_ZIP::ssf7.xml_;_
	wait 1
	Dim value
	value = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaEdit("Dirección de instalación:").GetROProperty("editable")
	
	If value = "0"  Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"CampoValidado.png", True
		imagenToWord "Se valida el campo: Dirección de Instalación se encuentra deshabilitado",RutaEvidencias() &"CampoValidado.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaEdit("Dirección de instalación:").CaptureBitmap RutaEvidencias() &"Campo.png", True
		imagenToWord "Se valida el campo: ",RutaEvidencias() &"Campo.png"
		DataTable("s_Resultado", dtlocalSheet) = "Exito"
		DataTable("s_Detalle", dtlocalSheet) = "Se valida que el campo: Dirección de Instalacion se encuentra deshabilitado"
		Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
	Else 
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"CampoValidado.png", True
		imagenToWord "El campo: Dirección de Instalación se encuentra habilitado",RutaEvidencias() &"CampoValidado.png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Detalles del producto").JavaEdit("Dirección de instalación:").CaptureBitmap RutaEvidencias() &"Campo.png", True
		imagenToWord "Se valida el campo: ",RutaEvidencias() &"Campo.png"
		DataTable("s_Resultado", dtlocalSheet) = "Fallido"
		DataTable("s_Detalle", dtlocalSheet) = "El campo Dirección de Instalacion se encuentra habilitado"
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
		
	End If
	

	
End Sub

