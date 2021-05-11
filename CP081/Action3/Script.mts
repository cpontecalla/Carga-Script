Dim e_Tipo1, e_Tipo2, e_Tipo3, e_IdentCuenta, filas, ruc
Dim habilitado

e_Tipo1 				= DataTable("e_Tipo1", dtLocalSheet)
e_Tipo2 				= DataTable("e_Tipo2", dtLocalSheet)
e_Tipo3 				= DataTable("e_Tipo3", dtLocalSheet)
e_IdentCuenta			= DataTable("e_IdentCuenta", dtLocalSheet)

'Bucles que esperan la carga de la pantalla Busqueda de contacto y suscripción.

	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaButton("-- Registros").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de Busqueda de contacto"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaBusquedaContacto.png", True
			imagenToWord "Error en la Carga de Busqueda de contacto",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaBusquedaContacto.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaButton("Buscar ahora").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de Busqueda de contacto"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaBusquedaContacto.png", True
			imagenToWord "Error en la Carga de Busqueda de contacto",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaBusquedaContacto.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
			ExitActionIteration
		End If	
	Wend
	
	JavaWindow("Ejecutivo de interacción").JavaMenu("Crear").JavaMenu("Soporte").JavaMenu("Caso").Select


	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaList("Prioridad:").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Crear Caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Encontrar Comunicante").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Crear Caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaEdit("Título del caso").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Crear Caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
	Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaEdit("ID de la cuenta:").SetFocus()
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaEdit("ID de la cuenta:").Set e_IdentCuenta
	wait 3
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"EncComunicante.png", True
	imagenToWord "Busqueada cuenta Corporativa ",RutaEvidencias() &Num_Iter&"_"&"EncComunicante.png"
	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Encontrar Comunicante").Click

'	t=0
'	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Búsqueda:").JavaButton("Buscar ahora").Exist = False
'		Wait 1
'		
'		t = t + 1
'		If (t >= 180) Then
'			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
'			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Buscar Contacto"
'			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaBuscarContacto.png", True
'			imagenToWord "Error en la Carga de la Pantalla Buscar Contacto",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaBuscarContacto.png"
'			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
'			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
'			End If
'			Wait 2
'			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Búsqueda:").JavaButton("Cerrar").Exist = True Then
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Búsqueda:").JavaButton("Cerrar").Click
'			End If
'			ExitActionIteration
'		End If	
'	Wend
'	t=0
'	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Búsqueda:").JavaTable("SearchJTable").Exist = False
'		Wait 1
'		
'		t = t + 1
'		If (t >= 180) Then
'			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
'			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Buscar Contacto"
'			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaBuscarContacto.png", True
'			imagenToWord "Error en la Carga de la Pantalla Buscar Contacto",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaBuscarContacto.png"
'			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
'			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
'			End If
'			Wait 2
'			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Búsqueda:").JavaButton("Cerrar").Exist = True Then
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Búsqueda:").JavaButton("Cerrar").Click
'			End If
'			ExitActionIteration
'		End If	
'	Wend
'
'	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Búsqueda:").JavaTable("SearchJTable").GetROProperty("rows")
'	t=0
'	While filas <=  0
'		Wait 1
'		
'		t = t + 1
'		If (t >= 180) Then
'			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
'			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
'			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
'			imagenToWord "Error en la Carga de la Pantalla Resultados de contacto",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
'			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
'			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
'			End If
'			Wait 2
'			ExitActionIteration
'		End If	
'		filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Búsqueda:").JavaTable("SearchJTable").GetROProperty("rows")
'	Wend
'	
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Búsqueda:").JavaTable("SearchJTable").SelectRow "#0"
'	wait 3
'	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"BuscarContacto.png", True
'	imagenToWord "Buscar Contacto",RutaEvidencias() &Num_Iter&"_"&"BuscarContacto.png"
'	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
''	
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Búsqueda:").JavaButton("Seleccionar").Click
'
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaList("Prioridad:").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Crear Caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Encontrar Comunicante").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Crear Caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaEdit("Título del caso").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Crear Caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
	Wend
	wait 3
	
	habilitado = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("DetailsButton_on").GetROProperty("enabled")
	While habilitado = "0"
		habilitado = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("DetailsButton_on").GetROProperty("enabled")
		wait 1 
	Wend
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("DetailsButton_on").Click
	wait 3
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Cuenta:").JavaEdit("ID de Compańía").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Crear Caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			
			Wait 2
			ExitActionIteration
		End If	
	Wend
	ruc = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Cuenta:").JavaEdit("ID de Compańía").GetROProperty("text")
	wait 3
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Cuenta:").JavaButton("Cerrar").Click
	wait 3
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaList("Medio de notificación:").Select "Carta"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaList("Tipo 1:").Select e_Tipo1
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaList("Tipo 2:").Select e_Tipo2
	wait 2

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaList("Tipo 3:").Select e_Tipo3
	wait 2
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SeleccionTipos.png", True
	imagenToWord "Se seleccionan los tipos del Caso",RutaEvidencias() &Num_Iter&"_"&"SeleccionTipos.png"
	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaTab("Número de Suscripción:").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Crear Caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			
			Wait 2
			ExitActionIteration
		End If	
	Wend
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaTab("Número de Suscripción:").Select "Suscripción"

	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaTable("Número de Suscripción:").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Crear Caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			
			Wait 2
			ExitActionIteration
		End If	
	Wend @@ hightlight id_;_8416259_;_script infofile_;_ZIP::ssf2.xml_;_
	
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Agregar").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Crear Caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			
			Wait 2
			ExitActionIteration
		End If	
	Wend
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"AggSuscripcion.png", True
	imagenToWord "Se agregan las Suscripciones",RutaEvidencias() &Num_Iter&"_"&"AggSuscripcion.png"
	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Agregar").Click
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Lista de").JavaTable("SearchJTable").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaSelSuscrip.png", True
			imagenToWord "Error en la Carga de la Pantalla Crear Caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaSelSuscrip.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			
			Wait 2
			ExitActionIteration
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Lista de").JavaButton("10 Registros").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Crear Caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			
			Wait 2
			ExitActionIteration
		End If	
	Wend
	
	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Lista de").JavaTable("SearchJTable").GetROProperty("rows")
	t=0
	While filas <=  0
		Wait 1
		
		t = t + 1
		If (t >= 30) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
		filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Lista de").JavaTable("SearchJTable").GetROProperty("rows")
	Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Lista de").JavaEdit("TextFieldNative$1").Set "Activo"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Lista de").JavaButton("Buscar ahora").Click
'	wait 1
'	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Lista de").JavaTable("SearchJTable").GetROProperty("rows")
'	t=0
'	While filas <=  0
'		Wait 1
'		
'		t = t + 1
'		If (t >= 180) Then
'			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
'			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
'			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
'			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
'			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
'			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
'			End If
'			Wait 2
'			ExitActionIteration
'		End If	
'		filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Lista de").JavaTable("SearchJTable").GetROProperty("rows")
'	Wend
'	
'	wait 2
'	For Iterator = 0 To 4 Step 1
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Lista de").JavaTable("SearchJTable").SelectRow "#"&Iterator
'		wait 2
'	Next
'	
'	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SelSuscripcion.png", True
'	imagenToWord "Selección de Suscripción",RutaEvidencias() &Num_Iter&"_"&"SelSuscripcion.png"
'	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Lista de").JavaButton("Seleccionar").Click
	wait 2
	
	t =0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaTable("Número de Suscripción:").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Crear Caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
	Wend
	t=0

	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Agregar").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Crear Caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Próximo caso").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Crear Caso",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCrearCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
	Wend
	
	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaTable("Número de Suscripción:").GetROProperty("rows")
	t=0
	While filas <= 0
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
		filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaTable("Número de Suscripción:").GetROProperty("rows")
	Wend
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SuscripcionSelec.png", True
	imagenToWord "Suscripciones Seleccionadas",RutaEvidencias() &Num_Iter&"_"&"SuscripcionSelec.png"
	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
	wait 2
'	For Iterator = 0 To 4 Step 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaTable("Número de Suscripción:").SelectRow "#0"
'		wait 2
'	Next
	
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ElimSuscripcion.png", True
	imagenToWord "Eliminamos Suscripción",RutaEvidencias() &Num_Iter&"_"&"ElimSuscripcion.png"
	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Eliminar").Click
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SuscripcionesCaso.png", True
	imagenToWord "Suscripciones del caso",RutaEvidencias() &Num_Iter&"_"&"SuscripcionesCaso.png"
	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
	wait 2 
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaList("ComboBoxNative$1").Select "Guardar y Continuar"
	

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Guardar").Click

	Select Case e_Tipo1
		Case "Solicitud Medicion de señal"
			t=0
			While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").Exist = False
				Wait 1
				
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtlocalSheet) = "Fallido"
					DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Atributos Flexibles"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
					imagenToWord "Error en la Carga de la Pantalla Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
					End If
					Wait 2
					ExitActionIteration
				End If	
			Wend
			t=0
			While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaButton("Guardar").Exist = False
				Wait 1
				
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtlocalSheet) = "Fallido"
					DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Atributos Flexibles"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
					imagenToWord "Error en la Carga de la Pantalla Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
					End If
					Wait 2
					ExitActionIteration
				End If	
			Wend
			filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").GetROProperty("rows")
			t=0
			While filas <= 0
				Wait 1
				
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtlocalSheet) = "Fallido"
					DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
					imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
					End If
					Wait 2
					ExitActionIteration
				End If	
				filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").GetROProperty("rows")
			Wend
	
	
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#0","#1","GC3 - Empresas" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf7.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#1","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#1","#1",ruc @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf8.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#2","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#2","#1","AROMAS" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf9.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#3","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#3","#1","SI" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf10.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#4","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#4","#1","PRUEBA" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf11.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#5","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#5","#1","NO TIENE" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf12.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#6","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#6","#1","925012456" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf13.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#7","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#7","#1","prueba@gmail.com" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf14.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#8","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#8","#1","10:00 AM" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf14.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"AtributosFlexibles.png", True
			imagenToWord "Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"AtributosFlexibles.png"
			Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaButton("Guardar").Click
		Case "Gestión Averías Negocios"
			t=0
			While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").Exist = False
				Wait 1
				
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtlocalSheet) = "Fallido"
					DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Atributos Flexibles"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
					imagenToWord "Error en la Carga de la Pantalla Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
					End If
					Wait 2
					ExitActionIteration
				End If	
			Wend
			t=0
			While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaButton("Guardar").Exist = False
				Wait 1
				
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtlocalSheet) = "Fallido"
					DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Atributos Flexibles"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
					imagenToWord "Error en la Carga de la Pantalla Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
					End If
					Wait 2
					ExitActionIteration
				End If	
			Wend
			filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").GetROProperty("rows")
			t=0
			While filas <= 0
				Wait 1
				
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtlocalSheet) = "Fallido"
					DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
					imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
					End If
					Wait 2
					ExitActionIteration
				End If	
				filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").GetROProperty("rows")
			Wend
	
	
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#0","#1","No" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf7.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#1","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#1","#1","Alta Nueva" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf8.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#2","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#2","#1","100" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf9.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#3","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#3","#1","Nuevos Soles" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf10.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#4","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#4","#1","Postpago" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf11.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#5","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#5","#1","01-06-2019" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf12.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#6","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#6","#1","Lima" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf13.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#7","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#7","#1","920951245" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf14.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#8","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#8","#1","920954178" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf14.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"AtributosFlexibles.png", True
			imagenToWord "Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"AtributosFlexibles.png"
			Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaButton("Guardar").Click
			
		Case "Gestión Averías Empresas"
			t=0
			While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").Exist = False
				Wait 1
				
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtlocalSheet) = "Fallido"
					DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Atributos Flexibles"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
					imagenToWord "Error en la Carga de la Pantalla Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
					End If
					Wait 2
					ExitActionIteration
				End If	
			Wend
			t=0
			While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaButton("Guardar").Exist = False
				Wait 1
				
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtlocalSheet) = "Fallido"
					DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Atributos Flexibles"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
					imagenToWord "Error en la Carga de la Pantalla Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
					End If
					Wait 2
					ExitActionIteration
				End If	
			Wend
			filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").GetROProperty("rows")
			t=0
			While filas <= 0
				Wait 1
				
				t = t + 1
				If (t >= 180) Then
					DataTable("s_Resultado", dtlocalSheet) = "Fallido"
					DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
					imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
					End If
					Wait 2
					ExitActionIteration
				End If	
				filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").GetROProperty("rows")
			Wend
	
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#0","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#0","#1","POSTPAGO" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf7.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#1","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#1","#1","No" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf8.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#2","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#2","#1","920959999" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf9.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#3","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#3","#1","920154783" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf10.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#4","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#4","#1","XXXXXXXXXXX" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf11.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#5","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#5","#1","lima" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf12.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#6","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#6","#1","Lima" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf13.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#7","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#7","#1","920951245" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf14.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#8","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#8","#1","920954178" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf14.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").DoubleClickCell "#9","#1"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaTable("SearchJTable").SetCellData "#9","#1","Alta Nueva" @@ hightlight id_;_14232694_;_script infofile_;_ZIP::ssf14.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"AtributosFlexibles.png", True
			imagenToWord "Atributos Flexibles",RutaEvidencias() &Num_Iter&"_"&"AtributosFlexibles.png"
			Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso  > Atributos").JavaButton("Guardar").Click
		
	End Select
	
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Caso:2547").JavaEdit("ID del caso:").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Caso:2547").JavaList("Estado").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
	Wend
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CreacionCaso.png", True
	imagenToWord "Creación del caso",RutaEvidencias() &Num_Iter&"_"&"CreacionCaso.png"
	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
	Caso = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Caso:2547").JavaEdit("ID del caso:").GetROProperty("text")
	DataTable("s_CasoCreado", dtLocalSheet) = Caso
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Caso:2547").JavaButton("Guardar").Click
	wait 3


	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Soporte").JavaMenu("Casos").Select @@ hightlight id_;_8516623_;_script infofile_;_ZIP::ssf4.xml_;_
	
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaButton("Buscar ahora").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
	Wend
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaEdit("TextFieldNative$1").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
	Wend
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaEdit("TextFieldNative$1").Set Caso
wait 2

JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaButton("Buscar ahora").Click

filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaTable("SearchJTable").GetROProperty("rows")
t=0
	While filas <= 0
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
		filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaTable("SearchJTable").GetROProperty("rows")
	Wend
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"BuscarCaso.png", True
	imagenToWord "Buscar caso",RutaEvidencias() &Num_Iter&"_"&"BuscarCaso.png"
	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaTable("SearchJTable").DoubleClickCell "#0","#0"	

	wait 2
	
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547").JavaEdit("Condición:").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
	Wend
	wait 3
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"VerCaso.png", True
	imagenToWord "Ver caso",RutaEvidencias() &Num_Iter&"_"&"VerCaso.png"
	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547").JavaTab("Id de TT externo:").Select "Suscripción" @@ hightlight id_;_25351375_;_script infofile_;_ZIP::ssf5.xml_;_
	

	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547").JavaTable("SearchJTable").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
	Wend
'	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547").JavaTable("SearchJTable").GetROProperty("rows")
'	t=0
'	While filas <= 0
'		Wait 1
'		
'		t = t + 1
'		If (t >= 180) Then
'			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
'			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
'			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
'			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
'			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
'			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
'			End If
'			Wait 2
'			ExitActionIteration
'		End If	
'		filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547").JavaTable("SearchJTable").GetROProperty("rows")
'	Wend
'	
'	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547").JavaTable("SearchJTable").GetROProperty("rows")
'	wait 2
'	For Iterator = 0 To filas - 1 Step 1
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547").JavaTable("SearchJTable").SelectRow "#"&Iterator
'	Next
'	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SelecSuscr.png", True
'	imagenToWord "Seleccionar Suscripciones para eliminar",RutaEvidencias() &Num_Iter&"_"&"SelecSuscr.png"
'	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
'	
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547").JavaButton("Eliminar").Click
'	wait 3
'	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SinSuscripciones.png", True
'	imagenToWord "Se visualiza el caso sin Suscripciones",RutaEvidencias() &Num_Iter&"_"&"SinSuscripciones.png"
'	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN}"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547").JavaButton("Agregar").Click
'	
'	
'	
'	
'	t=0
'	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547 > Lista").JavaTable("SearchJTable").Exist = False
'		Wait 1
'		
'		t = t + 1
'		If (t >= 180) Then
'			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
'			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
'			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
'			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
'			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
'			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
'			End If
'			Wait 2
'			ExitActionIteration
'		End If	
'	Wend
'	t=0
'	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547 > Lista").JavaButton("Buscar ahora").Exist = False
'		Wait 1
'		
'		t = t + 1
'		If (t >= 180) Then
'			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
'			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
'			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
'			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
'			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
'			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
'			End If
'			Wait 2
'			ExitActionIteration
'		End If	
'	Wend
'	t=0
'	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547 > Lista").JavaButton("10 Registros").Exist = False
'		Wait 1
'		
'		t = t + 1
'		If (t >= 180) Then
'			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
'			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
'			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
'			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
'			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
'			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
'			End If
'			Wait 2
'			ExitActionIteration
'		End If	
'	Wend
'	
	
	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547 > Lista").JavaTable("SearchJTable").GetROProperty("rows")
	t=0
	While filas <= 0
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
		filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547 > Lista").JavaTable("SearchJTable").GetROProperty("rows")
	Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547 > Lista").JavaEdit("TextFieldNative$1").Set "Activo"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547 > Lista").JavaButton("Buscar ahora").Click
	wait 4
'	wait 2
'	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547 > Lista").JavaTable("SearchJTable").GetROProperty("rows")
'	t=0
'	While filas <= 0
'		Wait 1
'		
'		t = t + 1
'		If (t >= 180) Then
'			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
'			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
'			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
'			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
'			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
'			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
'			End If
'			Wait 2
'			ExitActionIteration
'		End If	
'		filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547 > Lista").JavaTable("SearchJTable").GetROProperty("rows")
'	Wend
'	For Iterator = 0 To 4 Step 1
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547 > Lista").JavaTable("SearchJTable").SelectRow "#0"
'		wait 2
'	Next
	
'	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"SeleccionSuscripciones2.png", True
'	imagenToWord "Se visualiza las Suscripciones seleccionadas",RutaEvidencias() &Num_Iter&"_"&"SeleccionSuscripciones2.png"
'	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
'	
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547 > Lista").JavaButton("Seleccionar").Click
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547 > Lista").JavaTable("SearchJTable").SelectRow("#0")
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CasoConSuscripciones.png", True
	imagenToWord "Se visualiza el caso con las Suscripciones seleccionadas",RutaEvidencias() &Num_Iter&"_"&"CasoConSuscripciones.png"
	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
	
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547").JavaTab("Id de TT externo:").Select "Más información" @@ hightlight id_;_25351375_;_script infofile_;_ZIP::ssf6.xml_;_
	
	t=0
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547").JavaEdit("Dueńo:").Exist = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "Error en la Carga de la Pantalla Crear Caso"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png", True
			imagenToWord "Error en la Carga de la Pantalla Caso Guardado",RutaEvidencias() &Num_Iter&"_"&"ErrorCargaCaso.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Exist = true Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Caso").JavaButton("Cerrar").Click
			End If
			Wait 2
			ExitActionIteration
		End If	
	Wend
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"VisualizarCasoFinal.png", True
	imagenToWord "Se visualiza el caso con los cambios realizados",RutaEvidencias() &Num_Iter&"_"&"VisualizarCasoFinal.png"
	Reporter.ReportEvent micPass, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547").JavaButton("Guardar").Click
	wait 5
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547").JavaButton("Cerrar").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver caso: 2547").JavaButton("Cerrar").Click
	End If
	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaButton("Cerrar").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Resultados de").JavaButton("Cerrar").Click
	End If
	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Caso:2556 > Atributos").JavaButton("Cerrar").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Caso:2556 > Atributos").JavaButton("Cerrar").Click
	End If

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Caso:2547").JavaButton("Cerrar").Exist = True Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Caso:2547").JavaButton("Cerrar").Click
	End If

	


