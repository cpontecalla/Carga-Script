
Dim val1, val2, val3, val4, val5, val6, str_dia, Num_Iter, str_TipoCliente, str_TipoDoc, str_NumDoc, intStartTime, intStopTime, var1, var2, str_Genero,  str_FechaNac, str_Nombres, str_Apellidos, str_Dpto, str_Prov, str_Distr, str_TipoVia, str_NombreVia, str_Manzana, str_Lote, var4, str_Nac, var5
Dim shell, str_Numero
intStartTime = timer

Num_Iter        = Environment.Value("ActionIteration")
nuevoCiclo      = DataTable("e_CodigoCiclo", dtLocalsheet)
str_Cliente		= DataTable("e_TipoCliente", dtLocalSheet)

Call SelecciondeContacto()
Call CambiodeCiclo()

Sub SelecciondeContacto()

		while(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_2").JavaEdit("Número de documento").Exist)= false
			wait 1
		wend
		
	Call Captura("Panel de Interacción","panel_"&Num_Iter)
	JavaWindow("Ejecutivo de interacción").JavaButton("DIEGO ANTONIO PEREZ VARGAS").Click
	
End Sub
Sub CambiodeCiclo()

		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   mario andres").JavaList("Cliente actual").Exist = False
			wait 1
		Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   mario andres").JavaList("Cliente actual").Select "No se seleccionó cliente"
	Call Down(1)
	
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   mario andres").JavaStaticText("Id del Cliente en Legados:(st)").Exist = False
			wait 1
		Wend
		
	Call Captura("Contacto Creado","Contacto_Creado_"&Num_Iter)
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   mario andres").JavaButton("99").Click
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{ENTER}"
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA").JavaStaticText("Actual:(st)").Exist = False
			wait 1
		Wend
	Call Captura("Cambio de Ciclo","Cambio_Ciclo_"&Num_Iter)
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA").JavaButton("Lookup-notValidated").Click
	
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").Exist = False
			wait 1
		Wend
	wait 5
	
	Filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").GetROProperty("rows")
	Dim Filas, ciclo1, ciclo2, ciclo3, ciclo4
	If Filas = 2 Then
		ciclo1 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").GetCellData(0,1)
		ciclo1 = cstr(ciclo1)
		wait 1
		ciclo2 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").GetCellData(1,1)
		ciclo2 = cstr(ciclo2)
		wait 1
		
		Call Captura("Validacion de Ciclos: Se muestran Los ciclos: "&ciclo1&", "&ciclo2,"Valida_Ciclo_"&Num_Iter)
		
		Select Case nuevoCiclo
			Case ciclo1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").SelectRow "#0"
			Case ciclo2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").SelectRow "#1"
			Case else 
				Call Captura("El ciclo no existe se selcciona por defecto el primer ciclo: "&ciclo1,"CicloDefecto_"&ciclo1&"_"&Num_Iter)
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").SelectRow "#0"
		End Select
		wait 2
		'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").SelectRow "#"&nuevoCiclo
		Call Captura("Se selecciona el ciclo "&nuevoCiclo,"Ciclo_"&nuevoCiclo&"_"&Num_Iter)
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaButton("Seleccionar").Click
	End If
	
	If Filas = 3 Then
		ciclo1 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").GetCellData(0,1)
		ciclo1 = cstr(ciclo1)
		wait 1
		ciclo2 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").GetCellData(1,1)
		ciclo2 = cstr(ciclo2)
		wait 1
		ciclo3 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").GetCellData(2,1)
		ciclo3 = cstr(ciclo3)
		wait 1
		Call Captura("Validacion de Ciclos: Se muestran Los ciclos: "&ciclo1&", "&ciclo2&", "&ciclo3,"Valida_Ciclo_"&Num_Iter)
		
		Select Case nuevoCiclo
			Case ciclo1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").SelectRow "#0"
			Case ciclo2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").SelectRow "#1"
			Case ciclo3
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").SelectRow "#2"
			Case else 
				Call Captura("El ciclo no existe se selcciona por defecto el primer ciclo: "&ciclo1,"CicloDefecto_"&ciclo1&"_"&Num_Iter)
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").SelectRow "#0"
		End Select
		wait 2
		'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").SelectRow "#"&nuevoCiclo
		Call Captura("Se selecciona el ciclo "&nuevoCiclo,"Ciclo_"&nuevoCiclo&"_"&Num_Iter)
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaButton("Seleccionar").Click
	End If
	
	If Filas = 4 Then
	
		ciclo1 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").GetCellData(0,1)
		ciclo1 = cstr(ciclo1)
		wait 1
		ciclo2 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").GetCellData(1,1)
		ciclo2 = cstr(ciclo2)
		wait 1
		ciclo3 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").GetCellData(2,1)
		ciclo3 = cstr(ciclo3)
		wait 1
		ciclo4 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").GetCellData(3,1)
		ciclo4 = cstr(ciclo4)
		
		Call Captura("Validacion de Ciclos: Se muestran Los ciclos: "&ciclo1&", "&ciclo2&", "&ciclo3&", "&ciclo4,"Valida_Ciclo_"&Num_Iter)
		
		Select Case nuevoCiclo
			Case ciclo1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").SelectRow "#0"
			Case ciclo2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").SelectRow "#1"
			Case ciclo3
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").SelectRow "#2"
			Case ciclo4
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").SelectRow "#3"
			Case else 
				Call Captura("El ciclo no existe se selcciona por defecto el primer ciclo: "&ciclo1,"CicloDefecto_"&ciclo1&"_"&Num_Iter)
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").SelectRow "#0"
		End Select
		wait 2
		'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaTable("SearchJTable").SelectRow "#"&nuevoCiclo
		Call Captura("Se selecciona el ciclo "&nuevoCiclo,"Ciclo_"&nuevoCiclo&"_"&Num_Iter)
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA_2").JavaButton("Seleccionar").Click
	End If
	
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA").JavaStaticText("Actual:(st)").Exist = False
			wait 1
		Wend
	Call Captura("Se guarda el cambio de ciclo","CambCiclo23_"&Num_Iter)
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   ANA JULIA").JavaButton("Guardar").Click

		While JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar una opción").Exist = False
			wait 1
		Wend
	Call Captura("Validacion Cambio de Frecuencia","CambFrec_"&Num_Iter)
	
	JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar una opción").JavaButton("Sí").Click
	
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   mario andres").JavaButton("Guardar").Exist = False
			wait 1
		Wend
		
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   mario andres").JavaButton("Guardar").Click
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Contacto:   mario andres").JavaButton("Cerrar").Click

End Sub	
Sub Down(cantidad)
	For Iterator = 1 To cantidad Step 1
		WAIT 1
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{DOWN}"
	Next
End Sub
Sub Tab()
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{TAB}"
End Sub
Sub Captura(Texto,Imagen)
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Imagen&".png", True
	imagenToWord Texto,RutaEvidencias() &Imagen&".png"
End Sub
	

