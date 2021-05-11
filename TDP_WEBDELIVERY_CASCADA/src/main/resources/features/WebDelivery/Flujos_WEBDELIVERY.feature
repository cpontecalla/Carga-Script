Feature: FlujoWEB_DELIVERY


  @CREAR_PEDIDO
  Scenario Outline: Crear Pedido_ALTA_MASIVA_CAEQ_MASIVO
    Given INGRESAMOS A LA URL DE WEB DELIVERY "<caso_prueba>"
    When INGRESAMOS USUARIO A WEB DELIVERY"<caso_prueba>"
    And INGRESAMOS PASSWORD WEB DELIVER"<caso_prueba>"
    Then CLICK BOTON LOGIN INGRESANDO CORRECTAMENTE A LA PAGINA
    Given click en crear pedido
    When Ingresar y buscar el numero de RUC "<caso_prueba>"
    And Ingresar el tipo de pedido y almacen "<caso_prueba>"
    And Infromación del solicitante"<caso_prueba>"
    And Dirección de entrega
    And Información del receptor "<caso_prueba>"
    Then click en botón continuar
    Given Click botón fila nueva
    When Linea de detalle de solicitud (Alta)"<caso_prueba>"
    And Click botón consultar disponibilidad
    And Click botón realizar reserva
    And Click botón generar detalles del pedido
    And Click botón continuar
    And Click botón continuar siguiente
    And Click botón enviar "<caso_prueba>"
    Then Guardar el código de pedido "<caso_prueba>"

    Examples:
      | caso_prueba |
      |           1 |

  @CARGA_MATERIALES
  Scenario Outline: Carga de Materiales

    Given INGRESAR A LA URL DE WEB DELIVERY "<caso_prueba>"
    When INGRESAR USUARIO A WEB DELIVERY"<caso_prueba>"
    And INGRESAR CONTRASENA WEB DELIVER"<caso_prueba>"
    Then CLICK EN EL BOTON LOGIN INGRESANDO CORRECTAMENTE A LA PAGINA
    Given CLICK EN EL BOTON IR A EN WEB DELIVERY "<caso_prueba>"
    When SELECCIONAR AJUSTE DE INVENTARIO
    And CLICK EN EL BOTON NUEVO REGISTRO
    And INGRESAR TIPO ABASTECIMIENTO "<caso_prueba>"
    And INGRESAR COMENTARIO "<caso_prueba>"
    And INGRESAR GUIA DE REMISION "<caso_prueba>"
    And CARGAR ARCHIVO CSV
    And EJECUTAR AJUSTE Y ACEPTAR MENSAJE
    And VALIDAR LA CARGA DE ARCHIVO CSV

    Examples:
      | caso_prueba |
      |           1 |

  @RESIDENCIAL_1
  Scenario Outline: Flujo Residencial - Parte 1

    Given INGRESAR A LA URL DE WEB DELIVERY "<caso_prueba>"
    When INGRESAR USUARIO A WEB DELIVERY"<caso_prueba>"
    And INGRESAR CONTRASENA WEB DELIVER"<caso_prueba>"
    Then CLICK EN EL BOTON LOGIN INGRESANDO CORRECTAMENTE A LA PAGINA
    When SELECCIONAMOS AUDITORIA DE PEDIDO "<caso_prueba>"
    And BUSCAMOS ID DE ORDEN "<caso_prueba>"
    And SELECCIONAMOS PEDIDO
    Given AGENDAMOS EL PEDIDO "<caso_prueba>"
    When SELECCIONAMOS FECHA DE PEDIDO"<caso_prueba>"
    Then VALIDAMOS CAMBIO DE ESTADO DEL PEDIDO (AGENDADO) "<caso_prueba>"
    When SELECCIONAMOS ASIGNACION DE SERIES "<caso_prueba>"
    And BUSCAMOS EL ORDER ID  "<caso_prueba>"
    And INGRESAMOS MATERIALES "<caso_prueba>"
    Given VALIDAMOS SERIES
    And BUSCAMOS EL ORDER ID  "<caso_prueba>"
    Then VERIFICAMOS ESTADO DE VALIDACION DE SERIES "<caso_prueba>"
    When SELECCIONAMOS IMPRESION DE DOCUMENTOS"<caso_prueba>"
    And BUSCAMOS EL ORDER_ID  "<caso_prueba>"
    And SELECCIONAMOS EJECUTAR INFORMES "<caso_prueba>"
    Given IMPRESION DE GUIA DE REMISION
    And IMPRESION DE ETIQUETA  "<caso_prueba>"
    And BUSCAMOS EL ORDER_ID  "<caso_prueba>"
    Then VERIFICAMOS QUE EL ESTADO DE LA ORDEN SEA EL CORRECTO (REALIZADO)
    When SELECCIONAMOS DESPACHO DE HUB "<caso_prueba>"
    And BUSCAMOS EL ORDER_ID  "<caso_prueba>"
    And ASIGNAMOS MASTER BOX "<caso_prueba>"
    Then VERIFICAMOS LA ASIGNACION CORRECTA DEL CODIGO
    And DESPACHO DE PEDIDO
    When RECEPCIONAR PEDIDOS
    And BUSCAMOS EL ORDER_ID  "<caso_prueba>"
    And RECEPCIONAR PEDIDO RESI
    And BUSCAMOS EL ORDER_ID  "<caso_prueba>"

    Examples:
      | caso_prueba |
      |           1 |

  @RESIDENCIAL_2
  Scenario Outline: Flujo Residencial - Parte 2

    Given INGRESAR A LA URL DE WEB DELIVERY "<caso_prueba>"
    When INGRESAR USUARIO A WEB DELIVERY"<caso_prueba>"
    And INGRESAR CONTRASENA WEB DELIVER"<caso_prueba>"
    Then CLICK EN EL BOTON LOGIN INGRESANDO CORRECTAMENTE A LA PAGINA
    When RECEPCIONAR PEDIDOS
    And BUSCAMOS EL ORDER_ID  "<caso_prueba>"
    And EJECUTAR CARGA DE LOGICA DE RUTEO
    And BUSCAMOS EL ORDER_ID  "<caso_prueba>"
    And GUARDAR NUMERO DE ENVIO
    When DESPACHO DE PEDIDO ENVIADO
    And BUSCAMOS NUMERO DE ENVIO
    And BUSCAMOS MOTORIZADO
    And DESPACHAMOS PEDIDO
    Then VALIDAMOS EL ESTADO DE PEDIDO DESPACHADO

    Examples:
      | caso_prueba |
      |           1 |

  @CORPORATIVO_1
  Scenario Outline: Flujo corporativo (ID DE ORDEN)

    Given INGRESAMOS A LA URL WEB DELIVERY "<caso_prueba>"
    When INGRESAMOS USUARIO WEB DELIVERY"<caso_prueba>"
    And INGRESAMOS CONTRASEÑA WEB DELIVERY"<caso_prueba>"
    Then CLICK BOTON LOGIN WEB DELIVERY Y SE INGRESA CORRECTAMENTE
    Given SELECCIONAR ASIGNACIÓN DE SERIES(CORPORATIVO)
    When BUSCAR EL ID DE RESERVA(CORPORATIVO)"<caso_prueba>"
    And INGRESAR MATERIALES(CORPORATIVO) "<caso_prueba>"
    And VALIDAR SERIES(CORPORATIVO)
    When BUSCAR EL ID DE RESERVA(CORPORATIVO)"<caso_prueba>"
    Then VERFICIAR EL ESTADO DE LAS SERIES(CORPORATIVO)"<caso_prueba>"
    When SELECCIONAR IMPRESION DE DOCUMENTOS(CORPORATIVO)
    And BUSCAR ID DE RESERVA2(CORPORATIVO)"<caso_prueba>"
    And SELECCIONAR CONTRATO DIGITAL(CORPORATIVO)
    And SELECCIONAR PREPARACION DE LA FACTURA(CORPORATIVO)
    And IMPRIMIR FACTURA(CORPORATIVO)
    And EJECUTAR INFORMES(CORPORATIVO)
    And IMPRIMIR GUIA DE REMISION(CORPORATIVO)
    And IMPRIMIR ETIQUETA DE LINEA(CORPORATIVO)
    And IMPRIMIR ETIQUETA(CORPORATIVO)
    And BUSCAR ID DE RESERVA2(CORPORATIVO)"<caso_prueba>"
    Then VERIFICAR CAMBIO DE ESTADO
    When SELECCIONAR DESPACHO DE HUB(CORPORATIVO)
    And BUSCAR ID DE RESERVA2(CORPORATIVO)"<caso_prueba>"
    And EJECUTAR INFORMES(CORPORATIVO)
    And GENERAR MASTER BOX(CORPORATIVO)
    And BUSCAR ID DE RESERVA2(CORPORATIVO)"<caso_prueba>"
    Then DESPACHAR PEDIDO(CORPORATIVO)
    When SELECCIONAR RECEPCION DE PEDIDOS(CORPORATIVO)
    And BUSCAR ID DE RESERVA2(CORPORATIVO)"<caso_prueba>"
    And RECEPCIONAR PEDIDOS(CORPORATIVO)
    And EJECUTAR INFORMES(CORPORATIVO)
    And EJECUTAR PROCESO DE LOGICA DE RUTEO(CORPORATIVO)
    And EJECUTAR REPORTE DE LOGICA DE RUTEO(CORPORATIVO)
    And BUSCAR ID DE RESERVA2(CORPORATIVO)"<caso_prueba>"

    Examples:
      | caso_prueba |
      |           1 |


  @CORPORATIVO_2
  Scenario Outline: Flujo corporativo (ID DE ORDEN)

    Given INGRESAMOS A LA URL WEB DELIVERY "<caso_prueba>"
    When INGRESAMOS USUARIO WEB DELIVERY"<caso_prueba>"
    And INGRESAMOS CONTRASEÑA WEB DELIVERY"<caso_prueba>"
    Then CLICK BOTON LOGIN WEB DELIVERY Y SE INGRESA CORRECTAMENTE
    When Seleccionamos recepción de pedidos(corporativo)
    And BUSCAR ID DE RESERVA2(CORPORATIVO)"<caso_prueba>"
    And EJECUTAR CARGA LOGICA DE RUTEO (CORPORATIVO)
    And BUSCAR ID DE RESERVA2(CORPORATIVO)"<caso_prueba>"
    Then GUARDAMOS NUMERO DE ENVIO (CORPORATIVO)
    When DESPACHO A MOTORIZADO (CORPORATIVO)
    And BUSCAMOS NUMERO DE ENVIO (CORPORATIVO)
    And BUSCAMOS MOTORIZADO (CORPORATIVO)
    And DESPACHAMOS PEDIDO (CORPORATIVO)
    When MAESTRO DE PEDIDOS (CORPORATIVO)
    And BUSCAMOS EL ED RESERVA MAESTRO DE PEDIDOS(CORPORATIVO)"<caso_prueba>"
    Then VERIFICAMOS EL ESTADO DE PEDIDO (CORPORATIVO)
    And ENTREGAR PEDIDO (CORPORATIVO)
    Then VALIDAR ESTADO DE PEDIDO (CORPORATIVO)


    Examples:
      | caso_prueba |
      |           2 |
  @MASIVO_1
  Scenario Outline: FLUJO MASIVOS

    Given INGRESAMOS A LA URL WEB DELIVERY "<caso_prueba>"
    When INGRESAMOS USUARIO WEB DELIVERY"<caso_prueba>"
    And INGRESAMOS CONTRASEÑA WEB DELIVERY"<caso_prueba>"
    Then CLICK BOTON LOGIN WEB DELIVERY Y SE INGRESA CORRECTAMENTE
    When SELECCIONAMOS ASIGNACION DE SERIES(CORPORATIVO)
    And BUSCAMOS PEDIDO PARA ASIGNAR SERIES(CORPORATIVO)"<caso_prueba>"
    And INGRESAMOS MATERIALES(CORPORATIVO) "<caso_prueba>"
    And VALIDAMOS SERIES INGRESADAS (CORPORATIVO)
    And BUSCAMOS PEDIDO(CORPORATIVO)"<caso_prueba>"
    Then VERIFICAMOS ESTADO DE SERIES (CORPORATIVO)"<caso_prueba>"
    When SELECCIONAR IMPRRESION DE DOCUMENTOS(CORPORATIVO)
    And BUSCAMOS PEDIDO(CORPORATIVO)"<caso_prueba>"
    And Contrato digital (corporativo)
    And Preparación de la factura (corporativo)
    And Impresión de la factura (corporativo)
    And Ejecutar informes (corporativo)
    And Imprimir guia de remisión (corporativo)
    And Imprimir etiqueta de linea (corporativo
    And Imprimir etiqueta (corporativo)
    And BUSCAMOS PEDIDO(CORPORATIVO)"<caso_prueba>"
    Then Observamos el cambio de estado
    When Seleccionamos despacho de hub(corporativo)
    And BUSCAMOS PEDIDO(CORPORATIVO)"<caso_prueba>"
    And Ejecución de informes(corporativo)
    And Generación de Master BOX (corporativo)
    And BUSCAMOS PEDIDO(CORPORATIVO)"<caso_prueba>"
    Then Despachar pedido (corporativo)
    When Seleccionar recepción de pedidos (corporativo)
    And BUSCAMOS PEDIDO(CORPORATIVO)"<caso_prueba>"
    And Recepcionar de pedidos (corporativo)
    And Ejecución de informes(corporativo)
    And Proceso de lógica de ruteo (corporativo)
    And Reporte de logica de ruteo (corporativo)
    And BUSCAMOS PEDIDO(CORPORATIVO)"<caso_prueba>"

    Examples:
      | caso_prueba |
      |           1 |
      |           3 |

  @MASIVO_2
  Scenario Outline: FLUJO MASIVOS

    Given INGRESAMOS A LA URL WEB DELIVERY "<caso_prueba>"
    When INGRESAMOS USUARIO WEB DELIVERY"<caso_prueba>"
    And INGRESAMOS CONTRASEÑA WEB DELIVERY"<caso_prueba>"
    Then CLICK BOTON LOGIN WEB DELIVERY Y SE INGRESA CORRECTAMENTE
    When Seleccionamos recepción de pedidos C (corporativo)
    And BUSCAMOS PEDIDO(CORPORATIVO)"<caso_prueba>"
    And Ejecutar Carga de Ruteo (corporativo)
    And BUSCAMOS PEDIDO(CORPORATIVO)"<caso_prueba>"
    Then Guardamos número de envio (corporativo)
    When Despacho a motorizado (corporativo)
    And Buscamos numero de envio (corporativo)
    And Buscamos motorizado (corporativo)
    And Despachamos Pedido (corporativo)
    When Maestro de Pedidos (corporativo)
    And Buscamos el numero de pedido"<caso_prueba>"
    Then Verificamos el estado del pedido (corporativo)
    And Entregar pedido (corporativo)
    Then Validar estado del pedido (corporativo)

    Examples:
      | caso_prueba |
      |           1 |
      |           3 |