Feature: Ventas Prepago

  @VentasAltasPrepago
  Scenario Outline: Ventas-AltaPrepago

    Given Abrir la aplicación e ingresar número de Dni "<caso_prueba>"
    When Se da clic al boton ingresar se muestra formulario de configuracion
    And Se da clic en el boton guardar se muestra pop up
    Then Se da clic en el boton Acepto y muestra menu de productos
    Given Se da clic en el boton prepago muestra menu de operaciones
    When Se selecciona nueva linea muestra formulario datos del cliente
    And Se da clic en el boton Chip Solo muestra formulario de codigo de barras
    Then Se genera nro telefono y se selecciona preplan
    And Se selecciona centro poblado y muestra huellero

    Examples:
      | caso_prueba |
      |           1 |

  @VentasAltasPostpago
  Scenario Outline: Ventas-AltaPostpago
    Given Abrir la aplicación e ingresar número de Dni "<caso_prueba>"
    When Se da clic al boton ingresar se muestra formulario de configuracion
    And Se da clic en el boton guardar se muestra pop up
    Then Se da clic en el boton Acepto muestra menu de productos
    Given Se da clic en el boton postpago muestra menu de operaciones
    When Se selecciona nueva linea se ingresa datos del cliente
    And Se da clic en boton Chip Solo muestra formulario de codigo de barras
    And Selecciona centro poblado y muestra huellero

    Examples:
      | caso_prueba |
      |           1 |

  @VentasCasiPrepago
  Scenario Outline: Ventas-CasiPrepago
    Given Abrir la aplicación e ingresar número de Dni "<caso_prueba>"
    When Se da clic al boton ingresar se muestra formulario de configuracion
    And Se da clic en el boton guardar se muestra pop up
    Then Muestra menu de operaciones y se selecciona Prepago
    And Muestra opciones de prepago y Se selecciona Renovacion
    When Se ingresa datos del cliente se da clic en chip solo
    And Se ingresa serie y clic en continuar

    Examples:
      | caso_prueba |
      |           1 |

  @VentasCasiPostpago
  Scenario Outline: Ventas-CasiPrepago
    Given Abrir la aplicación e ingresar número de Dni "<caso_prueba>"
    When Se da clic al boton ingresar se muestra formulario de configuracion
    And Se da clic en el boton guardar se muestra pop up
    Then Muestra menu de operaciones y se selecciona Postpago
    And Muestra opciones de postpago y Se selecciona Renovacion
    When Se ingresa datos del cliente postpago se da clic en chip solo
    And Se ingresa serie plan y clic en continuar
    Then Se ingresa datos de facturacion
    And Muestra contrato y huellero

    Examples:
      | caso_prueba |
      |           1 |

  @VentasPortabilidadPrepago
  Scenario Outline: Ventas-PortabilidadPrepago
    Given Abrir la aplicación e ingresar número de Dni "<caso_prueba>"
    When Se da clic al boton ingresar se muestra formulario de configuracion
    And Se da clic en el boton guardar se muestra pop up
    Then Muestra menu de operaciones se selecciona Prepago
    And Muestra opciones de prepago y Se selecciona Portabilidad
    When Se ingresa datos del cliente y se da clic en validar linea
    And Se ingresa serie clic en continuar
    Then Se ingresa datos de centro poblado y huellero

    Examples:
      | caso_prueba |
      |           1 |

  @VentasPortabilidadPostpago
  Scenario Outline: Ventas-PortabilidadPostpago
    Given Abrir la aplicación e ingresar número de Dni "<caso_prueba>"
    When Se da clic al boton ingresar se muestra formulario de configuracion
    And Se da clic en el boton guardar se muestra pop up
    Then Muestra menu de operaciones se selecciona Postpago
    And Muestra opciones de postpago y Se selecciona Portabilidad
    When Se ingresa datos del cliente y se da clic en validar linea postpago
    And Se ingresa serie y plan clic en continuar
    Then Se ingresa datos de centro poblado facturacion y huellero

    Examples:
      | caso_prueba |
      |           1 |