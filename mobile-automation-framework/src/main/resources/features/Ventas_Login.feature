Feature: Ventas Login

  @Ventas-Login
  Scenario Outline: Ventas-Login
    Given Abrir la aplicación e ingresar número del Dni "<caso_prueba>"
    When Se da clic al boton ingresar se muestra formulario de configuracion
    And Se da clic en el boton guardar se muestra pop up
    Then Se da clic en el boton Acepto y muestra menu de productos
    Given Se da clic en el boton prepago muestra menu de operaciones
    When Se selecciona nueva linea muestra formulario datos del cliente
    And Se da clic en el boton Chip Solo muestra formulario de codigo de barras

    Examples:
      | caso_prueba |
      |           1 |
