  Feature: Ventas Prepago

    @Ventas-Prepago-Paso2
    Scenario Outline: Ventas Prepago Paso2
      Given Se da clic en el boton prepago muestra menu de operaciones
      When Se selecciona nueva linea muestra formulario datos del cliente
      And Se da clic en el boton Chip Solo muestra formulario de codigo de barras

      Examples:

        | caso_prueba |
        |           1 |
