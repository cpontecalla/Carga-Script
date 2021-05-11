  Feature: Ventas Altas Postpago

    @Ventas-Altas-Postpago
    Scenario Outline: Ventas Altas Postpago
      Given Se da clic en el boton postpago muestra menu de operaciones
      When Se selecciona nueva linea se ingresa datos del cliente
      And Se da clic en boton Chip Solo muestra formulario de codigo de barras

      Examples:

        | caso_prueba |
        |           1 |
