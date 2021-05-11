Feature: App Ventas Testing

  @App_Ventas_Test1
  Scenario Outline: Flujo de Login APP Ventas

    Given Se ingresa a la apk y se ingresa el DNI del vendedor "<caso_prueba>"
    When se clic en el boton ingresar

    Examples:
      | caso_prueba |
      | 1           |

  @App_Ventas_Test2
  Scenario Outline: Flujo de Alta PREPAGO

    Given Se ingresa a la apk y se ingresa el DNI del vendedor "<caso_prueba>"
    When se clic en el boton ingresar
    When se da click en el boton Guardar de App Ventas
    And se da click en el boton Acepto del Aviso
    And se elige la venta Prepago
    And se elige la operacion Nueva Linea
    And se ingresa el numero de documento "<caso_prueba>"
    Then se da click en el boton CHIP SOLO
    And se ingresa la serie del chip "<caso_prueba>"
    And se da click en el boton continuar


    Examples:
      | caso_prueba |
      | 1           |
