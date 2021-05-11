Feature: Consulta Lineas Moviles

  @Login_Ambiente
  Scenario Outline: Login Portal Web Exitoso

    Given Ingreso a la url del Portal "<caso_prueba>"
    When Ingreso el Nombre de usuario "<caso_prueba>"
    And Ingreso la Contraseña "<caso_prueba>"
    Then se da clic en el boton Acceder ingresando correctamente

    Examples:
      | caso_prueba |
      |           1 |

  @Consulta_Linea_B2B
  Scenario Outline: Consulta Lineas B2B

    Given Ingreso a la url del Portal "<caso_prueba>"
    When Ingreso el Nombre de usuario "<caso_prueba>"
    And Ingreso la Contraseña "<caso_prueba>"
    Then se da clic en el boton Acceder ingresando correctamente
    Given Carga la pagina Consulta mis lineas moviles
    When Selecciono el tipo de documento "<caso_prueba>"
    And Se ingresa el numero documento  "<caso_prueba>"
    And Se da clic en el boton Consultar
    Then Se valida respuesta de Lineas del cliente

    Examples:
      | caso_prueba |
      |           5 |

  @Consulta_Linea_ValidarTodoDoc
  Scenario Outline: Consultas B2B DNI Ce Pasaporte

    Given Ingreso a la url del Portal "<caso_prueba>"
    When Ingreso el Nombre de usuario "<caso_prueba>"
    And Ingreso la Contraseña "<caso_prueba>"
    Then se da clic en el boton Acceder ingresando correctamente
    Given Carga la pagina Consulta mis lineas moviles
    When Selecciono el tipo de documento "<caso_prueba>"
    And Se ingresa el numero documento  "<caso_prueba>"
    And Se da clic en el boton Consultar
    Then Se valida respuesta de Lineas del cliente

    Examples:
      | caso_prueba |
      |           1 |
      |           2 |
      |           3 |
      |           4 |
      |           5 |
      |           6 |


  @Consulta_Linea_B2C
  Scenario Outline: Consulta de Lineas B2C

    Given Ingreso a la url del Portal "<caso_prueba>"
    When Ingreso el Nombre de usuario "<caso_prueba>"
    And Ingreso la Contraseña "<caso_prueba>"
    Then se da clic en el boton Acceder ingresando correctamente
    Given Carga la pagina Consulta mis lineas moviles
    When Selecciono el tipo de documento "<caso_prueba>"
    And Se ingresa el numero documento  "<caso_prueba>"
    And Se da clic en el boton Consultar
    Then Se valida respuesta de Lineas del cliente

    Examples:
      | caso_prueba |
      |           1 |
      |           2 |