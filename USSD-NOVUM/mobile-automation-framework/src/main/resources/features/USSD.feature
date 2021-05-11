Feature: Flujos de USSD

  @USSD_Test1
  Scenario Outline: Consulta de Tarifas

    Given Se ingresa a USSD mediante "*515#"
    When se ingresa a la opcion Consultas
    And se ingresa a la opcion Mas
    Then se ingresa a la opcion Tarifas

    Examples:
      | caso_prueba |
      | 1           |

  @USSD_Test2
  Scenario Outline: Consulta de Perdida o Robo de tu Equipo

    Given Se ingresa a USSD mediante "*515#"
    When se ingresa a la opcion Consultas
    And se ingresa a la opcion Mas
    Then se ingresa a la opcion Perdida/robo de tu equipo

    Examples:
      | caso_prueba |
      | 1           |

  @USSD_Test3
  Scenario Outline: Consulta de Llamadas Internacionales

    Given Se ingresa a USSD mediante "*515#"
    When se ingresa a la opcion Consultas
    And se ingresa a la opcion Mas
    Then se ingresa a la opcion Llamadas Internacionales
    And se verifica para otros operadores

    Examples:
      | caso_prueba |
      | 1           |

  @USSD_Test4
  Scenario Outline: Duplicar MB + FB GRATIS

    Given Se ingresa a USSD mediante "*515#"
    When se ingresa a la opcion Duplicar MB / FB Gratis
    Then se ingresa la opcion de duplicar 66MB

    Examples:
      | caso_prueba |
      | 1           |

