Feature: Flujos de USSD Versión 2

  @USSD_Testing_1 @RegTesting
  Scenario Outline: Recibir SMS para descargar APP

    Given Se ingresa a USSD mediante *183# Testing
    When se ingresa a la opcion Consultar mi prepago
    And se ingresa la opcion Recibir SMS para descargar APP
    Then se verifica el envio de SMS

    Examples:
      | caso_prueba |
      | 1           |

  @USSD_Testing_2 @RegTesting
  Scenario Outline: Consulta Saldo Prepago

    Given Se ingresa a USSD mediante *183# Testing
    When se ingresa a la opcion Consultar mi prepago
    And se ingresa la opcion Continua aqui
    And se elige Consultar Saldo
    Then se obtiene el saldo actual

    Examples:
      | caso_prueba |
      | 1           |

  @USSD_Testing_3 @RegTesting
  Scenario Outline: Consulta Contrato Prepago

    Given Se ingresa a USSD mediante *183# Testing
    When se ingresa a la opcion Consultar mi prepago
    And se ingresa la opcion Continua aqui
    And se ingresa a Mas sobre mi Prepago
    And se elige Consultar Contrato Prepago
    Then se obtiene el contrato actual

    Examples:
      | caso_prueba |
      | 1           |

  @USSD_Testing_4 @RegTesting
  Scenario Outline: Consulta Tarifa de LLAMADAS

    Given Se ingresa a USSD mediante *183# Testing
    When se ingresa a la opcion Consultar mi prepago
    And se ingresa la opcion Continua aqui
    And se ingresa a Mas sobre mi Prepago
    And se elige Tarifas
    And se elige la opción LLAMADAS
    Then se obtiene la tarifa de LLAMADAS

    Examples:
      | caso_prueba |
      | 1           |

  @USSD_Testing_5 @RegTesting
  Scenario Outline: Consulta Tarifa de DATOS

    Given Se ingresa a USSD mediante *183# Testing
    When se ingresa a la opcion Consultar mi prepago
    And se ingresa la opcion Continua aqui
    And se ingresa a Mas sobre mi Prepago
    And se elige Tarifas
    And se elige la opción DATOS
    Then se obtiene la tarifa de DATOS

    Examples:
      | caso_prueba |
      | 1           |

  @USSD_Testing_6 @RegTesting
  Scenario Outline: Consulta Perdida o Robo

    Given Se ingresa a USSD mediante *183# Testing
    When se ingresa a la opcion Consultar mi prepago
    And se ingresa la opcion Continua aqui
    And se ingresa a Mas sobre mi Prepago
    And se elige Perdida o robo de un equipo
    Then se obtiene la información  solicitada

    Examples:
      | caso_prueba |
      | 1           |

  @USSD_Testing_7 @RegTesting
  Scenario Outline: Consulta Llamadas Internacionales de Movistar 1911

    Given Se ingresa a USSD mediante *183# Testing
    When se ingresa a la opcion Consultar mi prepago
    And se ingresa la opcion Continua aqui
    And se ingresa a Mas sobre mi Prepago
    And se elige Llamadas internacionales
    And se elige la opción Movistar 1911
    Then se obtiene la información de Movistar 1911

    Examples:
      | caso_prueba |
      | 1           |

  @USSD_Testing_8 @RegTesting
  Scenario Outline: Consulta Llamadas Internacionales de Otros Operadores

    Given Se ingresa a USSD mediante *183# Testing
    When se ingresa a la opcion Consultar mi prepago
    And se ingresa la opcion Continua aqui
    And se ingresa a Mas sobre mi Prepago
    And se elige Llamadas internacionales
    And se elige la opción Otros Operadores
    Then se obtiene la información de Otros Operadores

    Examples:
      | caso_prueba |
      | 1           |

  @USSD_Testing_9 @RegTesting
  Scenario Outline: Compra Paquetes LLAMADAS

    Given Se ingresa a USSD mediante *183# Testing
    When se ingresa a la opcion Comprar Paquetes
    And se ingresa la opcion LLamadas+Apps+Datos+SMS
    And se elige el monto a comprar (3 soles)
    Then se verifica el envio de SMS al comprar 3 soles

    Examples:
      | caso_prueba |
      | 1           |

  @USSD_Testing_10 @RegTesting
  Scenario Outline: Beneficios Prepago : Bono Por Antiguedad

    Given Se ingresa a USSD mediante *183# Testing
    When se ingresa a la opcion Beneficios Prepago
    And se ingresa la opcion Bono por antiguedad
    Then se verifica el mensaje validador

    Examples:
      | caso_prueba |
      | 1           |

  @USSD_Testing_11 @RegTesting
  Scenario Outline: Comparte tu Saldo

    Given Se ingresa a USSD mediante *183# Testing
    When se ingresa a la opcion Servicios
    And se ingresa la opcion Comparte tu Saldo
    And se ingresa el monto a compartir (1 sol)
    And se ingresa el telefono de destino
    And se confirma el envío
    Then se verifica el mensaje de saldo insuficiente


    Examples:
      | caso_prueba |
      | 1           |

  @USSD_Testing_12 @RegTesting
  Scenario Outline: Roaming : Cobertura Pasaporte Movistar

    Given Se ingresa a USSD mediante *183# Testing
    When se ingresa a la opcion Roaming
    And se ingresa la opcion consultar Cobertura
    And se ingresa la opcion ver más
    And se ingresa la opcion Continuar
    Then se verifica el mensaje de Cobertura Pasaporte Movistar


    Examples:
      | caso_prueba |
      | 1           |

