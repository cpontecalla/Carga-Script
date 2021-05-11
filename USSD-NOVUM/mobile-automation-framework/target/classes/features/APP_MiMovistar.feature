Feature: APP_MiMovistar NOVUM

  @APP_MiMovistar_TEST1
    Scenario Outline: Cambiar Nombre del Titular

      Given Se ingresa a la app Mi Movistar y se da click al boton empieza ahora
      And se selecciona el ingreso como titular
      When se ingresa el DNI "<caso_prueba>"
      And se ingresa la contrasenia de 6 numeros "<caso_prueba>"
      And se da click en el boton ingresar
      And se da click en ajustes y se selecciona informacion personal
      And se ingresa a datos personales y al nombre del titular
      Then se cambiar el nombre del titular por "<caso_prueba>"
      And se cambia el apellido del titular por "<caso_prueba>"
      And se da click en el boton guardar
      Then se realiza el logout de la app

      Examples:
        | caso_prueba |
        | 1           |

  @APP_MiMovistar_TEST2
  Scenario Outline: Comprar Paquete de Datos (Instagram)

    Given Se ingresa a la app Mi Movistar y se da click al boton empieza ahora
    And se selecciona el ingreso como titular
    When se ingresa el DNI "<caso_prueba>"
    And se ingresa la contrasenia de 6 numeros "<caso_prueba>"
    And se da click en el boton ingresar
    Then se da click en el boton compra paquetes
    And se selecciona paquete de datos
    And se agrega el paquete Instagram Ilim X
    And se visualiza las caracteristicas y se da click en el boton pagar
    Then se visualiza la confirmacion de compra y se da click a volver a mi linea
    Then se realiza el logout de la app

    Examples:
      | caso_prueba |
      | 1           |

  @APP_MiMovistar_TEST3
  Scenario Outline: Comprar Paquete de Minutos (45 Minutos)

    Given Se ingresa a la app Mi Movistar y se da click al boton empieza ahora
    And se selecciona el ingreso como titular
    When se ingresa el DNI "<caso_prueba>"
    And se ingresa la contrasenia de 6 numeros "<caso_prueba>"
    And se da click en el boton ingresar
    Then se da click en el boton compra paquetes
    And se selecciona paquete de minutos
    And se agrega el paquete de 45 minutos
    And se visualiza las caracteristicas y se da click en el boton pagar
    Then se visualiza la confirmacion de compra y se da click a volver a mi linea
    Then se realiza el logout de la app

    Examples:
      | caso_prueba |
      | 1           |


  @APP_MiMovistar_TEST4
  Scenario Outline: Validar Gestión de Sesiones (Equipo)

    Given Se ingresa a la app Mi Movistar y se da click al boton empieza ahora
    And se selecciona el ingreso como titular
    When se ingresa el DNI "<caso_prueba>"
    And se ingresa la contrasenia de 6 numeros "<caso_prueba>"
    And se da click en el boton ingresar
    And se da click en ajustes y se selecciona seguridad y privacidad
    And se da click en gestion de sesiones y se valida el dispositivo
    Then se realiza el logout de la app

    Examples:
      | caso_prueba |
      | 1           |