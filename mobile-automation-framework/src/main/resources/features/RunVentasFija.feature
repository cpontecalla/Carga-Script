Feature: Ventas Fijas

  @VentasAltasFijas
  Scenario Outline: Altas-Fijas

    Given Abrir la aplicación e ingresar codigo de vendedor "<caso_prueba>"
    When Muestra menu principal y se da clic en altas nuevas
    And Se ingresa datos del contratante y clic en evaluar
    Then Verificar cliente sin restriccion de deuda e iniciar venta
    Given Se selecciona direccion de instalacion
    When Se completa datos de la direccion de instalacion
    And Se selecciona campana disponible para cliente
    Then Se verifica detalle del producto
    And Se selecciona SVA
    When Se aceptan las condiciones
    And Se muestra el resumen de venta
    Then Se realiza la lectura del contrato
    And Se realiza validacion de identidad
    Then Se selecciona foto de dni
    And Se valida venta exitosa

    Examples:
      | caso_prueba |
      |           1 |

  @VentasMigraciones
  Scenario Outline: Migraciones-Fijas
    Given Abrir la aplicación e ingresar codigo de vendedor "<caso_prueba>"
    When Muestra menu principal y se da clic en migraciones
    And Se busca productos del cliente
    Then Se selecciona opcion linea o television
    And Se selecciona el producto
    Then Muestra lista de campanias

    Examples:
      | caso_prueba |
      |           1 |