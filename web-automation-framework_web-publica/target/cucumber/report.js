$(document).ready(function() {var formatter = new CucumberHTML.DOMFormatter($('.cucumber-report'));formatter.uri("consultaLineas.feature");
formatter.feature({
  "line": 1,
  "name": "Consulta Lineas Moviles",
  "description": "",
  "id": "consulta-lineas-moviles",
  "keyword": "Feature"
});
formatter.scenarioOutline({
  "line": 16,
  "name": "Consulta Lineas B2B",
  "description": "",
  "id": "consulta-lineas-moviles;consulta-lineas-b2b",
  "type": "scenario_outline",
  "keyword": "Scenario Outline",
  "tags": [
    {
      "line": 15,
      "name": "@Consulta_Linea_B2B"
    }
  ]
});
formatter.step({
  "line": 18,
  "name": "Ingreso a la url del Portal \"\u003ccaso_prueba\u003e\"",
  "keyword": "Given "
});
formatter.step({
  "line": 19,
  "name": "Ingreso el Nombre de usuario \"\u003ccaso_prueba\u003e\"",
  "keyword": "When "
});
formatter.step({
  "line": 20,
  "name": "Ingreso la Contraseña \"\u003ccaso_prueba\u003e\"",
  "keyword": "And "
});
formatter.step({
  "line": 21,
  "name": "se da clic en el boton Acceder ingresando correctamente",
  "keyword": "Then "
});
formatter.step({
  "line": 22,
  "name": "Carga la pagina Consulta mis lineas moviles",
  "keyword": "Given "
});
formatter.step({
  "line": 23,
  "name": "Selecciono el tipo de documento \"\u003ccaso_prueba\u003e\"",
  "keyword": "When "
});
formatter.step({
  "line": 24,
  "name": "Se ingresa el numero documento  \"\u003ccaso_prueba\u003e\"",
  "keyword": "And "
});
formatter.step({
  "line": 25,
  "name": "Se da clic en el boton Consultar",
  "keyword": "And "
});
formatter.step({
  "line": 26,
  "name": "Se valida respuesta de Lineas del cliente",
  "keyword": "Then "
});
formatter.examples({
  "line": 28,
  "name": "",
  "description": "",
  "id": "consulta-lineas-moviles;consulta-lineas-b2b;",
  "rows": [
    {
      "cells": [
        "caso_prueba"
      ],
      "line": 29,
      "id": "consulta-lineas-moviles;consulta-lineas-b2b;;1"
    },
    {
      "cells": [
        "5"
      ],
      "line": 30,
      "id": "consulta-lineas-moviles;consulta-lineas-b2b;;2"
    }
  ],
  "keyword": "Examples"
});
formatter.before({
  "duration": 413792400,
  "status": "passed"
});
formatter.before({
  "duration": 7109152600,
  "status": "passed"
});
formatter.scenario({
  "line": 30,
  "name": "Consulta Lineas B2B",
  "description": "",
  "id": "consulta-lineas-moviles;consulta-lineas-b2b;;2",
  "type": "scenario",
  "keyword": "Scenario Outline",
  "tags": [
    {
      "line": 15,
      "name": "@Consulta_Linea_B2B"
    }
  ]
});
formatter.step({
  "line": 18,
  "name": "Ingreso a la url del Portal \"5\"",
  "matchedColumns": [
    0
  ],
  "keyword": "Given "
});
formatter.step({
  "line": 19,
  "name": "Ingreso el Nombre de usuario \"5\"",
  "matchedColumns": [
    0
  ],
  "keyword": "When "
});
formatter.step({
  "line": 20,
  "name": "Ingreso la Contraseña \"5\"",
  "matchedColumns": [
    0
  ],
  "keyword": "And "
});
formatter.step({
  "line": 21,
  "name": "se da clic en el boton Acceder ingresando correctamente",
  "keyword": "Then "
});
formatter.step({
  "line": 22,
  "name": "Carga la pagina Consulta mis lineas moviles",
  "keyword": "Given "
});
formatter.step({
  "line": 23,
  "name": "Selecciono el tipo de documento \"5\"",
  "matchedColumns": [
    0
  ],
  "keyword": "When "
});
formatter.step({
  "line": 24,
  "name": "Se ingresa el numero documento  \"5\"",
  "matchedColumns": [
    0
  ],
  "keyword": "And "
});
formatter.step({
  "line": 25,
  "name": "Se da clic en el boton Consultar",
  "keyword": "And "
});
formatter.step({
  "line": 26,
  "name": "Se valida respuesta de Lineas del cliente",
  "keyword": "Then "
});
formatter.match({
  "arguments": [
    {
      "val": "5",
      "offset": 29
    }
  ],
  "location": "steps_consultaLineasB2B.ingresoALaUrlDelPortal(String)"
});
formatter.result({
  "duration": 2948248100,
  "status": "passed"
});
formatter.match({
  "arguments": [
    {
      "val": "5",
      "offset": 30
    }
  ],
  "location": "steps_consultaLineasB2B.ingresoElNombreDeUsuario(String)"
});
formatter.result({
  "duration": 776648500,
  "status": "passed"
});
formatter.match({
  "arguments": [
    {
      "val": "5",
      "offset": 23
    }
  ],
  "location": "steps_consultaLineasB2B.ingresoLaContraseña(String)"
});
formatter.result({
  "duration": 704246400,
  "status": "passed"
});
formatter.match({
  "location": "steps_consultaLineasB2B.seDaClicEnElBotonAccederIngresandoCorrectamente()"
});
formatter.result({
  "duration": 6731070500,
  "status": "passed"
});
formatter.match({
  "location": "steps_consultaLineasB2B.cargaLaPaginaConsultaMisLineasMoviles()"
});
formatter.result({
  "duration": 635096000,
  "status": "passed"
});
formatter.match({
  "arguments": [
    {
      "val": "5",
      "offset": 33
    }
  ],
  "location": "steps_consultaLineasB2B.seleccionoElTipoDeDocumento(String)"
});
formatter.result({
  "duration": 3772354400,
  "status": "passed"
});
formatter.match({
  "arguments": [
    {
      "val": "5",
      "offset": 33
    }
  ],
  "location": "steps_consultaLineasB2B.seIngresaElNumeroDocumento(String)"
});
formatter.result({
  "duration": 1459168800,
  "status": "passed"
});
formatter.match({
  "location": "steps_consultaLineasB2B.seDaClicEnElBotonConsultar()"
});
formatter.result({
  "duration": 5683151000,
  "status": "passed"
});
formatter.match({
  "location": "steps_consultaLineasB2B.seValidaRespuestaDeLineasDelCliente()"
});
formatter.result({
  "duration": 589987600,
  "status": "passed"
});
formatter.after({
  "duration": 1177147000,
  "status": "passed"
});
});