package com.tsoft.bot.frontend.steps.WebDelivery;

import com.tsoft.bot.frontend.pages.pages.P_Corporativo;
import cucumber.api.PendingException;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import org.openqa.selenium.WebDriver;

public class Corporativo {
    public WebDriver driver;
    P_Corporativo Corporativo = new P_Corporativo(driver);
    @Given("^INGRESAMOS A LA URL WEB DELIVERY \"([^\"]*)\"$")
    public void ingresamosALAURLWEBDELIVERY(String dato) throws Throwable {
        Corporativo.ingresamosALAURLWEBDELIVERY(dato);
    }
    @When("^INGRESAMOS USUARIO WEB DELIVERY\"([^\"]*)\"$")
    public void ingresamosUSUARIOWEBDELIVERY(String dato) throws Throwable {
        Corporativo.ingresamosUSUARIOWEBDELIVERY(dato);
    }

    @And("^INGRESAMOS CONTRASEÑA WEB DELIVERY\"([^\"]*)\"$")
    public void ingresamosCONTRASEÑAWEBDELIVERY(String dato) throws Throwable {
        Corporativo.ingresamosCONTRASEÑAWEBDELIVERY(dato);
    }
    @Then("^CLICK BOTON LOGIN WEB DELIVERY Y SE INGRESA CORRECTAMENTE$")
    public void clickBOTONLOGINWEBDELIVERYYSEINGRESACORRECTAMENTE() throws Throwable{
        Corporativo.clickBOTONLOGINWEBDELIVERYYSEINGRESACORRECTAMENTE();
    }

    @Given("^SELECCIONAR ASIGNACIÓN DE SERIES\\(CORPORATIVO\\)$")
    public void seleccionarASIGNACIÓNDESERIESCORPORATIVO()throws Throwable {
        Corporativo.seleccionamosAsignaciónDeSeriesCorporativo();
    }

    @When("^BUSCAR EL ID DE RESERVA\\(CORPORATIVO\\)\"([^\"]*)\"$")
    public void buscarELIDDERESERVACORPORATIVO(String dato) throws Throwable {
       Corporativo.buscamosElIdDeReservaCorporativo(dato);
    }

    @And("^INGRESAR MATERIALES\\(CORPORATIVO\\) \"([^\"]*)\"$")
    public void ingresarMATERIALESCORPORATIVO(String dato) throws Throwable {
        Corporativo.ingresamosMaterialesIMEIYSIMCARD(dato);
    }

    @And("^VALIDAR SERIES\\(CORPORATIVO\\)$")
    public void validarSERIESCORPORATIVO() throws Throwable {
        Corporativo.validamosSeriesCorporativo();
    }

    @Then("^VERFICIAR EL ESTADO DE LAS SERIES\\(CORPORATIVO\\)\"([^\"]*)\"$")
    public void verficiarELESTADODELASSERIESCORPORATIVO(String dato) throws Throwable {
        Corporativo.verificarElEstadoDeAsignaciónDeSerieCorporativos(dato);
    }

    @When("^SELECCIONAR IMPRESION DE DOCUMENTOS\\(CORPORATIVO\\)$")
    public void seleccionarIMPRESIONDEDOCUMENTOSCORPORATIVO() throws Throwable {
        Corporativo.seleccionamosImpresionDeDocumentosCorporativo();
    }

    @And("^BUSCAR ID DE RESERVA(\\d+)\\(CORPORATIVO\\)\"([^\"]*)\"$")
    public void buscarIDDERESERVACORPORATIVO(String casoDePrueba) throws Throwable {
        Corporativo.buscamosElIdDeReservaCorporativo_2(casoDePrueba);
    }

    @And("^SELECCIONAR CONTRATO DIGITAL\\(CORPORATIVO\\)$")
    public void seleccionarCONTRATODIGITALCORPORATIVO() throws Throwable{
        Corporativo.contratoDigitalCorporativo();
    }

    @And("^SELECCIONAR PREPARACION DE LA FACTURA\\(CORPORATIVO\\)$")
    public void seleccionarPREPARACIONDELAFACTURACORPORATIVO() throws Throwable{
        Corporativo.preparaciónDeLaFacturaCorporativo();
    }

    @And("^IMPRIMIR FACTURA\\(CORPORATIVO\\)$")
    public void imprimirFACTURACORPORATIVO()throws Throwable {
        Corporativo.impresiónDeLaFacturaCorporativo();
    }

    @And("^EJECUTAR INFORMES\\(CORPORATIVO\\)$")
    public void ejecutarINFORMESCORPORATIVO() throws Throwable{
        Corporativo.ejecutarInformesCorporativo();
    }

    @And("^IMPRIMIR GUIA DE REMISION\\(CORPORATIVO\\)$")
    public void imprimirGUIADEREMISIONCORPORATIVO() throws Throwable{
        Corporativo.imprimirGuiaDeRemisiónCorporativo();
    }

    @And("^IMPRIMIR ETIQUETA DE LINEA\\(CORPORATIVO\\)$")
    public void imprimirETIQUETADELINEACORPORATIVO() throws Throwable{
        Corporativo.imprimirEtiquetaDeLineaCorporativo();
    }

    @And("^IMPRIMIR ETIQUETA\\(CORPORATIVO\\)$")
    public void imprimirETIQUETACORPORATIVO() throws Throwable{
        Corporativo.imprimirEtiquetaCorporativo();
    }

    @Then("^VERIFICAR CAMBIO DE ESTADO$")
    public void verificarCAMBIODEESTADO() throws Throwable{
        Corporativo.observamosElCambioDeEstado();
    }

    @When("^SELECCIONAR DESPACHO DE HUB\\(CORPORATIVO\\)$")
    public void seleccionarDESPACHODEHUBCORPORATIVO() throws Throwable{
        Corporativo.seleccionamosDespachoDeHubCorporativo();
    }

    @And("^GENERAR MASTER BOX\\(CORPORATIVO\\)$")
    public void generarMASTERBOXCORPORATIVO()throws Throwable {
        Corporativo.generaciónDeMasterBOXCorporativo();
    }

    @Then("^DESPACHAR PEDIDO\\(CORPORATIVO\\)$")
    public void despacharPEDIDOCORPORATIVO()throws Throwable {
        Corporativo.despacharPedidoCorporativo();
    }

    @When("^SELECCIONAR RECEPCION DE PEDIDOS\\(CORPORATIVO\\)$")
    public void seleccionarRECEPCIONDEPEDIDOSCORPORATIVO()throws Throwable {
        Corporativo.seleccionamosRecepciónDePedidosCorporativo();
    }

    @And("^RECEPCIONAR PEDIDOS\\(CORPORATIVO\\)$")
    public void recepcionarPEDIDOSCORPORATIVO() throws Throwable{
        Corporativo.recepcionarDePedidosCorporativo();
    }

    @And("^EJECUTAR PROCESO DE LOGICA DE RUTEO\\(CORPORATIVO\\)$")
    public void ejecutarPROCESODELOGICADERUTEOCORPORATIVO() throws Throwable{
        Corporativo.procesoDeLógicaDeRuteoCorporativo();
    }

    @And("^EJECUTAR REPORTE DE LOGICA DE RUTEO\\(CORPORATIVO\\)$")
    public void ejecutarREPORTEDELOGICADERUTEOCORPORATIVO()throws Throwable {
        Corporativo.reporteDeLogicaDeRuteoCorporativo();
    }

    @When("^Seleccionamos recepción de pedidos\\(corporativo\\)$")
    public void seleccionamosRecepciónDePedidosCorporativo() throws Throwable{
        Corporativo.seleccionamosRecepciónDePedidosCorporativos();
    }

    @And("^EJECUTAR CARGA LOGICA DE RUTEO \\(CORPORATIVO\\)$")
    public void ejecutarCARGALOGICADERUTEOCORPORATIVO() throws Throwable{
        Corporativo.ejecutarCargaDeRuteoCorporativo();
    }

    @Then("^GUARDAMOS NUMERO DE ENVIO \\(CORPORATIVO\\)$")
    public void guardamosNUMERODEENVIOCORPORATIVO()throws Throwable {
        Corporativo.guardamosNúmeroDeEnvioCorporativo();
    }

    @When("^DESPACHO A MOTORIZADO \\(CORPORATIVO\\)$")
    public void despachoAMOTORIZADOCORPORATIVO()throws Throwable {
        Corporativo.despachoAMotorizadoCorporativo();
    }

    @And("^BUSCAMOS NUMERO DE ENVIO \\(CORPORATIVO\\)$")
    public void buscamosNUMERODEENVIOCORPORATIVO()throws Throwable {
        Corporativo.buscamosNumeroDeEnvioCorporativo();
    }


    @And("^BUSCAMOS MOTORIZADO \\(CORPORATIVO\\)$")
    public void buscamosMOTORIZADOCORPORATIVO()throws Throwable {
        Corporativo.buscamosMotorizadoCorporativo();
    }

    @And("^DESPACHAMOS PEDIDO \\(CORPORATIVO\\)$")
    public void despachamosPEDIDOCORPORATIVO() throws Throwable{
        Corporativo.despachamosPedidoCorporativo();
    }

    @When("^MAESTRO DE PEDIDOS \\(CORPORATIVO\\)$")
    public void maestroDEPEDIDOSCORPORATIVO() throws Throwable{
        Corporativo.maestroDePedidosCorporativo();
    }

    @And("^BUSCAMOS EL ED RESERVA MAESTRO DE PEDIDOS\\(CORPORATIVO\\)\"([^\"]*)\"$")
    public void buscamosELEDRESERVAMAESTRODEPEDIDOSCORPORATIVO(String arg0) throws Throwable {
      Corporativo.buscamosElIDReservaMaestroDePedidoCorporativo(arg0);
    }

    @Then("^VERIFICAMOS EL ESTADO DE PEDIDO \\(CORPORATIVO\\)$")
    public void verificamosELESTADODEPEDIDOCORPORATIVO()throws Throwable {
        Corporativo.verificamosElEstadoDelPedidoCorporativo();
    }

    @And("^ENTREGAR PEDIDO \\(CORPORATIVO\\)$")
    public void entregarPEDIDOCORPORATIVO() throws Throwable{
        Corporativo.entregarPedidoCorporativo();
    }

    @Then("^VALIDAR ESTADO DE PEDIDO \\(CORPORATIVO\\)$")
    public void validarESTADODEPEDIDOCORPORATIVO()throws Throwable {
        Corporativo.validarEstadoDelPedidoCorporativo();
    }

    @When("^SELECCIONAMOS ASIGNACION DE SERIES\\(CORPORATIVO\\)$")
    public void seleccionamosASIGNACIONDESERIESCORPORATIVO() throws Throwable{
        Corporativo.seleccionamosAsignaciónDeSeriesCorporativos();
    }

    @And("^BUSCAMOS PEDIDO PARA ASIGNAR SERIES\\(CORPORATIVO\\)\"([^\"]*)\"$")
    public void buscamosPEDIDOPARAASIGNARSERIESCORPORATIVO(String casoDePrueba) throws Throwable {
       Corporativo.buscamosElPedidoCorporativo(casoDePrueba);
    }

    @And("^INGRESAMOS MATERIALES\\(CORPORATIVO\\) \"([^\"]*)\"$")
    public void ingresamosMATERIALESCORPORATIVO(String arg0) throws Throwable {
        Corporativo.ingresamosMaterialesIMEIYSIMCARDS(arg0);
    }

    @And("^VALIDAMOS SERIES INGRESADAS \\(CORPORATIVO\\)$")
    public void validamosSERIESINGRESADASCORPORATIVO() throws Throwable {
        Corporativo.validamosSeriesCorporativos();
    }

    @And("^BUSCAMOS PEDIDO\\(CORPORATIVO\\)\"([^\"]*)\"$")
    public void buscamosPEDIDOCORPORATIVO(String arg0) throws Throwable {
        Corporativo.buscamosElPedidoCorporativos(arg0);
    }

    @Then("^VERIFICAMOS ESTADO DE SERIES \\(CORPORATIVO\\)\"([^\"]*)\"$")
    public void verificamosESTADODESERIESCORPORATIVO(String arg0) throws Throwable {
        Corporativo.verificarElEstadoDeAsignaciónDeSeriesCorporativos(arg0);
    }

    @When("^SELECCIONAR IMPRRESION DE DOCUMENTOS\\(CORPORATIVO\\)$")
    public void seleccionarIMPRRESIONDEDOCUMENTOSCORPORATIVO() throws Throwable{
        Corporativo.seleccionamosImpresionDeDocumentosCorporativos();
    }

    @And("^Contrato digital \\(corporativo\\)$")
    public void contratoDigitalCorporativo()throws Throwable {
        Corporativo.ContratoDigitalCorporativo();
    }

    @And("^Preparación de la factura \\(corporativo\\)$")
    public void preparaciónDeLaFacturaCorporativo() throws Throwable {
        Corporativo.preparaciónDeLaFacturaCorporativos();
    }

    @And("^Impresión de la factura \\(corporativo\\)$")
    public void impresiónDeLaFacturaCorporativo() throws Throwable{
        Corporativo.impresiónDeLaFacturaCorporativos();
    }

    @And("^Ejecutar informes \\(corporativo\\)$")
    public void ejecutarInformesCorporativo()throws Throwable {
        Corporativo.ejecutarInformesCorporativos();
    }

    @And("^Imprimir guia de remisión \\(corporativo\\)$")
    public void imprimirGuiaDeRemisiónCorporativo()throws Throwable {
        Corporativo.imprimirGuiaDeRemisiónCorporativos();
    }

    @And("^Imprimir etiqueta de linea \\(corporativo$")
    public void imprimirEtiquetaDeLineaCorporativo() throws Throwable{
        Corporativo.imprimirEtiquetaDeLineaCorporativos();
    }

    @And("^Imprimir etiqueta \\(corporativo\\)$")
    public void imprimirEtiquetaCorporativo()throws Throwable {
        Corporativo.imprimirEtiquetaCorporativos();
    }

    @Then("^Observamos el cambio de estado$")
    public void observamosElCambioDeEstado() throws Throwable{
        Corporativo.observamosElCambioDeEstadoC();
    }

    @When("^Seleccionamos despacho de hub\\(corporativo\\)$")
    public void seleccionamosDespachoDeHubCorporativo()throws Throwable {
        Corporativo.seleccionamosDespachoDeHubCorporativos();
    }

    @And("^Ejecución de informes\\(corporativo\\)$")
    public void ejecuciónDeInformesCorporativo() throws Throwable{
        Corporativo.ejecuciónDeInformesCorporativos();
    }

    @And("^Generación de Master BOX \\(corporativo\\)$")
    public void generaciónDeMasterBOXCorporativo()throws Throwable {
        Corporativo.generaciónDeMasterBOXCorporativos();
    }

    @Then("^Despachar pedido \\(corporativo\\)$")
    public void despacharPedidoCorporativo()throws Throwable {
        Corporativo.despacharPedidoCorporativos();
    }


    @When("^Seleccionar recepción de pedidos \\(corporativo\\)$")
    public void seleccionarRecepciónDePedidosCorporativo() throws Throwable {
        Corporativo.seleccionRecepciónDePedidosCorporativos();
    }

    @And("^Recepcionar de pedidos \\(corporativo\\)$")
    public void recepcionarDePedidosCorporativo() throws Throwable{
        Corporativo.recepcionarDePedidosCorporativos();
    }

    @And("^Proceso de lógica de ruteo \\(corporativo\\)$")
    public void procesoDeLógicaDeRuteoCorporativo() throws Throwable{
        Corporativo.procesoDeLógicaDeRuteoCorporativos();
    }


    @And("^Reporte de logica de ruteo \\(corporativo\\)$")
    public void reporteDeLogicaDeRuteoCorporativo()throws Throwable{
        Corporativo.ReporteDeLogicaDeRuteoCorporativos();
    }

    @When("^Seleccionamos recepción de pedidos C \\(corporativo\\)$")
    public void seleccionamosRecepciónDePedidosCCorporativo()throws Throwable {
        Corporativo.SeleccionamosRecepciónDePedidosCorporativos();
    }

    @And("^Ejecutar Carga de Ruteo \\(corporativo\\)$")
    public void ejecutarCargaDeRuteoCorporativo()throws Throwable  {
        Corporativo.EjecutarCargaDeRuteoCorporativos();
    }

    @Then("^Guardamos número de envio \\(corporativo\\)$")
    public void guardamosNúmeroDeEnvioCorporativo() throws Throwable{
        Corporativo.GuardamosNúmeroDeEnvioCorporativos();
    }

    @When("^Despacho a motorizado \\(corporativo\\)$")
    public void despachoAMotorizadoCorporativo()throws Throwable {
        Corporativo.DespachoAMotorizadoCorporativos();
    }

    @And("^Buscamos numero de envio \\(corporativo\\)$")
    public void buscamosNumeroDeEnvioCorporativo() throws Throwable{
        Corporativo.BuscamosNumeroDeEnvioCorporativos();
    }

    @And("^Buscamos motorizado \\(corporativo\\)$")
    public void buscamosMotorizadoCorporativo() throws Throwable{
        Corporativo.BuscamosMotorizadoCorporativos();
    }

    @And("^Despachamos Pedido \\(corporativo\\)$")
    public void despachamosPedidoCorporativo() throws Throwable{
        Corporativo.DespachamosPedidoCorporativos();
    }

    @When("^Maestro de Pedidos \\(corporativo\\)$")
    public void maestroDePedidosCorporativo() throws Throwable{
        Corporativo.MaestroDePedidosCorporativos();
    }


    @And("^Buscamos el numero de pedido\"([^\"]*)\"$")
    public void buscamosElNumeroDePedido(String arg0) throws Throwable {
       Corporativo.BuscamosElNumeroDePedidoC(arg0);
    }

    @Then("^Verificamos el estado del pedido \\(corporativo\\)$")
    public void verificamosElEstadoDelPedidoCorporativo()throws Throwable {
        Corporativo.VerificamosElEstadoDelPedidoCorporativos();
    }

    @And("^Entregar pedido \\(corporativo\\)$")
    public void entregarPedidoCorporativo() throws Throwable{
        Corporativo.EntregarPedidoCorporativos();
    }

    @Then("^Validar estado del pedido \\(corporativo\\)$")
    public void validarEstadoDelPedidoCorporativo()throws Throwable {
        Corporativo.ValidarEstadoDelPedidoCorporativos();
    }
}
