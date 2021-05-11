package com.tsoft.bot.frontend.steps.WebDelivery;

import com.tsoft.bot.frontend.pages.pages.P_Residential;
import cucumber.api.PendingException;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import org.openqa.selenium.WebDriver;

public class Residential {
    public WebDriver driver;
    P_Residential Residential = new P_Residential(driver);

    @When("^SELECCIONAMOS AUDITORIA DE PEDIDO \"([^\"]*)\"$")
    public void seleccionamosAUDITORIADEPEDIDO(String casoDePrueba) throws Throwable {
        Residential.seleccionamosAUDITORIADEPEDIDO(casoDePrueba);
    }

    @And("^BUSCAMOS ID DE ORDEN \"([^\"]*)\"$")
    public void buscamosIDDEORDEN(String casoDePrueba) throws Throwable {
        Residential.buscamosIDDEORDEN(casoDePrueba);
    }

    @And("^SELECCIONAMOS PEDIDO$")
    public void seleccionamosPEDIDO() throws Throwable{
        Residential.seleccionamosPEDIDO();
    }

    @Given("^AGENDAMOS EL PEDIDO \"([^\"]*)\"$")
    public void agendamosELPEDIDO(String arg0) throws Throwable {
        Residential.agendamosELPEDIDO(arg0);
    }

    @When("^SELECCIONAMOS FECHA DE PEDIDO\"([^\"]*)\"$")
    public void seleccionamosFECHADEPEDIDO(String arg0) throws Throwable {
        Residential.seleccionamosFECHADEPEDIDO(arg0);
    }

    @Then("^VALIDAMOS CAMBIO DE ESTADO DEL PEDIDO \\(AGENDADO\\) \"([^\"]*)\"$")
    public void validamosCAMBIODEESTADODELPEDIDOAGENDADO(String arg0) throws Throwable {
        Residential.validamosCAMBIODEESTADODELPEDIDOAGENDADO(arg0);
    }

    @When("^SELECCIONAMOS ASIGNACION DE SERIES \"([^\"]*)\"$")
    public void seleccionamosASIGNACIONDESERIES(String arg0) throws Throwable {
        Residential.seleccionamosASIGNACIONDESERIES(arg0);
    }

    @And("^BUSCAMOS EL ORDER ID  \"([^\"]*)\"$")
    public void buscamosELORDERID(String casoDePrueba) throws Throwable {
        Residential.buscamosELORDERID(casoDePrueba);
    }

    @And("^INGRESAMOS MATERIALES \"([^\"]*)\"$")
    public void ingresamosMATERIALES(String casoDePrueba) throws Throwable {
        Residential.ingresamosMATERIALES(casoDePrueba);
    }

    @Given("^VALIDAMOS SERIES$")
    public void validamosSERIES() throws Throwable{
        Residential.validamosSERIES();
    }

    @Then("^VERIFICAMOS ESTADO DE VALIDACION DE SERIES \"([^\"]*)\"$")
    public void verificamosESTADODEVALIDACIONDESERIES(String casoDePrueba) throws Throwable {
        Residential.verificamosESTADODEVALIDACIONDESERIES(casoDePrueba);
    }

    @When("^SELECCIONAMOS IMPRESION DE DOCUMENTOS\"([^\"]*)\"$")
    public void seleccionamosIMPRESIONDEDOCUMENTOS(String arg0) throws Throwable {
        Residential.seleccionamosIMPRESIONDEDOCUMENTOS(arg0);
    }

    @And("^BUSCAMOS EL ORDER_ID  \"([^\"]*)\"$")
    public void buscamosELORDER_ID(String casoDePrueba) throws Throwable {
        Residential.buscamosELORDER_ID(casoDePrueba);
    }

    @And("^SELECCIONAMOS EJECUTAR INFORMES \"([^\"]*)\"$")
    public void seleccionamosEJECUTARINFORMES(String arg0) throws Throwable {
        Residential.seleccionamosEJECUTARINFORMES(arg0);
    }

    @Given("^IMPRESION DE GUIA DE REMISION$")
    public void impresionDEGUIADEREMISION() throws Throwable{
        Residential.impresionDEGUIADEREMISION();
    }

    @And("^IMPRESION DE ETIQUETA  \"([^\"]*)\"$")
    public void impresionDEETIQUETA(String arg0) throws Throwable {
        Residential.impresionDEETIQUETA(arg0);
    }

    @Then("^VERIFICAMOS QUE EL ESTADO DE LA ORDEN SEA EL CORRECTO \\(REALIZADO\\)$")
    public void verificamosQUEELESTADODELAORDENSEAELCORRECTOREALIZADO() throws Throwable{
        Residential.verificamosQUEELESTADODELAORDENSEAELCORRECTOREALIZADO();
    }

    @When("^SELECCIONAMOS DESPACHO DE HUB \"([^\"]*)\"$")
    public void seleccionamosDESPACHODEHUB(String arg0) throws Throwable {
        Residential.seleccionamosDESPACHODEHUB(arg0);
    }

    @And("^ASIGNAMOS MASTER BOX \"([^\"]*)\"$")
    public void asignamosMASTERBOX(String arg0) throws Throwable {
        Residential.asignamosMASTERBOX(arg0);
    }

    @Then("^VERIFICAMOS LA ASIGNACION CORRECTA DEL CODIGO$")
    public void verificamosLAASIGNACIONCORRECTADELCODIGO() throws Throwable{
        Residential.verificamosLAASIGNACIONCORRECTADELCODIGO();
    }

    @And("^DESPACHO DE PEDIDO$")
    public void despachoDEPEDIDO()  throws Throwable{
        Residential.despachoDEPEDIDO();
    }

    @When("^RECEPCIONAR PEDIDOS$")
    public void recepcionarPEDIDOS() throws Throwable{
        Residential.recepcionarPEDIDOS();
    }

    @And("^RECEPCIONAR PEDIDO RESI$")
    public void recepcionarPEDIDORESI() throws Throwable {
        Residential.recepcionarPEDIDORESI();
    }

    @And("^EJECUTAR CARGA DE LOGICA DE RUTEO$")
    public void ejecutarCARGADELOGICADERUTEO() throws Throwable{
        Residential.ejecutarCARGADELOGICADERUTEO();
    }

    @And("^GUARDAR NUMERO DE ENVIO$")
    public void guardarNUMERODEENVIO() throws Throwable{
        Residential.guardarNUMERODEENVIO();
    }

    @When("^DESPACHO DE PEDIDO ENVIADO$")
    public void despachoDEPEDIDOENVIADO() throws Throwable{
        Residential.despachoDePedidoDeEnvio();
    }

    @And("^BUSCAMOS NUMERO DE ENVIO$")
    public void buscamosNUMERODEENVIO() throws Throwable{
        Residential.buscamosNumeroDeEnvio();
    }

    @And("^BUSCAMOS MOTORIZADO$")
    public void buscamosMOTORIZADO() throws Throwable{
        Residential.buscamosMotorizado();
    }

    @And("^DESPACHAMOS PEDIDO$")
    public void despachamosPEDIDO() throws Throwable {
        Residential.despachamosPedido();
    }

    @Then("^VALIDAMOS EL ESTADO DE PEDIDO DESPACHADO$")
    public void validamosELESTADODEPEDIDODESPACHADO() throws Throwable {
        Residential.validamosElEstadoDePedidoDESPACHADO();
    }

}
