package com.tsoft.bot.frontend.steps.WebDelivery;

import com.tsoft.bot.frontend.helpers.Hook;
import cucumber.api.PendingException;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import org.openqa.selenium.WebDriver;
import com.tsoft.bot.frontend.pages.pages.P_Residential;

public class CargaDeMateriales {
    public WebDriver driver;
    P_Residential CargaMateriales = new P_Residential(driver);

    public CargaDeMateriales() {
        this.driver = Hook.getDriver();
    }

    @Given("^INGRESAR A LA URL DE WEB DELIVERY \"([^\"]*)\"$")
    public void ingresarALAURLDEWEBDELIVERY(String casoPrueba) throws Throwable {
        CargaMateriales.ingresoALaUrlDeWEBDELIVERY(casoPrueba);
    }
    @When("^INGRESAR USUARIO A WEB DELIVERY\"([^\"]*)\"$")
    public void ingresarUSUARIOAWEBDELIVERY(String casoPrueba) throws Throwable {
        CargaMateriales.ingresoElUsuarioDeWEBDELIVERY(casoPrueba);
    }
    @And("^INGRESAR CONTRASENA WEB DELIVER\"([^\"]*)\"$")
    public void ingresarCONTRASENAWEBDELIVER(String casoPrueba) throws Throwable {
        CargaMateriales.laContraseñaDeWEBDELIVERY(casoPrueba);
    }
    @Then("^CLICK EN EL BOTON LOGIN INGRESANDO CORRECTAMENTE A LA PAGINA$")
    public void clickENELBOTONLOGININGRESANDOCORRECTAMENTEALAPAGINA()  throws Throwable{
        CargaMateriales.seDaClicEnElBotonLoginDeWEBDELIVERYIngresandoCorrectamente();
    }
    @Given("^CLICK EN EL BOTON IR A EN WEB DELIVERY \"([^\"]*)\"$")
    public void clickENELBOTONIRAENWEBDELIVERY(String casoPrueba) throws Throwable {
        CargaMateriales.seDaClickEnElBotonIRAEnWEBDELIVERY(casoPrueba);
    }
    @When("^SELECCIONAR AJUSTE DE INVENTARIO$")
    public void seleccionarAJUSTEDEINVENTARIO() throws Throwable{
        CargaMateriales.seleccionarAjusteDeInventario();
    }
    @And("^CLICK EN EL BOTON NUEVO REGISTRO$")
    public void clickENELBOTONNUEVOREGISTRO() throws Throwable {
        CargaMateriales.clickEnElBotonNuevoRegistro();
    }
    @And("^INGRESAR TIPO ABASTECIMIENTO \"([^\"]*)\"$")
    public void ingresarTIPOABASTECIMIENTO(String casoPrueba) throws Throwable {
        CargaMateriales.seleccionamosElTipoABASTECIMIENTO(casoPrueba);
    }
    @And("^INGRESAR COMENTARIO \"([^\"]*)\"$")
    public void ingresarCOMENTARIO(String casoPrueba) throws Throwable {
      CargaMateriales.ingresamosUnComentario(casoPrueba);
    }
    @And("^INGRESAR GUIA DE REMISION \"([^\"]*)\"$")
    public void ingresarGUIADEREMISION(String casoPrueba) throws Throwable {
       CargaMateriales.ingresamosGuiaDeRemision(casoPrueba);
    }
    @And("^CARGAR ARCHIVO CSV$")
    public void cargarARCHIVOCSV() throws Throwable{
        CargaMateriales.ingresamosElArchivo();
    }
    @And("^EJECUTAR AJUSTE Y ACEPTAR MENSAJE$")
    public void ejecutarAJUSTEYACEPTARMENSAJE()throws Throwable {
        CargaMateriales.clickEnEjecutarAjusteYAceptarMensaje();
    }
    @And("^VALIDAR LA CARGA DE ARCHIVO CSV$")
    public void validarLACARGADEARCHIVOCSV() throws Throwable{
        CargaMateriales.validarQueLosArchivosHayanCargado();
    }
}
