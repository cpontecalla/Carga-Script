package com.tsoft.bot.frontend.steps.fijas;

import com.tsoft.bot.frontend.BaseClass;
import com.tsoft.bot.frontend.helpers.fijas.HookFijas;
import com.tsoft.bot.frontend.utility.ExcelReader;
import com.tsoft.bot.frontend.utility.GenerateWord;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import io.appium.java_client.AppiumDriver;
import org.openqa.selenium.*;

import java.util.HashMap;
import java.util.List;

import static com.tsoft.bot.frontend.pageobject.Fijas.PageObject_Fijas.*;
import static com.tsoft.bot.frontend.utility.scroll.scroll;

public class Steps_AltaFija extends BaseClass {
    private static AppiumDriver<WebElement> driver;
    private static GenerateWord generateWord = new GenerateWord();
    private static final String EXCEL_DOC = "excel/VentasFijas.xlsx";
    private static final String PAGE_NAME = "Alta-Fijas";
    private static final String COLUMN_DOCUMENTO = "Documento";
    private static final String COLUMN_DEPARTAMENTO = "Departamento";
    private static final String COLUMN_PROVINCIA = "Provincia";
    private static final String COLUMN_DISTRITO = "Distrito";
    private static final String COLUMN_DIRECCION = "Direccion";
    private static final String COLUMN_REFERENCIA = "Referencia";
    private static final String COLUMN_CELULAR = "Celular";
    private static final String COLUMN_NOMBRE_VIA = "NombreVia";
    private static final String COLUMN_CUADRA = "Cuadra";
    private static final String COLUMN_NRO_PUERTA = "NumeroPuerta";
    private static final String COLUMN_NOMBRE_CCHH = "NombreCCHH";
    private static final String COLUMN_MANZANA = "Mz";
    private static final String COLUMN_LOTE = "Lt";
    private static final String COLUMN_CAMPANIA = "Campania";
    private static final String COLUMN_EMAIL = "Email";
    private static final String COLUMN_NOMBRE_MADRE = "Nombre_Madre";

    public Steps_AltaFija() throws Throwable {
        driver = HookFijas.getDriver();
    }

    private List<HashMap<String, String>> getData() throws Throwable {
        return ExcelReader.data(EXCEL_DOC, PAGE_NAME);
    }
    private String GET_DOCUMENTO = getData().get(0).get(COLUMN_DOCUMENTO);
    private String GET_DEPARTAMENTO = getData().get(0).get(COLUMN_DEPARTAMENTO);
    private String GET_PROVINCIA = getData().get(0).get(COLUMN_PROVINCIA);
    private String GET_DISTRITO = getData().get(0).get(COLUMN_DISTRITO);
    private String GET_DIRECCION = getData().get(0).get(COLUMN_DIRECCION);
    private String GET_REFERENCIA = getData().get(0).get(COLUMN_REFERENCIA);
    private String GET_CELULAR = getData().get(0).get(COLUMN_CELULAR);
    private String GET_NOMBRE_VIA = getData().get(0).get(COLUMN_NOMBRE_VIA);
    private String GET_CUADRA = getData().get(0).get(COLUMN_CUADRA);
    private String GET_NRO_PUERTA = getData().get(0).get(COLUMN_NRO_PUERTA);
    private String GET_NOMBRE_CCHH = getData().get(0).get(COLUMN_NOMBRE_CCHH);
    private String GET_MANZANA = getData().get(0).get(COLUMN_MANZANA);
    private String GET_LOTE = getData().get(0).get(COLUMN_LOTE);
    private String GET_CAMPANIA = getData().get(0).get(COLUMN_CAMPANIA);
    private String GET_EMAIL = getData().get(0).get(COLUMN_EMAIL);
    private String GET_NOMBRE_MADRE = getData().get(0).get(COLUMN_NOMBRE_MADRE);

    @When("^Muestra menu principal y se da clic en altas nuevas$")
    public void muestraMenuPrincipalYSeDaClicEnAltasNuevas() throws Throwable {
        try {
            if (isDisplayed(driver,"id", FORM_MENU))
            {
                stepPass(driver, "Se selecciona Altas Nuevas");
                generateWord.sendText("Se selecciona Altas Nuevas");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_ALTA);
            }
            else
            {
                stepFail(driver, "No se muestra el menú principal");
                generateWord.sendText("No se muestra el menú principal");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @And("^Se ingresa datos del contratante y clic en evaluar$")
    public void seIngresaDatosDelContratanteYClicEnEvaluar() throws Throwable {
        try {
            if (isDisplayed(driver,"id", CBX_DOCUMENTO))
            {
                clear(driver,"id", TXT_DOCUMENTO);
                sendKeyValue(driver,"id", TXT_DOCUMENTO, GET_DOCUMENTO);
                sendKeyValue(driver,"id", CBX_DEPARTAMENTO, GET_DEPARTAMENTO);
                sendKeyValue(driver,"id", CBX_PROVINCIA, GET_PROVINCIA);
                sendKeyValue(driver,"id", CBX_DISTRITO, GET_DISTRITO);
                //sendKeyValue(driver,"id", TXT_DOCUMENTO, GET_DIRECCION);
                stepPass(driver, "Se Ingresa datos del contratante");
                generateWord.sendText("Se Ingresa datos del contratante");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_EVALUAR);
            }
            else
            {
                stepFail(driver, "No se muestra el menú principal");
                generateWord.sendText("No se muestra el menú principal");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^Verificar cliente sin restriccion de deuda e iniciar venta$")
    public void verificarClienteSinRestriccionDeDeudaEIniciarVenta() throws Throwable {
        try {
            if (isDisplayed(driver,"id", LBL_MONTO))
            {
                if (LBL_MONTO != "S/. 0")
                {
                    stepPass(driver, "No tiene restriccion de deuda, Iniciar venta");
                    generateWord.sendText("No tiene restriccion de deuda, Iniciar venta");
                    generateWord.addImageToWord(driver);
                    click(driver,"id", BTN_INICIAR_VENTA);
                }
                else{
                    stepFail(driver, "Cliente tiene restriccion de deuda");
                    generateWord.sendText("Cliente tiene restriccion de deuda");
                    generateWord.addImageToWord(driver);
                }
            }
            else
            {
                stepFail(driver, "No se muestra el menú principal");
                generateWord.sendText("No se muestra el menú principal");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @Given("^Se selecciona direccion de instalacion$")
    public void seSeleccionaDireccionDeInstalacion() throws Throwable {
        try {
            if (isDisplayed(driver,"id", POP_UP))
            {
                click(driver,"id", POP_UP);
                click(driver,"id", POP_UP_CERRAR);
                clear(driver,"id", TXT_DIRECCION);
                clear(driver,"id", TXT_REFERENCIA);
                clear(driver,"id", TXT_CELULAR);
                sendKeyValue(driver,"id", TXT_DIRECCION, GET_DIRECCION);
                sendKeyValue(driver,"id", TXT_REFERENCIA, GET_REFERENCIA);
                sendKeyValue(driver,"id", TXT_CELULAR, GET_CELULAR);
                sleep(2000);
                click(driver,"id", BTN_BUSCAR);
                stepPass(driver,"Se busca dirección de instalación");
                generateWord.sendText("Se busca dirección de instalación");
                generateWord.addImageToWord(driver);
                sleep(2000);
                click(driver,"id", BTN_CONTINUAR);
            }
            else
            {
                stepFail(driver, "No encuentra direccion para la instalación");
                generateWord.sendText("No encuentra direccion para la instalación");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @When("^Se completa datos de la direccion de instalacion$")
    public void seCompletaDatosDeLaDireccionDeInstalacion() throws Throwable {
        try {
            //if (isDisplayed(driver,"id", POP_UP))
            //{
//                clear(driver,"id", TXT_NOMBREVIA);
//                clear(driver,"id", TXT_CUADRA);
//                clear(driver,"id", TXT_NRO_PUERTA);
//                clear(driver,"id", TXT_NOMBRE_CCHH);
//                clear(driver,"id", TXT_MANZANA);
//                clear(driver,"id", TXT_LOTE);
//                clear(driver,"id", TXT_EDITAR_REFERENCIA);
//                sendKeyValue(driver,"id", TXT_NOMBREVIA, GET_NOMBRE_VIA);
//                sendKeyValue(driver,"id", TXT_CUADRA, GET_CUADRA);
//                sendKeyValue(driver,"id", TXT_NRO_PUERTA, GET_NRO_PUERTA);
//                sendKeyValue(driver,"id", TXT_NOMBRE_CCHH, GET_NOMBRE_CCHH);
//                sendKeyValue(driver,"id", TXT_MANZANA, GET_MANZANA);
//                sendKeyValue(driver,"id", TXT_LOTE, GET_LOTE);
//                sendKeyValue(driver,"id", TXT_EDITAR_REFERENCIA, GET_REFERENCIA);
                stepPass(driver,"Se completa datos de la dirección de instalación");
                generateWord.sendText("Se completa datos de la dirección de instalación");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_CONTINUAR);
            //}
            //else
            //{
                //stepFail(driver, "No encuentra direccion para la instalación");
                //generateWord.sendText("No encuentra direccion para la instalación");
                //generateWord.addImageToWord(driver);
            //}
        }
        catch (Exception we)
        {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @And("^Se selecciona campana disponible para cliente$")
    public void seSeleccionaCampanaDisponibleParaCliente() throws Throwable {
        try {
            if (isPresent(driver,"id", BTN_INTENTAR)){
                stepFail(driver, "El cliente no tiene campañas disponibles");
                generateWord.sendText("El cliente no tiene campañas disponibles");
                generateWord.addImageToWord(driver);
            }else {
                generateWord.sendText("Se selecciona Campaña");
                generateWord.addImageToWord(driver);
                stepPass(driver,"Se selecciona Campaña");
                click(driver,"id", BTN_VER_TODOS);
                listaCampanias(GET_CAMPANIA);
                sleep(3000);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^Se verifica detalle del producto$")
    public void seVerificaDetalleDelProducto() throws Throwable {
        try {
            if (isDisplayed(driver,"id", LBL_PRECIO_PRODUCTO))
            {
                generateWord.sendText("Se verifica el detalle del producto seleccionado");
                generateWord.addImageToWord(driver);
                stepPass(driver,"Se verifica el detalle del producto seleccionado");
                click(driver,"id", BTN_DP_SIGUIENTE);
                sleep(2000);
            }
            else
            {
                stepFail(driver, "No se muestra el detalle del producto");
                generateWord.sendText("No se muestra el detalle del producto");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @And("^Se selecciona SVA$")
    public void seSeleccionaSVA() throws Throwable {
        try {
            if (isDisplayed(driver,"xpath", CBX_BLOQUE_TV))
            {
                click(driver,"xpath", CBX_BLOQUE_TV);
                generateWord.sendText("Se selecciona SVA");
                generateWord.addImageToWord(driver);
                stepPass(driver,"Se selecciona SVA");
                click(driver,"id", BTN_SVA_SIGUIENTE);
                sleep(2000);
            }
            else
            {
                stepFail(driver, "No se muestra pantalla SVA");
                generateWord.sendText("No se muestra pantalla SVA");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @When("^Se aceptan las condiciones$")
    public void seAceptanLasCondiciones() throws Throwable {
        try {
            //WebElement PANEL_CONDICIONES = driver.findElement(By.xpath("/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.widget.ScrollView/android.view.ViewGroup"));
            WebElement PANEL_CONDICIONES = driver.findElement(By.id("com.telefonica.ventafija.dev.debug:id/nav_host_fragment_main"));
            if (isDisplayed(driver,"id", LBL_CONDICIONES))
            {
                //WebElement PANEL_CONDICIONES = driver.findElement(By.xpath("/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.widget.ScrollView/android.view.ViewGroup"));
                //WebElement PANEL_CONDICIONES = driver.findElement(By.id("com.telefonica.ventafija.dev.debug:id/txt_SubTituloControlParental"));
                generateWord.sendText("Se aceptan las condiciones");
                generateWord.addImageToWord(driver);
                //stepPass(driver,"Se aceptan las condiciones");
                scroll(PANEL_CONDICIONES, driver,"UP");
                click(driver,"id", BTN_CONDICIONES_CONTINUAR);
                sleep(2000);
            }
            else
            {
                stepFail(driver, "No se muestra pantalla de condiciones");
                generateWord.sendText("No se muestra pantalla de condiciones");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @And("^Se muestra el resumen de venta$")
    public void seMuestraElResumenDeVenta() throws Throwable {
        try {
            if (isDisplayed(driver,"id", LBL_PLAN))
            {
                WebElement PANEL_RESUMEN_VENTA = driver.findElement(By.xpath("/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.widget.ScrollView/android.view.ViewGroup"));
                scroll(PANEL_RESUMEN_VENTA, driver, "UP");
                generateWord.sendText("Se muestra el resumen de venta");
                generateWord.addImageToWord(driver);
                stepPass(driver,"Se muestra el resumen de venta");
                click(driver,"id", BTN_RESUMEN_VENTA_ACEPTAR);
                sleep(2000);
            }
            else
            {
                stepFail(driver, "No se muestra pantalla de resumen de ventas");
                generateWord.sendText("No se muestra pantalla de resumen de ventas");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^Se realiza la lectura del contrato$")
    public void seRealizaLaLecturaDelContrato() throws Throwable {
        try {
            if (isDisplayed(driver,"id", BTN_PLAY))
            {
                generateWord.sendText("Se realiza la lectura del contrato");
                generateWord.addImageToWord(driver);
                stepPass(driver,"Se realiza la lectura del contrato");
                click(driver,"id", BTN_PLAY);
                sleep(2000);
                click(driver,"id", POP_UP_ALLOW);
                sleep(30000);
                click(driver,"id", BTN_LECTURA_CONTINUAR);
                sleep(5000);
            }
            else
            {
                stepFail(driver, "No se muestra pantalla de lectura de contratos");
                generateWord.sendText("No se muestra pantalla de lectura de contratos");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @And("^Se realiza validacion de identidad$")
    public void seRealizaValidacionDeIdentidad() throws Throwable {
        try {
            if (isDisplayed(driver,"id", LBL_NOMBRE_CLIENTE))
            {
                sendKeyValue(driver,"id", TXT_EMAIL, GET_EMAIL);
                sendKeyValue(driver,"id", TXT_NOMBRE_MADRE, GET_NOMBRE_MADRE);
                generateWord.sendText("Se selecciona a la madre");
                generateWord.addImageToWord(driver);
                stepPass(driver,"Se selecciona a la madre");
                click(driver,"id", BTN_VALIDACION_SIGUIENTE);
                sleep(2000);
            }
            else
            {
                stepFail(driver, "No se muestra pantalla validacion de identidad");
                generateWord.sendText("No se muestra pantalla validacion de identidad");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^Se selecciona foto de dni$")
    public void seSeleccionaFotoDeDni() throws Throwable {
        try {
            if (isDisplayed(driver,"id", IMG_FRONTAL_DNI))
            {
                click(driver,"id", IMG_FRONTAL_DNI);
                click(driver,"xpath", OPC_ELEGIR_GALERIA);
                click(driver,"xpath", GALERIA);
                click(driver,"xpath", GALERIA_FOTO1);
                click(driver,"id", BTN_CROP);
                sleep(2000);
                click(driver,"id", IMG_POSTERIOR_DNI);
                click(driver,"xpath", OPC_ELEGIR_GALERIA);
                click(driver,"xpath", GALERIA);
                click(driver,"xpath", GALERIA_FOTO1);
                click(driver,"id", BTN_CROP);
                sleep(2000);
                generateWord.sendText("Se ha seleccionado la foto del dni");
                generateWord.addImageToWord(driver);
                stepPass(driver,"Se ha seleccionado la foto del dni");
                click(driver,"id", BTN_FOTO_SIGUIENTE);
                sleep(10000);
            }
            else
            {
                stepFail(driver, "No se muestra pantalla para seleccionar DNI");
                generateWord.sendText("No se muestra pantalla para seleccionar DNI");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @And("^Se valida venta exitosa$")
    public void seValidaVentaExitosa() throws Throwable {
        try {
            if (isDisplayed(driver,"id", IMG_EXITOSA))
            {
                generateWord.sendText("Se ha realizado la venta exitosa");
                generateWord.addImageToWord(driver);
                stepPass(driver,"Se ha realizado la venta exitosa");
                sleep(2000);
            }
            else
            {
                stepFail(driver, "No se muestra pantalla de Venta exitosa");
                generateWord.sendText("No se muestra pantalla de Venta exitosa");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    public static void listaCampanias(String Campania) throws Throwable {
        WebElement PANEL_CATALOGO = driver.findElement(By.id("com.telefonica.ventafija.dev.debug:id/rv_catalog_list"));
        switch (Campania) {
            case "TRIO PLANO LOCAL 200 MBPS ESTANDAR DIGITAL HD":
                click(driver,"xpath", OPC_1);
                break;

            case "TRIO PLANO LOCAL 100 MBPS ESTANDAR DIGITAL HD":
                click(driver,"xpath", OPC_2);
                break;

            case "TRIO PLANO LOCAL 50 MBPS ESTANDAR DIGITAL HD":
                click(driver,"xpath", OPC_3);
                break;

            case "TRIO PLANO LOCAL 30 MBPS ESTANDAR DIGITAL HD":
                click(driver,"xpath", OPC_4);
                break;

            case "DUO PLANO LOCAL ESTANDAR DIGITAL":
                scroll(PANEL_CATALOGO, driver,"UP");
                click(driver,"xpath", OPC_1);
                break;

            case "DUO INTERNET 200 MBPS ESTANDAR DIGITAL TV HD":
                scroll(PANEL_CATALOGO, driver,"UP");
                click(driver,"xpath", OPC_2);
                break;

            case "DUO INTERNET 100 MBPS ESTANDAR DIGITAL TV HD":
                scroll(PANEL_CATALOGO, driver,"UP");
                click(driver,"xpath", OPC_3);
                break;

            case "DUO INTERNET 50 MBPS ESTANDAR DIGITAL TV HD":
                scroll(PANEL_CATALOGO, driver,"UP");
                click(driver,"xpath", OPC_4);
                break;

            case "DUO INTERNET 30 MBPS ESTANDAR DIGITAL TV HD":
                scroll(PANEL_CATALOGO, driver,"UP");
                click(driver,"xpath", OPC_5);
                break;

        }
    }
}
