package com.tsoft.bot.frontend.steps.ventas.postpago;

import com.tsoft.bot.frontend.BaseClass;
import com.tsoft.bot.frontend.helpers.mobile.Hook;
import com.tsoft.bot.frontend.utility.ExcelReader;
import com.tsoft.bot.frontend.utility.GenerateWord;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import io.appium.java_client.AppiumDriver;
import io.appium.java_client.MobileElement;

import java.util.HashMap;
import java.util.List;

import static com.tsoft.bot.frontend.pageobject.Ventas.PageObject_Ventas.*;

public class Steps_PortabilidadPostpago extends BaseClass {
    private static AppiumDriver<MobileElement> driver;
    private static GenerateWord generateWord = new GenerateWord();
    private static final String EXCEL_DOC = "excel/Ventas.xlsx";
    private static final String PAGE_NAME = "Portabilidad-Prepago";
    private static final String COLUMN_TIPO_DOCUMENTO = "TipoDocumento";
    private static final String COLUMN_DOCUMENTO = "Documento";
    private static final String COLUMN_TELEFONO = "Telefono";
    private static final String COLUMN_OPERADOR = "Operador";
    private static final String COLUMN_PRODUCTO_ACTUAL = "ProductoActual";
    private static final String COLUMN_CORREO = "Correo";
    private static final String COLUMN_SERIE = "Serie";
    private static final String COLUMN_DIRECCION = "Direccion";

    public Steps_PortabilidadPostpago() throws Throwable {
        driver = Hook.getDriver();
    }
    private List<HashMap<String, String>> getData() throws Throwable {
        return ExcelReader.data(EXCEL_DOC, PAGE_NAME);
    }
    String GET_TIPO_DOCUMENTO = getData().get(0).get(COLUMN_TIPO_DOCUMENTO);
    String GET_DOCUMENTO = getData().get(0).get(COLUMN_DOCUMENTO);
    String GET_TELEFONO = getData().get(0).get(COLUMN_TELEFONO);
    String GET_OPERADOR = getData().get(0).get(COLUMN_OPERADOR);
    String GET_PRODUCTO_ACTUAL = getData().get(0).get(COLUMN_PRODUCTO_ACTUAL);
    String GET_CORREO = getData().get(0).get(COLUMN_CORREO);
    String GET_SERIE = getData().get(0).get(COLUMN_SERIE);
    String GET_DIRECCION = getData().get(0).get(COLUMN_DIRECCION);

    @Then("^Muestra menu de operaciones se selecciona Postpago$")
    public void muestraMenuDeOperacionesSeSeleccionaPostpago() throws Throwable {
        try {
            click(driver,"id", BTN_ACEPTAR_CONFIGURACION);
            if (isDisplayed(driver,"id", FORM_PRODUCTOS))
            {
                stepPass(driver,"Se valida ingreso al menu principal y se da clic en Postpago");
                generateWord.sendText("Se valida ingreso al menu principal y se da clic en Postpago");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_POSTPAGO);
            }
            else{
                stepFail(driver, "No muestra el menú principal");
                generateWord.sendText("No muestra el menú principal");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Throwable we)
        {
            stepFail(driver,"Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^Muestra opciones de postpago y Se selecciona Portabilidad$")
    public void muestraOpcionesDePostpagoYSeSeleccionaPortabilidad() throws Throwable {
        try {
            if (isDisplayed(driver,"id", LBL_GENERAL))
            {
                stepPass(driver,"Se selecciona la opción portabilidad");
                generateWord.sendText("Se selecciona la opción portabilidad");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_PORTABILIDAD);
            }
            else
            {
                stepFail(driver,"No muestra formulario de operaciones");
                generateWord.sendText("No muestra formulario de operaciones");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver,"Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta");
            generateWord.addImageToWord(driver);
        }
    }

    @When("^Se ingresa datos del cliente y se da clic en validar linea postpago$")
    public void seIngresaDatosDelClienteYSeDaClicEnValidarLineaPostpago() throws Throwable {
        try {
            if (isDisplayed(driver,"id", BTN_VALIDAR_LINEA))
            {
                clear(driver,"id", TXT_DOCUMENTO);
                clear(driver,"id", TXT_NUEVO_NUMERO);
                clear(driver,"id", TXT_EMAIL);
                //sendKeyValue(driver,"id", SELECT_DOCUMENTO, GET_TIPO_DOCUMENTO);
                //sendKeyValue(driver,"id", TXT_DOCUMENTO, GET_DOCUMENTO);
                sendKeyValue(driver,"id", TXT_NUEVO_NUMERO, GET_TELEFONO);
                sendKeyValue(driver,"id", CBX_OPERADOR, GET_OPERADOR);
                sendKeyValue(driver,"id", CBX_OPERADOR_ACTUAL, GET_PRODUCTO_ACTUAL);
                //selectByVisibleText(driver,"id", GET_OPERADOR);
                //selectByVisibleText(driver,"id", GET_PRODUCTO_ACTUAL);
                sendKeyValue(driver,"id", TXT_EMAIL, GET_CORREO);
                stepPass(driver,"Se ingresa datos del cliente y se valida la línea");
                generateWord.sendText("Se ingresa datos del cliente y se valida la línea");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_VALIDAR_LINEA);
            }
            else
            {
                stepFail(driver,"No muestra formulario de datos del cliente");
                generateWord.sendText("No muestra formulario de datos del cliente");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver,"Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^Se ingresa serie y plan clic en continuar$")
    public void seIngresaSerieYPlanClicEnContinuar() throws Throwable {
        try {
            if (isPresent(driver,"id", BTN_AVISO))
            {
                generateWord.sendText("Muestra mensaje de error");
                generateWord.addImageToWord(driver);
                stepFail(driver,"Muestra mensaje de error");
            }
            if (isDisplayed(driver,"id", TXT_SERIE_CHIP))
            {
                clear(driver,"id", TXT_SERIE_CHIP);
                sendKeyValue(driver,"id", TXT_SERIE_CHIP, GET_SERIE);
                sleep(10000);
                stepPass(driver,"Se ingresa datos del cliente y se valida la línea");
                generateWord.sendText("Se ingresa datos del cliente y se valida la línea");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_CONTINUAR);
                click(driver,"id", BTN_AVISO);//-------------------------------------
            }
            else{
                stepFail(driver,"No muestra formulario para ingresar la serie");
                generateWord.sendText("No muestra formulario para ingresar la serie");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver,"Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^Se ingresa datos de centro poblado facturacion y huellero$")
    public void seIngresaDatosDeCentroPoblado() throws Throwable {
        try {
            if (isDisplayed(driver,"id", LBL_GENERAL))
            {
                stepPass(driver,"Muestra Centro poblado y continuar");
                generateWord.sendText("Muestra Centro poblado y continuar");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_CONTINUAR);
                sleep(3000);
                clear(driver,"id", TXT_DIRECCION);
                sendKeyValue(driver,"id", TXT_DIRECCION, GET_DIRECCION);
                stepPass(driver,"Ingresa datos de facturación");
                generateWord.sendText("Ingresa datos de facturación");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_CONTINUAR2);
                stepPass(driver,"Proceso Exitoso - Clic aceptar con huella");
                generateWord.sendText("Proceso Exitoso - Clic aceptar con huella");
                generateWord.addImageToWord(driver);
            }
            else
            {
                stepFail(driver,"No muestra Centro poblado");
                generateWord.sendText("No muestra Centro poblado");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver,"Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta");
            generateWord.addImageToWord(driver);
        }
    }
}
