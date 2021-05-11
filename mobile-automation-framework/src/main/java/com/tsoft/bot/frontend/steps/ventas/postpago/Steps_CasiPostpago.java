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

public class Steps_CasiPostpago extends BaseClass {
    private static AppiumDriver<MobileElement> driver;
    private static GenerateWord generateWord = new GenerateWord();
    private static final String EXCEL_DOC = "excel/Ventas.xlsx";
    private static final String PAGE_NAME = "Casi-Postpago";
    private static final String COLUMN_TIPO_DOCUMENTO = "TipoDocumento";
    private static final String COLUMN_DOCUMENTO = "Documento";
    private static final String COLUMN_TELEFONO = "Telefono";
    private static final String COLUMN_CORREO = "Correo";
    private static final String COLUMN_SERIE = "Serie";
    private static final String COLUMN_PLAN = "Plan";
    private static final String COLUMN_DIRECCION = "Direccion";

    public Steps_CasiPostpago() throws Throwable {
        driver = Hook.getDriver();
    }
    private List<HashMap<String, String>> getData() throws Throwable {
        return ExcelReader.data(EXCEL_DOC, PAGE_NAME);
    }
    String GET_TIPO_DOCUMENTO = getData().get(0).get(COLUMN_TIPO_DOCUMENTO);
    String GET_DOCUMENTO = getData().get(0).get(COLUMN_DOCUMENTO);
    String GET_TELEFONO = getData().get(0).get(COLUMN_TELEFONO);
    String GET_CORREO = getData().get(0).get(COLUMN_CORREO);
    String GET_SERIE = getData().get(0).get(COLUMN_SERIE);
    String GET_PLAN = getData().get(0).get(COLUMN_PLAN);
    String GET_DIRECCION = getData().get(0).get(COLUMN_DIRECCION);

    @Then("^Muestra menu de operaciones y se selecciona Postpago$")
    public void muestraMenuDeOperacionesYSeSeleccionaPostpago() throws Throwable {
        try {
            click(driver,"id", BTN_ACEPTAR_CONFIGURACION);
            if (isDisplayed(driver,"id", FORM_PRODUCTOS))
            {
                stepPass(driver,"Acceso al menú principal - Postpago");
                generateWord.sendText("Acceso al menú principal - Postpago");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_POSTPAGO);
            }
        }
        catch (Throwable we)
        {
            stepFail(driver,"Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^Muestra opciones de postpago y Se selecciona Renovacion$")
    public void muestraOpcionesDePostpagoYSeSeleccionaRenovacion() throws Throwable {
        try {
            if (isDisplayed(driver,"id", LBL_GENERAL))
            {
                stepPass(driver,"Se selecciona la opción renovación");
                generateWord.sendText("Se selecciona la opción renovación");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_RENOVACION);
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

    @When("^Se ingresa datos del cliente postpago se da clic en chip solo$")
    public void seIngresaDatosDelClientePostpagoSeDaClicEnChipSolo() throws Throwable {
        try {
            if (isDisplayed(driver,"id", FORM_CONFIGURACION))
            {
                clear(driver,"id", TXT_DOCUMENTO);
                clear(driver,"id", TXT_NUEVO_NUMERO);
                clear(driver,"id", GET_CORREO);
                //sendKeyValue(driver,"id", SELECT_DOCUMENTO, GET_TIPO_DOCUMENTO);
                sendKeyValue(driver,"id", TXT_DOCUMENTO, GET_DOCUMENTO);
                sendKeyValue(driver,"id", TXT_NUEVO_NUMERO, GET_TELEFONO);
                sendKeyValue(driver,"id", TXT_EMAIL, GET_CORREO);
                stepPass(driver,"Muestra formulario de Datos del cliente y clic Chip solo");
                generateWord.sendText("Muestra formulario de Datos del cliente y clic Chip solo");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_CHIP_SOLO);
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

    @And("^Se ingresa serie plan y clic en continuar$")
    public void seIngresaSeriePlanYClicEnContinuar() throws Throwable {
        try {
            if (isDisplayed(driver,"id", LBL_CODIGO_BARRAS))
            {
                clear(driver,"id", TXT_SERIE_CHIP);
                sendKeyValue(driver,"id", TXT_SERIE_CHIP, GET_SERIE);
                sleep(5000);
                sendKeyValue(driver,"id", CBX_PLAN, GET_PLAN);
                stepPass(driver,"Se ingresa la serie y el plan");
                generateWord.sendText("Se ingresa la serie y el plan");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_CONTINUAR);
            }
            else
            {
                stepFail(driver,"No muestra formulario de código de barras");
                generateWord.sendText("No muestra formulario de código de barras");
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

    @Then("^Se ingresa datos de facturacion$")
    public void seIngresaDatosDeFacturacion() throws Throwable {
        try {
            if (isDisplayed(driver,"id", FORM_CONFIGURACION))
            {
                clear(driver,"id", TXT_DIRECCION);
                sendKeyValue(driver,"id", TXT_DIRECCION, GET_DIRECCION);
                sleep(5000);
                stepPass(driver,"Se ingresa la serie y el plan");
                generateWord.sendText("Se ingresa la serie y el plan");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_CONTINUAR);
            }
            else
            {
                stepFail(driver,"No muestra formulario de código de barras");
                generateWord.sendText("No muestra formulario de código de barras");
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

    @And("^Muestra contrato y huellero$")
    public void muestraContratoYHuellero() throws Throwable {
        try {
            if (isDisplayed(driver,"id", TXT_CONDICIONES))
            {
                stepPass(driver,"Proceso Exitoso - Clic aceptar con huella");
                generateWord.sendText("Proceso Exitoso - Clic aceptar con huella");
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
