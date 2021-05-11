package com.tsoft.bot.frontend.steps.ventas.prepago;

import com.tsoft.bot.frontend.BaseClass;
import com.tsoft.bot.frontend.helpers.mobile.Hook;
import com.tsoft.bot.frontend.utility.ExcelReader;
import com.tsoft.bot.frontend.utility.GenerateWord;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import io.appium.java_client.AppiumDriver;
import io.appium.java_client.MobileElement;

import java.util.HashMap;
import java.util.List;

import static com.tsoft.bot.frontend.pageobject.Ventas.PageObject_Ventas.*;

public class Steps_AltaPrepago extends BaseClass {
    private static AppiumDriver<MobileElement> driver;
    private static GenerateWord generateWord = new GenerateWord();
    private static final String EXCEL_DOC = "excel/Ventas.xlsx";
    private static final String PAGE_NAME = "Alta-Prepago";
    private static final String COLUMN_TIPO_DOCUMENTO = "TipoDocumento";
    private static final String COLUMN_DOCUMENTO = "Documento";
    private static final String COLUMN_SERIE = "Serie";
    private static final String COLUMNA_TELEFONO = "Telefono";
    private static final String COLUMN_PREPLAN = "PrePlan";

    public Steps_AltaPrepago() throws Throwable {
        driver = Hook.getDriver();
    }

    private List<HashMap<String, String>> getData() throws Throwable {
        return ExcelReader.data(EXCEL_DOC, PAGE_NAME);
    }
    String GET_SERIE = getData().get(0).get(COLUMN_SERIE);
    String GET_TIPO_DOCUMENTO = getData().get(0).get(COLUMN_TIPO_DOCUMENTO);
    String GET_DOCUMENTO = getData().get(0).get(COLUMN_DOCUMENTO);
    String GET_TELEFONO = getData().get(0).get(COLUMNA_TELEFONO);
    String GET_COLUMN_PREPLAN = getData().get(0).get(COLUMN_PREPLAN);
    String SELECT_TIPO_DOCUMENTO = "//android.widget.TextView[@text='" + GET_TIPO_DOCUMENTO + "']";

    @Then("^Se da clic en el boton Acepto y muestra menu de productos$")
    public void seDaClicEnElBotonAceptoYMuestraMenuDeProductos() throws Throwable {
        try {
            click(driver,"id", BTN_ACEPTAR_CONFIGURACION);
            if (isDisplayed(driver,"id", FORM_PRODUCTOS))
            {
                System.out.println("[LOG] - Se valida ingreso al menu principal y se da clic en Prepago");
                stepPass(driver,"Se valida ingreso al menu principal y se da clic en Prepago");
                generateWord.sendText("Se valida ingreso al menu principal y se da clic en Prepago");
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

    @Given("^Se da clic en el boton prepago muestra menu de operaciones$")
    public void seDaClicEnElBotonPrepagoMuestraMenuDeOperaciones() throws Throwable {
        try {
            click(driver,"id", BTN_PREPAGO);
            if (isDisplayed(driver,"id", LBL_GENERAL))
            {
                System.out.println("[LOG] - Se valida el acceso al menú de operaciones y se da clic en Nueva Línea");
                stepPass(driver,"Se valida el acceso al menú de operaciones y se da clic en Nueva Línea");
                generateWord.sendText("Se valida el acceso al menú de operaciones y se da clic en Nueva Línea");
                generateWord.addImageToWord(driver);
            }
            else
            {
                System.out.println("[LOG] - No muestra formulario de operaciones");
                stepFail(driver,"No muestra formulario de operaciones");
                generateWord.sendText("No muestra formulario de operaciones");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            System.out.println("[LOG] - Error en tiempo de respuesta " + we.getMessage());
            stepFail(driver,"Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta");
            generateWord.addImageToWord(driver);
        }
    }

    @When("^Se selecciona nueva linea muestra formulario datos del cliente$")
    public void seSeleccionaNuevaLineaMuestraFormularioDatosDelCliente() throws Throwable {
        try {
            click(driver,"id", BTN_NUEVA_LINEA);
            if (isDisplayed(driver,"id", LBL_ESCOGER_MODALIDAD))
            {
                //click(driver,"xpath", SELECT_DOCUMENTO);
                //click(driver,"xpath", SELECT_TIPO_DOCUMENTO);
                clear(driver,"id", TXT_DOCUMENTO);
                //sendKeyValue(driver,"id", SELECT_DOCUMENTO, GET_TIPO_DOCUMENTO);
                sendKeyValue(driver,"id", TXT_DOCUMENTO, GET_DOCUMENTO);
                System.out.println("[LOG] - Muestra formulario de Datos del cliente, se ingresa documento y clic Chip solo");
                stepPass(driver,"Muestra formulario de Datos del cliente, se ingresa documento y clic Chip solo");
                generateWord.sendText("Muestra formulario de Datos del cliente, se ingresa documento y clic Chip solo");
                generateWord.addImageToWord(driver);
            }
            else
            {
                System.out.println("[LOG] - No muestra formulario de datos del cliente");
                stepFail(driver,"No muestra formulario de datos del cliente");
                generateWord.sendText("No muestra formulario de datos del cliente");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            System.out.println("[LOG] - Error en tiempo de respuesta " + we.getMessage());
            stepFail(driver,"Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^Se da clic en el boton Chip Solo muestra formulario de codigo de barras$")
    public void seDaClicEnElBotonChipSoloMuestraFormularioDeCodigoDeBarras() throws Throwable {
        try {
            click(driver,"id", BTN_CHIP_SOLO);
            if (isDisplayed(driver,"id", LBL_CODIGO_BARRAS))
            {
                clear(driver,"id", TXT_SERIE_CHIP);
                sendKeyValue(driver,"id", TXT_SERIE_CHIP, GET_SERIE);
                System.out.println("Serie: " + GET_SERIE);
                System.out.println("[LOG] - Muestra formulario de Código de Barras, se ingresa serie del chip");
                stepPass(driver,"Muestra formulario de Código de Barras, se ingresa serie del chip");
                if (isDisplayed(driver,"id", BTN_AVISO))
                {
                    click(driver,"id", BTN_AVISO);
                }
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

    @Then("^Se genera nro telefono y se selecciona preplan$")
    public void seGeneraNroTelefonoYSeSeleccionaPreplan() throws Throwable {
        sleep(5000);
        try {
            if (isDisplayed(driver,"id", TXT_NUEVO_NUMERO))
            {
                sendKeyValue(driver,"id", TXT_NUEVO_NUMERO, GET_TELEFONO);
                if (GET_COLUMN_PREPLAN.equals("Prepan Flex"))
                {
                    clear(driver,"xpath", CHBX_PREPLAN_FLEX);
                    sleep(5000);
                    click(driver,"xpath", CHBX_PREPLAN_FLEX);
                    stepPass(driver,"Ventana de Datos plan preplan flex");
                    generateWord.sendText("Ventana de Datos plan preplan flex");
                    generateWord.addImageToWord(driver);
                    sleep(2000);
                    click(driver,"id", BTN_CONTINUAR);
                    click(driver,"id", BTN_AVISO);
                }else
                {
                    clear(driver,"xpath", CHBX_TARIFA_UNICA);
                    sleep(5000);
                    click(driver,"xpath", CHBX_TARIFA_UNICA);
                    stepPass(driver,"Ventana de Datos plan tarifa única");
                    generateWord.sendText("Ventana de Datos plan tarifa única");
                    generateWord.addImageToWord(driver);
                    sleep(2000);
                    click(driver,"id", BTN_CONTINUAR);
                    click(driver,"id", BTN_AVISO);
                }
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

    @And("^Se selecciona centro poblado y muestra huellero$")
    public void seSeleccionaCentroPobladoYMuestraHuellero() throws Throwable {
        sleep(10000);
        try {
            if (isDisplayed(driver,"id", LBL_GENERAL))
            {
                stepPass(driver,"Muestra Centro poblado y continuar");
                generateWord.sendText("Muestra Centro poblado y continuar");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_CONTINUAR);
                sleep(3000);
                if (isDisplayed(driver,"id", TXT_NUEVO_NUMERO))
                {
                    stepPass(driver,"Proceso Exitoso - Clic aceptar con huella");
                    generateWord.sendText("Proceso Exitoso - Clic aceptar con huella");
                    generateWord.addImageToWord(driver);
                }
                else{
                    stepFail(driver,"No muestra proceso aceptar con huella");
                    generateWord.sendText("No muestra proceso aceptar con huella");
                    generateWord.addImageToWord(driver);
                }
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
