package com.tsoft.bot.frontend.steps.ventas.postpago;

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
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

import java.util.HashMap;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.concurrent.TimeUnit;
import java.util.function.Function;

import static com.tsoft.bot.frontend.pageobject.Ventas.PageObject_Ventas.*;

public class Steps_AltaPostpago extends BaseClass {
    private static AppiumDriver<MobileElement> driver;
    private static GenerateWord generateWord = new GenerateWord();
    private static final String EXCEL_VENTAS = "excel/Ventas.xlsx";
    private static final String PAGE_POSTPAGO = "Alta-Postpago";
    private static final String COLUMN_TIPO_DOCUMENTO = "TipoDocumento";
    private static final String COLUMNA_DOCUMENTO = "Documento";
    private static final String COLUMNA_EMAIL = "Correo";
    private static final String COLUMNA_SERIE = "Serie";
    private static final String COLUMNA_TELEFONO = "Telefono";
    private static final String COLUMNA_PLAN = "Plan";
    private static final String COLUMNA_DIRECCION = "Direccion";


    private List<HashMap<String, String>> getData() throws Throwable {
        return ExcelReader.data(EXCEL_VENTAS, PAGE_POSTPAGO);
    }

    public Steps_AltaPostpago() throws Throwable {
        this.driver = Hook.getDriver();
    }
    String GET_TIPO_DOCUMENTO = getData().get(0).get(COLUMN_TIPO_DOCUMENTO);
    String GET_DOCUMENTO = getData().get(0).get(COLUMNA_DOCUMENTO);
    String GET_EMAIL = getData().get(0).get(COLUMNA_EMAIL);
    String GET_SERIE = getData().get(0).get(COLUMNA_SERIE);
    String GET_TELEFONO = getData().get(0).get(COLUMNA_TELEFONO);
    String GET_PLAN = getData().get(0).get(COLUMNA_PLAN);
    String GET_DIRECCION = getData().get(0).get(COLUMNA_DIRECCION);
    String SELECT_TIPO_DOCUMENTO = "//android.widget.TextView[@text='" + GET_TIPO_DOCUMENTO + "']";
    String SELECT_PLAN = "//android.widget.CheckedTextView[@text='" + GET_PLAN + "']";

    @Then("^Se da clic en el boton Acepto muestra menu de productos$")
    public void seDaClicEnElBotonAceptoMuestraMenuDeProductos() throws Throwable {
        try {
            click(driver,"id", BTN_ACEPTAR_CONFIGURACION);
            if (isDisplayed(driver,"id", FORM_PRODUCTOS))
            {
                stepPass(driver,"Se valida ingreso al menu principal y se da clic en Postpago");
                generateWord.sendText("Se valida ingreso al menu principal y se da clic en Postpago");
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

    @Given("^Se da clic en el boton postpago muestra menu de operaciones$")
    public void seDaClicEnElBotonPostpagoMuestraMenuDeOperaciones() throws Throwable {
        try {
            click(driver,"id", BTN_POSTPAGO);
            if (isDisplayed(driver,"id", LBL_GENERAL))
            {
                System.out.println("[LOG] - Muestra formulario de operaciones y se da clic en el botón Nueva Línea");
                stepPass(driver,"Muestra formulario de operaciones y se da clic en el botón Nueva Línea");
                generateWord.sendText("Muestra formulario de operaciones y se da clic en el botón Nueva Línea");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_NUEVA_LINEA);
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

    @When("^Se selecciona nueva linea se ingresa datos del cliente$")
    public void seSeleccionaNuevaLineaSeIngresaDatosDelCliente() throws Exception {
        try {
            if (isDisplayed(driver,"id", LBL_ESCOGER_MODALIDAD))
            {
                if (GET_TIPO_DOCUMENTO.equals("DNI"))
                {
                    sendKeyValue(driver,"id", TXT_DOCUMENTO, GET_DOCUMENTO);
                    sendKeyValue(driver,"id", TXT_EMAIL, GET_EMAIL);
                    System.out.println("[LOG] - Muestra formulario de Datos del cliente y se da clic en el botón Chip solo");
                    stepPass(driver,"Muestra formulario de Datos del cliente y se da clic en el botón Chip solo");
                    generateWord.sendText("Muestra formulario de Datos del cliente y se da clic en el botón Chip solo");
                    generateWord.addImageToWord(driver);
                    sleep(2000);
                    click(driver,"id", BTN_CHIP_SOLO);
                    sleep(5000);
                }
                else{
                    click(driver,"xpath", SELECT_DOCUMENTO);
                    click(driver,"xpath", SELECT_TIPO_DOCUMENTO);
                    System.out.println("Valor es: " + SELECT_TIPO_DOCUMENTO);
                    sendKeyValue(driver,"id", TXT_DOCUMENTO, GET_DOCUMENTO);
                    sendKeyValue(driver,"id", TXT_EMAIL, GET_EMAIL);
                    System.out.println("[LOG] - Muestra formulario de Datos del cliente y se da clic en el botón Chip solo");
                    stepPass(driver,"Muestra formulario de Datos del cliente y se da clic en el botón Chip solo");
                    generateWord.sendText("Muestra formulario de Datos del cliente y se da clic en el botón Chip solo");
                    generateWord.addImageToWord(driver);
                    sleep(2000);
                    click(driver,"id", BTN_CHIP_SOLO);
                    sleep(5000);
                }

                try {
                    Boolean element = driver.findElements(By.id(TXT_SERIE_CHIP)).size() >0;
                    if (element){sleep(500);}
                }catch (Exception we){
                    generateWord.sendText("Error con el documento");
                    generateWord.addImageToWord(driver);
                    System.out.println("[LOG] - Error con el documento " + we.getMessage());
                    stepFail(driver,"Error con el documento" + we.getMessage());
                    driver.close();
                }finally {
                    System.out.println("[LOG] - finalizó catch");
                    sleep(500);
                }
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
        } catch (Throwable throwable) {
            throwable.printStackTrace();
        }
    }

    @And("^Se da clic en boton Chip Solo muestra formulario de codigo de barras$")
    public void seDaClicEnBotonChipSoloMuestraFormularioDeCodigoDeBarras() throws Throwable {
        try {
            if (isDisplayed(driver,"id", LBL_CODIGO_BARRAS))
            {
                clear(driver,"id", TXT_SERIE_CHIP);
                sendKeyValue(driver,"id", TXT_SERIE_CHIP, GET_SERIE);
                System.out.println("[LOG] - Muestra formulario de Código de Barras, se ingresa serie del chip " + GET_SERIE);
                stepPass(driver,"Muestra formulario de Código de Barras, se ingresa serie del chip " + GET_SERIE);
                generateWord.sendText("Muestra formulario de Código de Barras y se ingresa serie del chip " + GET_SERIE);
                generateWord.addImageToWord(driver);

                /*Boolean element = driver.findElements(By.id(BTN_AVISO)).size() >0;
                if (element){
                    generateWord.sendText("Muestra mensaje de error");
                    generateWord.addImageToWord(driver);
                    click(driver,"id", BTN_AVISO);
                }*/

                //sendKeyValue(driver,"id", TXT_NUEVO_NUMERO, GET_TELEFONO);
                click(driver,"xpath", CBX_PLAN);
                click(driver,"xpath", SELECT_PLAN);
                System.out.println("[LOG] - Se Selecciona plan: " + GET_PLAN);

                //Boolean element2 = driver.findElements(By.id(BTN_AVISO2)).size() >0;
                /*if (element){
                    generateWord.sendText("Muestra mensaje de error");
                    generateWord.addImageToWord(driver);
                    click(driver,"id", BTN_AVISO2);
                }*/

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

    @And("^Selecciona centro poblado y muestra huellero$")
    public void seleccionaCentroPobladoYMuestraHuellero() throws Throwable {
        //sleep(10000);
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
