package com.tsoft.bot.frontend.steps.fijas;

import com.tsoft.bot.frontend.BaseClass;
import com.tsoft.bot.frontend.helpers.fijas.HookFijas;
import com.tsoft.bot.frontend.utility.ExcelReader;
import com.tsoft.bot.frontend.utility.GenerateWord;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import io.appium.java_client.*;
import org.openqa.selenium.*;

import java.util.HashMap;
import java.util.List;

import static com.tsoft.bot.frontend.pageobject.Fijas.PageObject_Fijas.*;
import static com.tsoft.bot.frontend.utility.scroll.*;

public class Steps_Migraciones extends BaseClass {
    private static AppiumDriver<WebElement> driver;
    private static GenerateWord generateWord = new GenerateWord();
    private static final String EXCEL_DOC = "excel/VentasFijas.xlsx";
    private static final String PAGE_NAME = "Migraciones";
    private static final String COLUMN_DOCUMENTO = "Documento";
    private static final String COLUMN_CELULAR = "Celular";
    private static final String COLUMN_NOMBRE_VIA = "NombreVia";
    private static final String COLUMN_CUADRA = "Cuadra";
    private static final String COLUMN_NRO_PUERTA = "NumeroPuerta";
    private static final String COLUMN_NOMBRE_CCHH = "NombreCCHH";
    private static final String COLUMN_MANZANA = "Mz";
    private static final String COLUMN_LOTE = "Lt";

    public Steps_Migraciones() throws Throwable {
        driver = HookFijas.getDriver();
    }

    private List<HashMap<String, String>> getData() throws Throwable {
        return ExcelReader.data(EXCEL_DOC, PAGE_NAME);
    }
    private String GET_DOCUMENTO = getData().get(0).get(COLUMN_DOCUMENTO);
    private String GET_CELULAR = getData().get(0).get(COLUMN_CELULAR);
    private String GET_NOMBRE_VIA = getData().get(0).get(COLUMN_NOMBRE_VIA);
    private String GET_CUADRA = getData().get(0).get(COLUMN_CUADRA);
    private String GET_NRO_PUERTA = getData().get(0).get(COLUMN_NRO_PUERTA);
    private String GET_NOMBRE_CCHH = getData().get(0).get(COLUMN_NOMBRE_CCHH);
    private String GET_MANZANA = getData().get(0).get(COLUMN_MANZANA);
    private String GET_LOTE = getData().get(0).get(COLUMN_LOTE);

    @When("^Muestra menu principal y se da clic en migraciones$")
    public void muestraMenuPrincipalYSeDaClicEnMigraciones() throws Throwable {
        try {
            if (isDisplayed(driver,"id", FORM_MENU))
            {
                stepPass(driver, "Se selecciona Migraciones / Completas");
                generateWord.sendText("Se selecciona Migraciones / Completas");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_MIGRACION);
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
            stepFail(driver, "No se muestra el menu de opciones " + we.getMessage());
            generateWord.sendText("No se muestra el menu de opciones " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @And("^Se busca productos del cliente$")
    public void seBuscaProductosDelCliente() throws Throwable {
        try {
            if (isDisplayed(driver,"id", TXT_DOCUMENTO))
            {
                clear(driver,"id", TXT_DOCUMENTO);
                clear(driver,"id", TXT_CELULAR);
                sendKeyValue(driver,"id", TXT_DOCUMENTO, GET_DOCUMENTO);
                sendKeyValue(driver,"id", TXT_CELULAR, GET_CELULAR);
                click(driver,"id", BTN_BUSCAR_PRODUCTO);
                sleep(10000);
                stepPass(driver, "Muestra productos del cliente");
                generateWord.sendText("Muestra productos del cliente");
                generateWord.addImageToWord(driver);
            }
            else
            {
                stepFail(driver, "No se muestra los productos del cliente");
                generateWord.sendText("No se muestra los productos del cliente");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "No se muestra el formulario de busqueda de productos " + we.getMessage());
            generateWord.sendText("No se muestra el formulario de busqueda de productos " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^Se selecciona opcion linea o television$")
    public void seSeleccionaOpcionLineaOTelevision() throws Throwable {
        try {
            WebElement PANEL = driver.findElement(By.id("com.telefonica.ventafija.dev.debug:id/rv_clients"));
            if (isPresent(driver,"xpath", OPCION_PRODUCTO))
            {
                stepPass(driver,"Muestra productos de la Linea");
                generateWord.sendText("Muestra productos de la Linea");
                generateWord.addImageToWord(driver);
                click(driver,"xpath", OPCION_PRODUCTO);
            }
            else
            {
                scroll(PANEL, driver, "LEFT");
                sleep(3000);
                stepPass(driver,"Muestra productos de television");
                generateWord.sendText("Muestra productos de television");
                generateWord.addImageToWord(driver);
                click(driver,"xpath", OPCION_PRODUCTO);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "No se muestra el formulario de productos " + we.getMessage());
            generateWord.sendText("No se muestra el formulario de productos " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @And("^Se selecciona el producto$")
    public void seSeleccionaElProducto() throws Throwable {
        try {
            if (isPresent(driver,"id", LBL_DIRECCION))
            {
                generateWord.sendText("Muestra detalles del servicio");
                generateWord.addImageToWord(driver);
                stepPass(driver,"Muestra detalles del servicio");
                sleep(2000);
                click(driver,"xpath", BTN_SELECCIONAR);
            }
            else
            {
                stepFail(driver,"No muestra detalles del servicio");
                generateWord.sendText("No muestra detalles del servicio");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "No se muestra el formulario de productos " + we.getMessage());
            generateWord.sendText("No se muestra el formulario de productos " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^Muestra lista de campanias$")
    public void muestraListaDeCampanias() throws Throwable {
        try {
            if (isPresent(driver,"id", LBL_DIRECCION))
            {
                generateWord.sendText("Muestra detalles del servicio");
                generateWord.addImageToWord(driver);
                stepPass(driver,"Muestra detalles del servicio");
                sleep(2000);
                click(driver,"xpath", BTN_SELECCIONAR);
            }
            else
            {
                stepFail(driver,"No muestra detalles del servicio");
                generateWord.sendText("No muestra detalles del servicio");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "No se muestra el formulario de productos " + we.getMessage());
            generateWord.sendText("No se muestra el formulario de productos " + we.getMessage());
            generateWord.addImageToWord(driver);
        }
    }
}
