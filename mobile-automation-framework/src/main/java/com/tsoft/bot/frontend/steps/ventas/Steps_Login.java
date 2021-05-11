package com.tsoft.bot.frontend.steps.ventas;

import com.tsoft.bot.frontend.BaseClass;
import com.tsoft.bot.frontend.helpers.mobile.Hook;
import com.tsoft.bot.frontend.utility.ExcelReader;
import com.tsoft.bot.frontend.utility.GenerateWord;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.When;
import io.appium.java_client.AppiumDriver;
import io.appium.java_client.MobileElement;

import java.util.HashMap;
import java.util.List;

import static com.tsoft.bot.frontend.pageobject.Ventas.PageObject_Ventas.*;

public class Steps_Login extends BaseClass {

    private static AppiumDriver<MobileElement> driver;
    private static GenerateWord generateWord = new GenerateWord();
    private static final String EXCEL_DOC = "excel/Ventas.xlsx";
    private static final String PAGE_NAME = "Login";
    private static final String COLUMN_DNI = "DNI";

    public Steps_Login() throws Throwable {
        driver = Hook.getDriver();
    }

    private List<HashMap<String, String>> getData() throws Throwable {
        return ExcelReader.data(EXCEL_DOC, PAGE_NAME);
    }
    private String GET_DNI = getData().get(0).get(COLUMN_DNI);

    @Given("^Abrir la aplicación e ingresar número de Dni \"([^\"]*)\"$")
    public void abrirLaAplicacionEIngresarNumeroDeDni(String arg0) throws Throwable {
        try {
            click(driver,"id", POPUP_AUTH);
            sleep(3500);
            click(driver,"id", POPUP_AUTH2);
            sleep(3500);
            clear(driver,"id", TXT_DNI_VENDEDOR);
            sendKeyValue(driver,"id", TXT_DNI_VENDEDOR, GET_DNI);
            stepPass(driver,"Se muestra login y se ingresó nro de documento del vendedor");
            generateWord.sendText("Se muestra login y se ingresó nro de documento del vendedor");
            generateWord.addImageToWord(driver);
        }
        catch (Exception we)
        {
            stepFail(driver,"Error en tiempo de respuesta " + we.getMessage());
            generateWord.sendText("Error en tiempo de respuesta");
            generateWord.addImageToWord(driver);
        }
    }

    @When("^Se da clic al boton ingresar se muestra formulario de configuracion$")
    public void seDaClicAlBotonIngresarSeMuestraFormularioDeConfiguracion() throws Throwable {
        try {
            click(driver,"id", BTN_INGRESAR);
            if (isDisplayed(driver,"id", FORM_CONFIGURACION))
            {
                stepPass(driver, "Se valida formulario de Configuración y se da clic en el botón guardar");
            }
            else
            {
                stepFail(driver, "No se muestra el formulario de configuración");
            }
        } catch (Exception we) {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
        }
    }

    @And("^Se da clic en el boton guardar se muestra pop up$")
    public void seDaClicEnElBotonGuardarSeMuestraPopUp() throws Throwable {
        try {
            click(driver,"id", BTN_GUARDAR);
            sleep(3500);
            stepPass(driver, "Se valida formulario de Configuración y se da clic en el botón Acepto");
        }
        catch (Exception we)
        {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
        }
    }
}
