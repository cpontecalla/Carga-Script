package com.tsoft.bot.frontend.steps.fijas;

import com.tsoft.bot.frontend.BaseClass;
import com.tsoft.bot.frontend.helpers.fijas.HookFijas;
import com.tsoft.bot.frontend.utility.ExcelReader;
import com.tsoft.bot.frontend.utility.GenerateWord;
import cucumber.api.java.en.Given;
import io.appium.java_client.AppiumDriver;
import org.openqa.selenium.WebElement;

import java.util.HashMap;
import java.util.List;

import static com.tsoft.bot.frontend.pageobject.Fijas.PageObject_Fijas.*;
import static com.tsoft.bot.frontend.utility.ConectionClass.*;

public class Steps_LoginFijas extends BaseClass {

    private static AppiumDriver<WebElement> driver;
    private static GenerateWord generateWord = new GenerateWord();
    private static final String EXCEL_DOC = "excel/VentasFijas.xlsx";
    private static final String PAGE_NAME = "Login";
    private static final String COLUMN_CODIGO = "Codigo";
    private static final String COLUMN_PASSWORD = "Password";
    private static final String COLUMN_TOKEN = "Token";

    public Steps_LoginFijas() throws Throwable {
        driver = HookFijas.getDriver();
    }

    private List<HashMap<String, String>> getData() throws Throwable {
        return ExcelReader.data(EXCEL_DOC, PAGE_NAME);
    }
    private String GET_CODIGO = getData().get(0).get(COLUMN_CODIGO);
    private String GET_PASSWORD = getData().get(0).get(COLUMN_PASSWORD);

    @Given("^Abrir la aplicación e ingresar codigo de vendedor \"([^\"]*)\"$")
    public void abrirLaAplicaciónEIngresarCodigoDeVendedor(String arg0) throws Throwable {
        try {
            if (isDisplayed(driver,"id", TXT_CODIGO))
            {
                clear(driver,"id", TXT_CODIGO);
                sendKeyValue(driver,"id", TXT_CODIGO, GET_CODIGO);
                stepPass(driver, "Se ingresa codigo de vendedor: " + GET_CODIGO + " y clic en siguiente");
                generateWord.sendText("Se ingresa codigo de vendedor: " + GET_CODIGO + "  y clic en siguiente");
                generateWord.addImageToWord(driver);
                click(driver,"id", BTN_PRE_LOGUEO);
                if (isPresent(driver,"id", POPUP_AGREE))
                {
                    stepFail(driver, "Error de servicio");
                    generateWord.sendText("Error de servicio");
                    generateWord.addImageToWord(driver);
                }
                else {
                    sleep(5000);
                    String token = executeQuerySelect("select token from ibmx_a07e6d02edaf552.tdp_token_vendedor where codatis = '" + GET_CODIGO + "';");
                    ExcelReader.writeCellValue(EXCEL_DOC, PAGE_NAME, 1, 3, token);
                    sendKeyValue(driver,"id", TXT_PASSWORD, GET_PASSWORD);
                    String GET_TOKEN = getData().get(0).get(COLUMN_TOKEN);
                    sendKeyValue(driver,"id", TXT_TOKEN, GET_TOKEN);
                    stepPass(driver, "Se ingresa codigo de vendedor y clic en siguiente");
                    generateWord.sendText("Se ingresa codigo de vendedor y clic en siguiente");
                    generateWord.addImageToWord(driver);
                    click(driver,"id", BTN_LOGIN);
                    if (isPresent(driver,"id", POPUP_AGREE))
                    {
                        stepFail(driver, "Error de servicio");
                        generateWord.sendText("Error de servicio");
                        generateWord.addImageToWord(driver);
                    }
                    else {
                        sleep(1000);
                    }
                }

//                if (isPresent(driver,"id", POPUP_SAVE_PASSWORD))
//                {
//                    click(driver,"id", POPUP_SAVE_PASSWORD);
//                }
//                else{
//
//                }
            }
            else {
                stepFail(driver, "No se muestra el formulario de Logueo");
                generateWord.sendText("No se muestra el formulario de Logueo");
                generateWord.addImageToWord(driver);
            }
        }
        catch (Exception we)
        {
            stepFail(driver, "Error en tiempo de respuesta " + we.getMessage());
        }
    }
}
