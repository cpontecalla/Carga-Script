package com.tsoft.bot.frontend.steps.APP_AppVentas;

import com.tsoft.bot.frontend.helpers.Hook;
import com.tsoft.bot.frontend.utility.ExcelReader;
import com.tsoft.bot.frontend.utility.ExtentReportUtil;
import com.tsoft.bot.frontend.utility.GenerateWord;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import io.appium.java_client.AppiumDriver;
import io.appium.java_client.MobileElement;
import org.openqa.selenium.By;

import java.util.HashMap;
import java.util.List;

import static com.tsoft.bot.frontend.pageobject.APP_AppVentas.PageObject_AppVentas_Login.*;

public class steps_appVentas {

    private static final String EXCEL_APK = "excel/App_Ventas_Login.xlsx";
    private static final String LOGIN_APK = "Login";
    private static final String COLUMNA_DNI = "DNI_VENDEDOR";
    private static final String COLUMNA_DNI2 = "DNI_CLIENTE";
    private static final String COLUMNA_CHIP = "SERIE_CHIP";


    private static GenerateWord generateWord = new GenerateWord();
    private AppiumDriver<MobileElement> driver;

    public steps_appVentas() {
        this.driver = Hook.getDriver();
    }

    private List<HashMap<String, String>> getData() throws Throwable {
        return ExcelReader.data(EXCEL_APK, LOGIN_APK);
    }

    @Given("^Se ingresa a la apk y se ingresa el DNI del vendedor \"([^\"]*)\"$")
    public void seIngresaALaApkYSeIngresaElDNIDelVendedor(String casoDePrueba) throws Throwable {
        try {
            int apk = Integer.parseInt(casoDePrueba) - 1;
            String dni = getData().get(apk).get(COLUMNA_DNI);
            driver.findElement(By.id(TXT_DNI)).clear();
            driver.findElement(By.id(TXT_DNI)).sendKeys(dni);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se inició correctamente la apk y se ingreso el vendedor");
            generateWord.sendText("Se inició correctamente la apk y se ingreso el vendedor");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL_APK, LOGIN_APK, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }


    @When("^se clic en el boton ingresar$")
    public void seClicEnElBotonIngresar() throws Exception {
        try {
            driver.findElement(By.id(BTN_INGRESAR)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en el botón Ingresar");
            generateWord.sendText("Se da click en el botón Ingresar");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL_APK, LOGIN_APK, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }



    @When("^se da click en el boton Guardar de App Ventas$")
    public void seDaClickEnElBotonGuardarDeAppVentas() throws Exception {
        try {
            driver.findElement(By.id(BTN_GUARDAR)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en el botón Guardar Configuración");
            generateWord.sendText("Se da click en el botón Guardar Configuración");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL_APK, LOGIN_APK, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se da click en el boton Acepto del Aviso$")
    public void seDaClickEnElBotonAceptoDelAviso() throws Exception {
        try {
            driver.findElement(By.id(BTN_ACEPTO)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en el botón Acepto del Aviso");
            generateWord.sendText("Se da click en el botón Acepto del Aviso");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL_APK, LOGIN_APK, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se elige la venta Prepago$")
    public void seEligeLaVentaPrepago() throws Exception {
        try {
            driver.findElement(By.id(BTN_PREPAGO)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en la opción PREPAGO");
            generateWord.sendText("Se da click en la opción PREPAGO");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL_APK, LOGIN_APK, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se elige la operacion Nueva Linea$")
    public void seEligeLaOperacionNuevaLinea() throws Exception {
        try {
            driver.findElement(By.id(BTN_NUEVALINEA)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en la operación NUEVA LINEA");
            generateWord.sendText("Se da click en la operación NUEVA LINEA");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL_APK, LOGIN_APK, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se ingresa el numero de documento \"([^\"]*)\"$")
    public void seIngresaElNumeroDeDocumento(String casoDePrueba) throws Throwable {
        try {
            int apk = Integer.parseInt(casoDePrueba) - 1;
            String dni = getData().get(apk).get(COLUMNA_DNI2);
            driver.findElement(By.id(TXT_DOC)).clear();
            driver.findElement(By.id(TXT_DOC)).sendKeys(dni);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresa el DNI del cliente");
            generateWord.sendText("Se ingresa el DNI del cliente");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL_APK, LOGIN_APK, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se da click en el boton CHIP SOLO$")
    public void seDaClickEnElBotonCHIPSOLO() throws Exception {
        try {
            driver.findElement(By.id(BTN_CHIPSOLO)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en el botón Chip Solo");
            generateWord.sendText("Se da click en el botón Chip Solo");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL_APK, LOGIN_APK, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se ingresa la serie del chip \"([^\"]*)\"$")
    public void seIngresaLaSerieDelChip(String casoDePrueba) throws Throwable {
        try {
            int apk = Integer.parseInt(casoDePrueba) - 1;
            String chip = getData().get(apk).get(COLUMNA_CHIP);
            driver.findElement(By.id(TXT_SERIECHIP)).clear();
            driver.findElement(By.id(TXT_SERIECHIP)).sendKeys(chip);
            Thread.sleep(8000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresa la serie del CHIP");
            generateWord.sendText("Se ingresa la serie del CHIP");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL_APK, LOGIN_APK, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se da click en el boton continuar$")
    public void seDaClickEnElBotonContinuar() throws Exception {
        try {
            driver.findElement(By.id(BTN_CONTINUAR)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en el botón Continuar");
            generateWord.sendText("Se da click en el botón Continuar");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL_APK, LOGIN_APK, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

}
