package com.tsoft.bot.frontend.steps.APP_MiMovistar;

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

import static com.tsoft.bot.frontend.pageobject.APP_MiMovistar.PageObject_MiMovistar.*;


public class steps_appMiMovistar {

    private static final String EXCEL1_TEST1 = "excel/APP_MiMovistar_TEST1.xlsx";
    private static final String PEST_EXCEL1 = "TEST1";
    private static final String COLUMNA_DNI = "DNI";
    private static final String COLUMNA_PASS = "PASSWORD";
    private static final String COLUMNA_NOMBRE = "NUEVO_NOMBRE";
    private static final String COLUMNA_APILLIDO = "NUEVO_APELLIDO";

    private static GenerateWord generateWord = new GenerateWord();
    private AppiumDriver<MobileElement> driver;

    public steps_appMiMovistar() { this.driver = Hook.getDriver(); }

    private List<HashMap<String, String>> getData() throws Throwable {
        return ExcelReader.data(EXCEL1_TEST1, PEST_EXCEL1);
    }




    @Given("^Se ingresa a la app Mi Movistar y se da click al boton empieza ahora$")
    public void seIngresaALaAppMiMovistarYSeDaClickAlBotonEmpiezaAhora() throws Exception {
        try {
            driver.findElement(By.id(BTN_EMPIEZAAHORA)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en el botón empieza ahora");
            generateWord.sendText("Se da click en el botón empieza ahora");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se selecciona el ingreso como titular$")
    public void seSeleccionaElIngresoComoTitular() throws Exception {
        try {
            driver.findElement(By.xpath(BTN_COMOTITULAR)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se selecciona el ingreso como titular");
            generateWord.sendText("Se selecciona el ingreso como titular");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @When("^se ingresa el DNI \"([^\"]*)\"$")
    public void seIngresaElDNI(String casoDePrueba) throws Throwable {
        try {
            int apk = Integer.parseInt(casoDePrueba) - 1;
            String dni = getData().get(apk).get(COLUMNA_DNI);
            driver.findElement(By.xpath(TXT_NUMDOCUMENTO)).click();
            driver.findElement(By.xpath(TXT_NUMDOCUMENTO)).sendKeys(dni);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresa el DNI: "+dni);
            generateWord.sendText("Se ingresa el DNI: "+dni);
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se ingresa la contrasenia de (\\d+) numeros \"([^\"]*)\"$")
    public void seIngresaLaContraseniaDeNumeros(int arg0, String casoDePrueba) throws Throwable {
        try {
            int apk = Integer.parseInt(casoDePrueba) - 1;
            String pass = getData().get(apk).get(COLUMNA_PASS);
            driver.findElement(By.xpath(TXT_CONTRASENIA)).click();
            driver.findElement(By.xpath(TXT_CONTRASENIA)).sendKeys(pass);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresa el password: "+pass);
            generateWord.sendText("Se ingresa el password: "+pass);
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se da click en el boton ingresar$")
    public void seDaClickEnElBotonIngresar() throws Exception {
        try {
            driver.navigate().back();
            driver.findElement(By.xpath(BTN_INGRESAR)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en el botón ingresar");
            generateWord.sendText("Se da click en el botón ingresar");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se da click en ajustes y se selecciona informacion personal$")
    public void seDaClickEnAjustesYSeSeleccionaInformacionPersonal() throws Exception {
        try {
            driver.findElement(By.xpath(BTN_AJUSTES)).click();
            driver.findElement(By.xpath(BTN_INFOPERSONAL)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se selecciona Información Personal");
            generateWord.sendText("Se selecciona Información Personal");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se ingresa a datos personales y al nombre del titular$")
    public void seIngresaADatosPersonalesYAlNombreDelTitular() throws Exception {
        try {
            driver.findElement(By.xpath(BTN_DATOSPERSONALES)).click();
            driver.findElement(By.xpath(BTN_NOMBRE)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se selecciona el nombre del titular");
            generateWord.sendText("Se selecciona el nombre del titular");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se cambiar el nombre del titular por \"([^\"]*)\"$")
    public void seCambiarElNombreDelTitularPor(String casoDePrueba) throws Throwable {
        try {
            int apk = Integer.parseInt(casoDePrueba) - 1;
            String nombre = getData().get(apk).get(COLUMNA_NOMBRE);
            driver.findElement(By.id(TXT_NOMBRE)).clear();
            driver.findElement(By.id(TXT_NOMBRE)).sendKeys(nombre);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresa el nuevo nombre de titular: "+nombre);
            generateWord.sendText("Se ingresa el nuevo nombre de titular: "+nombre);
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se cambia el apellido del titular por \"([^\"]*)\"$")
    public void seCambiaElApellidoDelTitularPor(String casoDePrueba) throws Throwable {
        try {
            int apk = Integer.parseInt(casoDePrueba) - 1;
            String apellido = getData().get(apk).get(COLUMNA_APILLIDO);
            driver.findElement(By.id(TXT_NOMBRE2)).clear();
            driver.findElement(By.id(TXT_NOMBRE2)).sendKeys(apellido);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresa el nuevo apellido de titular: "+apellido);
            generateWord.sendText("Se ingresa el nuevo apellido de titular: "+apellido);
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se da click en el boton guardar$")
    public void seDaClickEnElBotonGuardar() throws Exception {
        try {
            driver.findElement(By.id(BTN_GUARDAR)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en el botón Guardar");
            generateWord.sendText("Se da click en el botón Guardar");
            generateWord.addImageToWord(driver);
            driver.navigate().back();
            driver.navigate().back();
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se realiza el logout de la app$")
    public void seRealizaElLogoutDeLaApp() throws Exception {
        try {
            driver.findElement(By.xpath(BTN_OPCIONES)).click();
            driver.findElement(By.id(SELECT_CERRAR)).click();
            driver.findElement(By.xpath(BTN_CERRAR)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en el botón Cerrar");
            generateWord.sendText("Se da click en el botón Cerrar");
            generateWord.addImageToWord(driver);
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "PASS");
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }


    }


    @Then("^se da click en el boton compra paquetes$")
    public void seDaClickEnElBotonCompraPaquetes() throws Exception {
        try {
            driver.findElement(By.xpath(BTN_COMPRAPAQUETES)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en el botón Compra de Paquetes");
            generateWord.sendText("Se da click en el botón Compra de Paquetes");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se selecciona paquete de datos$")
    public void seSeleccionaPaqueteDeDatos() throws Exception {
        try {
            driver.findElement(By.xpath(BTN_PAQUETESDATOS)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en el botón Compra de Paquetes de Datos");
            generateWord.sendText("Se da click en el botón Compra de Paquetes de Datos");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }


    }

    @And("^se agrega el paquete Instagram Ilim X$")
    public void seAgregaElPaqueteInstagramIlimX() throws Exception {
        try {
            driver.findElement(By.xpath(BTN_ADD_INSTAGRAM)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en el botón Añadir Paquete Instagram");
            generateWord.sendText("Se da click en el botón Añadir Paquete Instagram");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se visualiza las caracteristicas y se da click en el boton pagar$")
    public void seVisualizaLasCaracteristicasYSeDaClickEnElBotonPagar() throws Exception {
        try {
            driver.findElement(By.xpath(BTN_PAGAR)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en el botón Pagar");
            generateWord.sendText("Se da click en el botón Pagar");
            generateWord.addImageToWord(driver);

        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se visualiza la confirmacion de compra y se da click a volver a mi linea$")
    public void seVisualizaLaConfirmacionDeCompraYSeDaClickAVolverAMiLinea() throws Exception {
        try {
            driver.findElement(By.xpath(BTN_VOLVER)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en el botón Volver a mi Linea");
            generateWord.sendText("Se da click en el botón Volver a mi Linea");
            generateWord.addImageToWord(driver);
            Thread.sleep(2000);
            driver.findElement(By.xpath(BTN_AJUSTES)).click();
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se selecciona paquete de minutos$")
    public void seSeleccionaPaqueteDeMinutos() throws Exception {
        try {
            driver.findElement(By.xpath(BTN_PAQUETESMINUTOS)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en el botón Compra de Paquetes de Minutos");
            generateWord.sendText("Se da click en el botón Compra de Paquetes de Minutos");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se agrega el paquete de (\\d+) minutos$")
    public void seAgregaElPaqueteDeMinutos(int arg0) throws Exception {
        try {
            driver.findElement(By.xpath(BTN_ADD_PAQ45MIN)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se da click en el botón Añadir Paquete 45 Minutos");
            generateWord.sendText("Se da click en el botón Añadir Paquete 45 Minutos");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }


    @And("^se da click en ajustes y se selecciona seguridad y privacidad$")
    public void seDaClickEnAjustesYSeSeleccionaSeguridadYPrivasidad() throws Exception {
        try {
            driver.findElement(By.xpath(BTN_AJUSTES)).click();
            driver.findElement(By.xpath(BTN_SEGURIDAD)).click();
            driver.findElement(By.xpath(BTN_GESTIONSESIONES)).wait();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se selecciona Seguridad y privacidad");
            generateWord.sendText("Se selecciona Seguridad y privacidad");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se da click en gestion de sesiones y se valida el dispositivo$")
    public void seDaClickEnGestionDeSesionesYSeValidaElDispositivo() throws Exception {
        try {
            driver.findElement(By.xpath(BTN_GESTIONSESIONES)).click();
            Thread.sleep(5000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se verifica el Dispositivo");
            generateWord.sendText("Se verifica el Dispositivo");
            generateWord.addImageToWord(driver);
            driver.navigate().back();
            driver.navigate().back();
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL1_TEST1, PEST_EXCEL1, 1, 5, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
}
