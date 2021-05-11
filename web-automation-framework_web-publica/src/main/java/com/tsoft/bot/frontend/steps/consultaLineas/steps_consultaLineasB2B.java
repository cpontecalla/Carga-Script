package com.tsoft.bot.frontend.steps.consultaLineas;

import com.tsoft.bot.frontend.helpers.Hook;
import com.tsoft.bot.frontend.utility.ExcelReader;
import com.tsoft.bot.frontend.utility.ExtentReportUtil;
import com.tsoft.bot.frontend.utility.GenerateWord;
import com.tsoft.bot.frontend.utility.Sleeper;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;

import java.util.HashMap;
import java.util.List;

import static com.tsoft.bot.frontend.pageobject.consultaLineas.PageObjectCnsltLineas.*;


public class steps_consultaLineasB2B {
    private static final String EXCEL_WEB = "excel/consultaLineas.xlsx";
    private static final String CONSULTA_LINEAS_WEB = "ConsultaLineas";
    private static final String COLUMNA_URL = "URL";
    private static final String COLUMNA_USUARIO= "Usuario";
    private static final String COLUMNA_PASS= "Contraseña";
    private static final String COLUMNA_TIPO_DOCUMENTO = "Tipo_Documento";
    private static final String COLUMNA_DOCUMENTO = "Num_Documento";
    private static final String COLUMNA_TIPO_DOCUMENTORL = "Tipo_DocumentoRL";
    private static final String COLUMNA_DOCUMENTORL  = "Num_DocumentoRL";

    private static GenerateWord generateWord = new GenerateWord();
    private WebDriver driver;
    public steps_consultaLineasB2B() { this.driver = Hook.getDriver(); }

    private List<HashMap<String, String>> getData() throws Throwable {
        return ExcelReader.data(EXCEL_WEB, CONSULTA_LINEAS_WEB);
    }


    @Given("^Ingreso a la url del Portal \"([^\"]*)\"$")
    public void ingresoALaUrlDelPortal(String casoPrueba) throws Throwable {
        // Write code here that turns the phrase above into concrete actions
        try {
            int consLineas = Integer.parseInt(casoPrueba) - 1;
            String url = getData().get(consLineas).get(COLUMNA_URL);
            driver.get(url);

            ExtentReportUtil.INSTANCE.stepPass(driver, "Se inició correctamente la carga del ambiente");
            generateWord.sendText("Se inició correctamente la carga del ambiente");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL_WEB, CONSULTA_LINEAS_WEB, 1, 19, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }

    }

    @When("^Ingreso el Nombre de usuario \"([^\"]*)\"$")
    public void ingresoElNombreDeUsuario(String casoPrueba) throws Throwable {
        try{
            int consLineas = Integer.parseInt(casoPrueba) - 1;
            String Usuario = getData().get(consLineas).get(COLUMNA_USUARIO);
            driver.findElement(TXT_USER).sendKeys(Usuario);
            ExtentReportUtil.INSTANCE.stepPass(driver,"Se ingreso el Usuario Correctamente");
            generateWord.sendText("Usuario Ingresado Correctamente");
            generateWord.addImageToWord(driver);
        }catch (Exception e){
            ExcelReader.writeCellValue(EXCEL_WEB,CONSULTA_LINEAS_WEB,1,0,"FALLO " +
                    "EL INGRESO DEL USUARIO");
            ExtentReportUtil.INSTANCE.stepFail(driver,"Fallo el ingreso de Usuario : " + e.getMessage());
            generateWord.sendText("Usuario no ingresado - Tiempo de espera excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^Ingreso la Contraseña \"([^\"]*)\"$")
    public void ingresoLaContraseña(String casoPrueba) throws Throwable {
        try{
            int consLineas = Integer.parseInt(casoPrueba) - 1;
            String Pass = getData().get(consLineas).get(COLUMNA_PASS);
            driver.findElement(TXT_PASS).sendKeys(Pass);
            ExtentReportUtil.INSTANCE.stepPass(driver,"Se ingreso la contraseña Correctamente");
            generateWord.sendText("Contraseña Ingresada Correctamente");
            generateWord.addImageToWord(driver);
        }catch (Exception e){
            ExcelReader.writeCellValue(EXCEL_WEB,CONSULTA_LINEAS_WEB,1,0,"FALLO " +
                    "EL INGRESO DE LA CONSTRASEÑA");
            ExtentReportUtil.INSTANCE.stepFail(driver,"Fallo el ingreso de Contraseña : " + e.getMessage());
            generateWord.sendText("Contraseña no ingresado - Tiempo de espera excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se da clic en el boton Acceder ingresando correctamente$")
    public void seDaClicEnElBotonAccederIngresandoCorrectamente() throws Exception {
        try {
            driver.findElement(BTN_Acceder).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se dio clic en el botón Acceder");
            generateWord.sendText("Se dio clic en el botón Acceder ");
            generateWord.addImageToWord(driver);
        }catch (Exception ex){
            ExcelReader.writeCellValue(EXCEL_WEB, CONSULTA_LINEAS_WEB, 1, 0, "FALLO" +
                    " INICIO DE SESION");
            ExtentReportUtil.INSTANCE.stepFail(driver,"Fallo el inicio de sesion" + ex.getMessage());

            generateWord.sendText("Fallo el inicio de Sesion - Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Given("^Carga la pagina Consulta mis lineas moviles$")
    public void cargaLaPaginaConsultaMisLineasMoviles() throws Exception {
        try {
            driver.findElement(LBL_CONSULTA).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Cargo la Pagina consulta de Líneas Móviles");
            generateWord.sendText("Cargo la Pagina consulta de Líneas Móviles");
            generateWord.addImageToWord(driver);
        }catch (Exception ex){
            ExcelReader.writeCellValue(EXCEL_WEB, CONSULTA_LINEAS_WEB, 1, 0, "FALLO" +
                    " la Carga la Pagina consulta de Líneas Móviles");
            ExtentReportUtil.INSTANCE.stepFail(driver,"Fallo Carga de Página" + ex.getMessage());

            generateWord.sendText("Fallo el inicio de Sesion - Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }

    }

    @When("^Selecciono el tipo de documento \"([^\"]*)\"$")
    public void seleccionoElTipoDeDocumento(String casoPrueba) throws Throwable {
        try {
            int consLineas = Integer.parseInt(casoPrueba) - 1;
            String tipoDocumento = getData().get(consLineas).get(COLUMNA_TIPO_DOCUMENTO);
            String tipoDocumentoRl = getData().get(consLineas).get(COLUMNA_TIPO_DOCUMENTORL);
            switch (tipoDocumento.toUpperCase()){
                case "DNI":
                    driver.findElement(LST_TIPO_DOC).sendKeys(Keys.DOWN,Keys.DOWN,Keys.ENTER);
                    break;
                case "RUC":
                    driver.findElement(LST_TIPO_DOC).sendKeys(Keys.DOWN,Keys.DOWN,Keys.DOWN,Keys.ENTER);
                    Sleeper.Sleep(2500);
                    switch (tipoDocumentoRl.toUpperCase()){
                        case "DNI":

                            break;
                        case "CEX":
                            driver.findElement(LST_TIPO_DOCRL).sendKeys(Keys.DOWN,Keys.DOWN,Keys.ENTER);
                            break;
                        case "PASAPORTE":
                            driver.findElement(LST_TIPO_DOCRL).sendKeys(Keys.DOWN,Keys.DOWN,Keys.DOWN,Keys.ENTER);
                            break;
                        default:
                            throw new IllegalStateException("Unexpected value: " + tipoDocumento);
                    }
                    break;

                case "CEX":
                    driver.findElement(LST_TIPO_DOC).sendKeys(Keys.DOWN,Keys.DOWN,Keys.DOWN,Keys.DOWN,Keys.ENTER);
                    break;
                case "PASAPORTE":
                    driver.findElement(LST_TIPO_DOC).sendKeys(Keys.DOWN,Keys.DOWN,Keys.DOWN,Keys.DOWN,Keys.DOWN,Keys.ENTER);
                    break;
                default:
                    throw new IllegalStateException("Unexpected value: " + tipoDocumento);
            }

            Thread.sleep(3000);
            if(tipoDocumento.toUpperCase().equals("RUC"))
            {
                ExtentReportUtil.INSTANCE.stepPass(driver, "Se selecciona el tipo de documento "
                        + tipoDocumento + " y Representante Legal: " +tipoDocumentoRl);
                generateWord.sendText("Se selecciona el tipo de documento " + tipoDocumento
                        + " y Representante Legal: " +tipoDocumentoRl);
                generateWord.addImageToWord(driver);
            }else {
                ExtentReportUtil.INSTANCE.stepPass(driver, "Se selecciona el tipo de documento " + tipoDocumento);
                generateWord.sendText("Se selecciona el tipo de documento " + tipoDocumento);
                generateWord.addImageToWord(driver);
            }
        }catch (Exception e) {
            ExcelReader.writeCellValue(EXCEL_WEB, CONSULTA_LINEAS_WEB, 1, 0, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^Se ingresa el numero documento  \"([^\"]*)\"$")
    public void seIngresaElNumeroDocumento(String casoPrueba) throws Throwable {
        try{
            int consLineas = Integer.parseInt(casoPrueba) - 1;
            String Documento = getData().get(consLineas).get(COLUMNA_DOCUMENTO);

            driver.findElement(LBL_DOCUMENTO).click();
            driver.findElement(TXT_DOCUMENTO).sendKeys(Documento);
            ExtentReportUtil.INSTANCE.stepPass(driver,"Se ingreso el numero de documento Correctamente");
            generateWord.sendText("Numero de documento ingresado correctamente");
            generateWord.addImageToWord(driver);
            String TipoDoc =  getData().get(consLineas).get(COLUMNA_TIPO_DOCUMENTO);
            if (TipoDoc.equals("RUC")){
                String DocumentoRL = getData().get(consLineas).get(COLUMNA_DOCUMENTORL);
                driver.findElement(LBL_DOCUMENTORL).click();
                Sleeper.Sleep(1000);
                driver.findElement(TXT_DOCUMENTORL).sendKeys(DocumentoRL);
                ExtentReportUtil.INSTANCE.stepPass(driver,"Se ingreso el numero de documento" +
                        "del Representante Legal Correctamente");
                generateWord.sendText("Numero de documento del representante legal ingresado correctamente");
                generateWord.addImageToWord(driver);
            }


        }catch (Exception e){
            ExcelReader.writeCellValue(EXCEL_WEB,CONSULTA_LINEAS_WEB,1,0,"FALLO " +
                    "EL INGRESO DEL DOCUMENTO");
            ExtentReportUtil.INSTANCE.stepFail(driver,"Fallo el ingreso del Documento : " + e.getMessage());
            generateWord.sendText("Docuemento no ingresado - Tiempo de espera excedido");
            generateWord.addImageToWord(driver);
        }

    }

    @And("^Se da clic en el boton Consultar$")
    public void seDaClicEnElBotonConsultar() throws Throwable {
        try {
            driver.findElement(BTN_CONSULTAR).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se dio clic en el botón Consultar");
            generateWord.sendText("Se dio clic en el botón Consultar ");
            generateWord.addImageToWord(driver);
        }catch (Exception ex){
            ExcelReader.writeCellValue(EXCEL_WEB, CONSULTA_LINEAS_WEB, 1, 0, "FALLO" +
                    " LA CONSULTA DE LINEA");
            ExtentReportUtil.INSTANCE.stepFail(driver,"Fallo la consulta" + ex.getMessage());

            generateWord.sendText("Fallo la consulta - Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }

    }

    @Then("^Se valida respuesta de Lineas del cliente$")
    public void seValidaRespuestaDeLineasDelCliente() throws Throwable {
        try {

           // driver.findElement(LBL_TABLA_RESULTADO).click();
            //driver.findElement(LBL_TABLA_RESULTADO).sendKeys(Keys.UP,Keys.UP,Keys.UP,Keys.UP);

            ExtentReportUtil.INSTANCE.stepPass(driver, "Se visualiza respuesta de la consulta de " +
                    "líneas del cliente");
            generateWord.sendText("Se visualiza respuesta de la consulta de líneas del cliente ");
            generateWord.addImageToWord(driver);
        }catch (Exception ex){
            ExcelReader.writeCellValue(EXCEL_WEB, CONSULTA_LINEAS_WEB, 1, 0, "FALLO" +
                    " LA CONSULTA DE LINEA");
            ExtentReportUtil.INSTANCE.stepFail(driver,"Fallo la consulta" + ex.getMessage());

            generateWord.sendText("Fallo la consulta - Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }



}
