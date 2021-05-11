package com.tsoft.bot.frontend.pages.pages;

import com.tsoft.bot.frontend.Base.BaseClass;
import com.tsoft.bot.frontend.helpers.Hook;
import com.tsoft.bot.frontend.pages.objects.ExcelCorporativo;
import com.tsoft.bot.frontend.pages.objects.O_Corporate;
import com.tsoft.bot.frontend.utility.ExcelReader;
import com.tsoft.bot.frontend.utility.ExtentReportUtil;
import com.tsoft.bot.frontend.utility.GenerateWord;
import cucumber.api.java.en.Given;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.util.HashMap;
import java.util.List;
import java.util.Set;

public class P_Corporativo extends BaseClass {
    String NUM_ENVIO;
    String Fecha;
    String reserva;
    String pedido;
    String flujo;
    String tipo;
    static GenerateWord generateWord = new GenerateWord();
    public WebDriver driver;
    public P_Corporativo( WebDriver driver) {
        super(driver);
        this.driver = Hook.getDriver();
    }
    private List<HashMap<String, String>> getData() throws Throwable {
        return ExcelReader.data(ExcelCorporativo.EXCEL_WEB, ExcelCorporativo.ORDEN);
    }

    public void ingresamosALAURLWEBDELIVERY(String casoDePrueba) throws Throwable {
        try {
            int LoginWD = Integer.parseInt(casoDePrueba) - 1;
            String url = getData().get(LoginWD).get(ExcelCorporativo.COLUMNA_URL);
            driver.get(url);
            stepPass(driver,"Se cargó correctamente la página");
            generateWord.sendText("Se cargó correctamente la página");
            generateWord.addImageToWord(driver);
            println("[LOG] Se cargó correctamente la página");
            generateWord.sendBreak();
        }catch (Exception e){
            ExcelReader.writeCellValue(ExcelCorporativo.EXCEL_WEB, ExcelCorporativo.ORDEN, 1, 19, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    public void ingresamosUSUARIOWEBDELIVERY(String casoDePrueba) throws Throwable {
        try {
            int user = Integer.parseInt(casoDePrueba) - 1;
            String usuario = getData().get(user).get(ExcelCorporativo.COLUMNA_USUARIO);
            wait(driver, O_Corporate.TXT_USER,60);
            if (driver.findElement(O_Corporate.TXT_USER).isDisplayed()){
                clear(driver,O_Corporate.TXT_USER);
                sendKeys(driver,O_Corporate.TXT_USER,usuario);
            }
            stepPass(driver,"Ingresamos el usuario");
            generateWord.sendText("Ingresamos el usuario");
            generateWord.addImageToWord(driver);
            println("[LOG] Ingreso correcto del usuario");
            generateWord.sendBreak();
        }catch (Exception e){
            ExcelReader.writeCellValue(ExcelCorporativo.EXCEL_WEB, ExcelCorporativo.ORDEN, 1, 19, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void ingresamosCONTRASEÑAWEBDELIVERY(String casoDePrueba) throws Throwable {
        try {
            int PASS = Integer.parseInt(casoDePrueba) - 1;
            wait(driver,O_Corporate.TXT_PASSWORD,60);
            clear(driver,O_Corporate.TXT_PASSWORD);
            String contrasenia = getData().get(PASS).get(ExcelCorporativo.COLUMNA_CONTRASENIA);
            sendKeys(driver,O_Corporate.TXT_PASSWORD,contrasenia);
            stepPass(driver,"Ingresamos contraseña");
            generateWord.sendText("Ingresamos contraseña");
            generateWord.addImageToWord(driver);
            println("[LOG] Ingreso correcto de la contraseña");
            generateWord.sendBreak();
        }catch (Exception e){
            ExcelReader.writeCellValue(ExcelCorporativo.EXCEL_WEB, ExcelCorporativo.ORDEN, 1, 19, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void clickBOTONLOGINWEBDELIVERYYSEINGRESACORRECTAMENTE() throws Throwable {
        try {
            click(driver,O_Corporate.BTN_LOGIN);
            sleep(2000);
            wait(driver,O_Corporate.LST_IR_A,60);
            stepPass(driver,"Se ingresa correctamente a la pagina");
            generateWord.sendText("Se ingresa correctamente a la pagina");
            generateWord.addImageToWord(driver);
            println("[LOG] Logueo exitoso");
            generateWord.sendBreak();
        }catch (Exception e){
            ExcelReader.writeCellValue(ExcelCorporativo.EXCEL_WEB, ExcelCorporativo.ORDEN, 1, 19, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void seleccionamosAsignaciónDeSeriesCorporativo() throws Throwable {
        click(driver,O_Corporate.LST_IR_A);
        sleep(2000);
        MoveToElement(driver,O_Corporate.LNK_GEST_PED_CORP);
        sleep(1000);
        MoveToElement(driver,O_Corporate.LNK_PREP_PEDIDO);
        click(driver,O_Corporate.LNK_ASIG_SERIES);
        wait(driver,O_Corporate.TXT_ID_RESERVA,60);
        stepPass(driver,"Asignación de series");
        generateWord.sendText("Asignación de series");
        generateWord.addImageToWord(driver);
        println("[LOG] Asignación de series");
        generateWord.sendBreak();
    }
    public void buscamosElIdDeReservaCorporativo(String casoDePrueba) throws Throwable {
        try {
            int pedido1 = Integer.parseInt(casoDePrueba) - 1;
            String user1 = getData().get(pedido1).get(ExcelCorporativo.IDRESERVA);
            clear(driver,O_Corporate.TXT_ID_RESERVA);
            sendKeys(driver,O_Corporate.TXT_ID_RESERVA,user1);
            println("[LOG] Id de reserva ingresado");
            sendKeysRobot(driver,O_Corporate.TXT_ID_RESERVA, Keys.ENTER);
            String f = driver.findElement(O_Corporate.TXT_ORDEN).getText();
            while (!f.equals(user1)) {
                sleep(1000);
                String g;
                g = driver.findElement(O_Corporate.TXT_VACIO).getText();
                if (g.equals("0 - 0 de 0")){
                    stepFail(driver,"ID Reserva no encontrado");
                    generateWord.sendText("ID Reserva no encontrado");
                    generateWord.addImageToWord(driver);
                    println("[LOG] Id de reserva no encontrado");
                    driver.quit();
                }else{
                    f = driver.findElement(O_Corporate.TXT_ORDEN).getText();
                }
            }
            stepPass(driver,"ID Reserva encontrado");
            generateWord.sendText("ID Reserva encontrado");
            generateWord.addImageToWord(driver);
            println("[LOG] Id de reserva encontrado");

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void ingresamosMaterialesIMEIYSIMCARD(String casoDePrueba) throws Throwable {
        try {
            int CantFilas = Integer.parseInt(casoDePrueba) - 1;
            String filas;
            filas = driver.findElement(O_Corporate.TABLE2).getAttribute("displayrows");
            int num = Integer.parseInt(filas);
            int s = CantFilas;
            int f = CantFilas;
            for (int  i =0; (i<num); i++){
                String TipoMaterial = driver.findElement(By.id("md6723283_tdrow_[C:6]-c[R:"+i+"]")).getText();
                if (TipoMaterial.equals("IMEI")){
                    String imei = getData().get(s).get(ExcelCorporativo.NUMERO_IMEI);
                    driver.findElement(By.id("md6723283_tdrow_[C:9]_txt-tb[R:"+i+"]")).clear();
                    driver.findElement(By.id("md6723283_tdrow_[C:9]_txt-tb[R:"+i+"]")).sendKeys(imei);
                    sleep(1000);
                    s++;
                }
                if (TipoMaterial.equals("ICCID")){
                    String sim = getData().get(f).get(ExcelCorporativo.NUMERO_SIMCARD);
                    driver.findElement(By.id("md6723283_tdrow_[C:9]_txt-tb[R:"+i+"]")).clear();
                    driver.findElement(By.id("md6723283_tdrow_[C:9]_txt-tb[R:"+i+"]")).sendKeys(sim);
                    sleep(1000);
                    f++;
                }
                if (i==num-1){
                    stepPass(driver,"Materiales ingresados");
                    generateWord.sendText("Materiales ingresados");
                    generateWord.addImageToWord(driver);
                    println("[LOG] Materiales ingresados");
                    break;
                }
            }
            sleep(2000);

/*
            int CantFilas = Integer.parseInt(casoDePrueba) - 1;
            String Cant_Filas = getData().get(CantFilas).get(CANT_FILAS);
            int Pedido = Integer.parseInt(casoDePrueba) - 1;
            String Pedidos = getData().get(Pedido).get(TIPO_PEDIDO);
            if (Cant_Filas.equals("6")&& Pedidos.equals("EQUIPO+SIM")){
                int imei1 = Integer.parseInt(casoDePrueba) - 1;
                String imei_1 = getData().get(imei1).get(NUMERO_IMEI);
                driver.findElement(TXT_MATERIAL_1).clear();
                driver.findElement(TXT_MATERIAL_1).sendKeys(imei_1);
                int sim1 = Integer.parseInt(casoDePrueba) - 1;
                String sim_1 = getData().get(sim1).get(NUMERO_SIMCARD);
                driver.findElement(TXT_MATERIAL_2).clear();
                driver.findElement(TXT_MATERIAL_2).sendKeys(sim_1);
                String imei_2 = getData().get(1).get(NUMERO_IMEI);
                driver.findElement(TXT_MATERIAL_3).clear();
                driver.findElement(TXT_MATERIAL_3).sendKeys(imei_2);
                String sim_2 = getData().get(1).get(NUMERO_SIMCARD);
                driver.findElement(TXT_MATERIAL_4).clear();
                driver.findElement(TXT_MATERIAL_4).sendKeys(sim_2);
                String imei_3 = getData().get(2).get(NUMERO_IMEI);
                driver.findElement(TXT_MATERIAL_5).clear();
                driver.findElement(TXT_MATERIAL_5).sendKeys(imei_3);
                String sim_3 = getData().get(2).get(NUMERO_SIMCARD);
                driver.findElement(TXT_MATERIAL_6).clear();
                driver.findElement(TXT_MATERIAL_6).sendKeys(sim_3);
                String imei_4 = getData().get(3).get(NUMERO_IMEI);
                driver.findElement(TXT_MATERIAL_7).clear();
                driver.findElement(TXT_MATERIAL_7).sendKeys(imei_4);
                String sim_4 = getData().get(3).get(NUMERO_SIMCARD);
                driver.findElement(TXT_MATERIAL_8).clear();
                driver.findElement(TXT_MATERIAL_8).sendKeys(sim_4);
                String imei_5 = getData().get(4).get(NUMERO_IMEI);
                driver.findElement(TXT_MATERIAL_9).clear();
                driver.findElement(TXT_MATERIAL_9).sendKeys(imei_5);
                String sim_5 = getData().get(4).get(NUMERO_SIMCARD);
                driver.findElement(TXT_MATERIAL_10).clear();
                driver.findElement(TXT_MATERIAL_10).sendKeys(sim_5);
                String imei_6 = getData().get(5).get(NUMERO_IMEI);
                driver.findElement(TXT_MATERIAL_11).clear();
                driver.findElement(TXT_MATERIAL_11).sendKeys(imei_6);
                String sim_6 = getData().get(5).get(NUMERO_SIMCARD);
                driver.findElement(TXT_MATERIAL_12).clear();
                driver.findElement(TXT_MATERIAL_12).sendKeys(sim_6);
            }
            if (Cant_Filas.equals("1")&& Pedidos.equals("EQUIPO+SIM")) {
                int imei1 = Integer.parseInt(casoDePrueba) - 1;
                String imei_1 = getData().get(imei1).get(NUMERO_IMEI);
                driver.findElement(TXT_MATERIAL_1).clear();
                driver.findElement(TXT_MATERIAL_1).sendKeys(imei_1);
                int sim1 = Integer.parseInt(casoDePrueba) - 1;
                String sim_1 = getData().get(sim1).get(NUMERO_SIMCARD);
                driver.findElement(TXT_MATERIAL_2).clear();
                driver.findElement(TXT_MATERIAL_2).sendKeys(sim_1);
            }
            if  (Cant_Filas.equals("1")&& Pedidos.equals("SOLO SIM")) {
                int sim1 = Integer.parseInt(casoDePrueba) - 1;
                String sim_1 = getData().get(sim1).get(NUMERO_SIMCARD);
                driver.findElement(TXT_MATERIAL_1).clear();
                driver.findElement(TXT_MATERIAL_1).sendKeys(sim_1);
            }
*/
        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void validamosSeriesCorporativo() throws Throwable {
        try {
            click(driver,O_Corporate.BTN_VALIDAR_SERIE);
            println("[LOG] Validamos materiales");
            wait(driver,O_Corporate.BTN_SI,60);
            stepPass(driver,"Mensaje: Validar Serie");
            generateWord.sendText("Mensaje: Validar Serie");
            generateWord.addImageToWord(driver);
            click(driver,O_Corporate.BTN_SI);
            wait(driver,O_Corporate.BTN_ACEPTAR_MENS_SIST,60);
            stepPass(driver,"Mensaje de Validación");
            generateWord.sendText("Mensaje de Validación");
            generateWord.addImageToWord(driver);
            click(driver,O_Corporate.BTN_ACEPTAR_MENS_SIST);
            sleep(2000);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void verificarElEstadoDeAsignaciónDeSerieCorporativos(String casoDePrueba) throws Throwable {
        try {
            String filas;
            int error = 0;
            int bien = 0;
            filas = driver.findElement(O_Corporate.TABLE2).getAttribute("displayrows");
            int num = Integer.parseInt(filas);
            for(int  i =0; (i<num); i++){
                String estado = driver.findElement(By.id("md6723283_tdrow_[C:10]-c[R:"+i+"]")).getText();
                if (estado.equals("ERROR")){
                    error++;
                    String valor = driver.findElement(By.id("md6723283_tdrow_[C:9]_txt-tb[R:"+i+"]")).getAttribute("value");
                    println(valor +" --> " + estado);
                }
                if (estado.equals("VALIDADO")){
                    bien++;
                    String valor = driver.findElement(By.id("md6723283_tdrow_[C:9]_txt-tb[R:"+i+"]")).getAttribute("value");
                    println(valor +" --> " + estado);
                }
                if (i == num-1){
                    if (error>0){
                        println("Se obtuvo un total de: " +error+ " materiales con error");
                        println("Se obtuvo un total de: " +bien+ " materiales validados");
                        stepFail(driver,"Materiales no validados");
                        generateWord.sendText("Materiales no validados");
                        println("[LOG] Materiales no validados");
                        generateWord.addImageToWord(driver);
                        driver.quit();
                        break;
                    }else{
                        println("Se obtuvo un total de: " +bien+ " materiales validados");
                        println("[LOG] Materiales validados");
                        ExtentReportUtil.INSTANCE.stepFail(driver, "Materiales validados");
                        generateWord.sendText("Materiales validados");
                        generateWord.addImageToWord(driver);
                        break;
                    }
                }
            }
            /*
            int CantFilas = Integer.parseInt(casoDePrueba) - 1;
            String Cant_Filas = getData().get(CantFilas).get(CANT_FILAS);
            int Pedido = Integer.parseInt(casoDePrueba) - 1;
            String Pedidos = getData().get(Pedido).get(TIPO_PEDIDO);
            if (Cant_Filas.equals("6")&& Pedidos.equals("EQUIPO+SIM")){
                String Estado1 = driver.findElement(LBL_ESTADO_VAL_SERIE1).getText();
                String Estado2 = driver.findElement(LBL_ESTADO_VAL_SERIE2).getText();
                String Estado3 = driver.findElement(LBL_ESTADO_VAL_SERIE3).getText();
                String Estado4 = driver.findElement(LBL_ESTADO_VAL_SERIE4).getText();
                String Estado5 = driver.findElement(LBL_ESTADO_VAL_SERIE5).getText();
                String Estado6 = driver.findElement(LBL_ESTADO_VAL_SERIE6).getText();
                String Estado7 = driver.findElement(LBL_ESTADO_VAL_SERIE7).getText();
                String Estado8 = driver.findElement(LBL_ESTADO_VAL_SERIE8).getText();
                String Estado9 = driver.findElement(LBL_ESTADO_VAL_SERIE9).getText();
                String Estado10 = driver.findElement(LBL_ESTADO_VAL_SERIE10).getText();
                String Estado11 = driver.findElement(LBL_ESTADO_VAL_SERIE11).getText();
                String Estado12 = driver.findElement(LBL_ESTADO_VAL_SERIE12).getText();
                if (Estado1.equals("VALIDADO") && Estado2.equals("VALIDADO") && Estado3.equals("VALIDADO") && Estado4.equals("VALIDADO")&& Estado5.equals("VALIDADO") && Estado6.equals("VALIDADO") && Estado7.equals("VALIDADO") && Estado8.equals("VALIDADO") && Estado9.equals("VALIDADO") && Estado10.equals("VALIDADO") && Estado11.equals("VALIDADO") && Estado12.equals("VALIDADO")){
                    ExtentReportUtil.INSTANCE.stepPass(driver, "Estado de Validación de series: VALIDADO");
                    generateWord.sendText("Estado de Validación de series: VALIDADO");
                    generateWord.addImageToWord(driver);
                }else {
                    ExtentReportUtil.INSTANCE.stepFail(driver, "IMEI Y SIMCARD no validados");
                    generateWord.sendText("IMEI Y SIMCARD no validados");
                    generateWord.addImageToWord(driver);
                }
            }
            if (Cant_Filas.equals("1")&& Pedidos.equals("EQUIPO+SIM")){
                String Estado1 = driver.findElement(LBL_ESTADO_VAL_SERIE1).getText();
                String Estado2 = driver.findElement(LBL_ESTADO_VAL_SERIE2).getText();
                if (Estado1.equals("VALIDADO") && Estado2.equals("VALIDADO")){
                    ExtentReportUtil.INSTANCE.stepPass(driver, "Estado de Validación de series: VALIDADO");
                    generateWord.sendText("Estado de Validación de series: VALIDADO");
                    generateWord.addImageToWord(driver);
                }else {

                        ExtentReportUtil.INSTANCE.stepFail(driver, "IMEI Y SIMCARD no validados");
                        generateWord.sendText("IMEI Y SIMCARD no validados");
                        generateWord.addImageToWord(driver);
                        driver.quit();
                }
            }
            if (Cant_Filas.equals("1")&& Pedidos.equals("SOLO SIM")){
                String Estado1 = driver.findElement(LBL_ESTADO_VAL_SERIE1).getText();
                if (Estado1.equals("VALIDADO")){
                    ExtentReportUtil.INSTANCE.stepPass(driver, "Estado de Validación de series: VALIDADO" );
                    generateWord.sendText("Estado de Validación de series: VALIDADO");
                    generateWord.addImageToWord(driver);
                }else {
                    ExtentReportUtil.INSTANCE.stepFail(driver, "Materiales no validados");
                    generateWord.sendText("Materiales no validados");
                    generateWord.addImageToWord(driver);
                }
            }*/

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void seleccionamosImpresionDeDocumentosCorporativo()  throws Throwable {
        click(driver,O_Corporate.LST_IR_A);
        MoveToElement(driver,O_Corporate.LNK_GEST_PED_CORP);
        MoveToElement(driver,O_Corporate.LNK_PREP_PEDIDO);
        click(driver, O_Corporate.LNK_IMPR_DOC);
        wait(driver,O_Corporate.TXT_ID_RESERVA2,60);
        ExtentReportUtil.INSTANCE.stepPass(driver, "Impresión de documentos");
        generateWord.sendText("Impresión de documentos");
        generateWord.addImageToWord(driver);
    }
    public void buscamosElIdDeReservaCorporativo_2(String casoDePrueba) throws Throwable {
        try {
            int pedido1 = Integer.parseInt(casoDePrueba) - 1;
            String user1 = getData().get(pedido1).get(ExcelCorporativo.IDRESERVA);
            wait(driver, O_Corporate.TXT_ID_RESERVA, 60);
            clear(driver, O_Corporate.TXT_ID_RESERVA);
            sendKeys(driver, O_Corporate.TXT_ID_RESERVA, user1);
            sendKeysRobot(driver, O_Corporate.TXT_ID_RESERVA, Keys.ENTER);
            String f = driver.findElement(O_Corporate.TXT_ORDEN).getText();
            while (!f.equals(user1)) {
                Thread.sleep(1000);
                String g;
                g = driver.findElement(O_Corporate.TXT_VACIO).getText();
                if (g.equals("0 - 0 de 0")) {
                    ExtentReportUtil.INSTANCE.stepPass(driver, "ID Reserva no encontrado");
                    generateWord.sendText("ID Reserva no encontrado");
                    generateWord.addImageToWord(driver);
                    driver.quit();
                } else {
                    f = driver.findElement(O_Corporate.TXT_ORDEN).getText();
                }
            }
            ExtentReportUtil.INSTANCE.stepPass(driver, "ID Reserva encontrado");
            generateWord.sendText("ID Reserva encontrado");
            generateWord.addImageToWord(driver);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void contratoDigitalCorporativo() throws Exception {
        try {
            click(driver,O_Corporate.BTN_CARGA_CONTRATO_DIGITAL);
            wait(driver,O_Corporate.BTN_SI_CARGA_CONTRATO_DIGITAL,60);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Contrato Digital");
            generateWord.sendText("Contrato Digital");
            generateWord.addImageToWord(driver);
            click(driver,O_Corporate.BTN_SI_CARGA_CONTRATO_DIGITAL);
            wait(driver,O_Corporate.BTN_ACEPTAR_MENS_SIST,60);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje del sistema");
            generateWord.sendText("Mensaje del sistema");
            generateWord.addImageToWord(driver);
            click(driver,O_Corporate.BTN_ACEPTAR_MENS_SIST);
            Thread.sleep(2000);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }

    }
    public void preparaciónDeLaFacturaCorporativo() throws Exception {
        try {
            click(driver,O_Corporate.CMB_SELECT_ACTION);
            Thread.sleep(1000);
            click(driver,O_Corporate.CMB_SELECT_ACTION_PREP_FACTURA);
            wait(driver,O_Corporate.BTN_ACEPTAR_MENS_SIST,60);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje del sistema");
            generateWord.sendText("Mensaje del sistema");
            generateWord.addImageToWord(driver);
            click(driver,O_Corporate.BTN_ACEPTAR_MENS_SIST);
            wait(driver,O_Corporate.BTN_SI_2,60);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Preparación de la factura");
            generateWord.sendText("Preparación de la factura");
            generateWord.addImageToWord(driver);
            click(driver,O_Corporate.BTN_SI_2);
            wait(driver,O_Corporate.BTN_ACEPTAR_MENS_SIST,60);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje del sistema");
            generateWord.sendText("Mensaje del sistema");
            generateWord.addImageToWord(driver);
            click(driver,O_Corporate.BTN_ACEPTAR_MENS_SIST);
            Thread.sleep(2000);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void impresiónDeLaFacturaCorporativo() throws Exception {
        try {
            click(driver,O_Corporate.CMB_SELECT_ACTION);
            Thread.sleep(1000);
            click(driver,O_Corporate.CMB_SELECT_ACTION_IMPR_FACTURA);
            wait(driver,O_Corporate.BTN_SI_3,60);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Impresión de la factura");
            generateWord.sendText("Impresión de la factura");
            generateWord.addImageToWord(driver);
            click(driver,O_Corporate.BTN_SI_3);
            Thread.sleep(5000);
            String parentWindow = driver.getWindowHandle();
            Set<String> handles = driver.getWindowHandles();
            for (String windowHandle : handles) {
                if (!windowHandle.equals(parentWindow)) {
                    ExtentReportUtil.INSTANCE.stepPass(driver.switchTo().window(windowHandle), "Reporte de impresion de guia de remision");
                    generateWord.sendText("Reporte de impresion de guia de remision");
                    generateWord.addImageToWord(driver.switchTo().window(windowHandle));
                    driver.switchTo().window(windowHandle).close();
                    driver.switchTo().window(parentWindow);
                    break;
                }
            }
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void ejecutarInformesCorporativo() throws Exception {
        try {
            click(driver,O_Corporate.CMB_SELECT_ACTION);
            Thread.sleep(1000);
            click(driver,O_Corporate.CMB_SELECT_ACTION_EJEC_INFORMES);
            wait(driver,O_Corporate.LNK_IMPR_GUIA_REMISION,60);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Ejecutar informes");
            generateWord.sendText("Ejecutar informes");
            generateWord.addImageToWord(driver);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void imprimirGuiaDeRemisiónCorporativo() throws Exception {
        try {
            click(driver,O_Corporate.LNK_IMPR_GUIA_REMISION);
            wait(driver,O_Corporate.TXT_SERIE_GUIA_REMISION,60);
            clear(driver,O_Corporate.TXT_SERIE_GUIA_REMISION);
            sendKeys(driver,O_Corporate.TXT_SERIE_GUIA_REMISION,"12345");
            Thread.sleep(100);
            clear(driver,O_Corporate.TXT_CORRELATIVO_GUIA);
            sendKeys(driver,O_Corporate.TXT_CORRELATIVO_GUIA,"12345");
            Thread.sleep(100);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Imprimir guia de remisión");
            generateWord.sendText("Imprimir guia de remisión");
            generateWord.addImageToWord(driver);
            click(driver,O_Corporate.BTN_ACEPTAR);
            Thread.sleep(3000);
            Boolean isPresent = driver.findElements(O_Corporate.TXT_IMAGEN).size() > 0;
            if (isPresent.equals(true)){
                ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje error");
                generateWord.sendText("Mensaje error");
                generateWord.addImageToWord(driver);
                driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
                driver.findElement(O_Corporate.TXT_SERIE_GUIA_REMISION).clear();
                driver.findElement(O_Corporate.TXT_SERIE_GUIA_REMISION).sendKeys("12345");
                Thread.sleep(100);
                driver.findElement(O_Corporate.TXT_CORRELATIVO_GUIA).clear();
                driver.findElement(O_Corporate.TXT_CORRELATIVO_GUIA).sendKeys("12345");
                Thread.sleep(100);
                ExtentReportUtil.INSTANCE.stepPass(driver, "Imprimir guia de remisión");
                generateWord.sendText("Imprimir guia de remisión");
                generateWord.addImageToWord(driver);
                driver.findElement(O_Corporate.BTN_ACEPTAR).click();
                Thread.sleep(3000);
                Boolean isPresent2 = driver.findElements(O_Corporate.TXT_IMAGEN).size() > 0;
                if (isPresent2.equals(true)){
                    driver.quit();
                }
            }
            String parentWindow = driver.getWindowHandle();
            Set<String> handles = driver.getWindowHandles();
            for (String windowHandle : handles) {
                if (!windowHandle.equals(parentWindow)) {
                    wait(driver,O_Corporate.TXT_NUMRUC,60);
                    ExtentReportUtil.INSTANCE.stepPass(driver.switchTo().window(windowHandle), "Reporte de impresion de guia de remision");
                    generateWord.sendText("Reporte de impresion de guia de remision");
                    generateWord.addImageToWord(driver.switchTo().window(windowHandle));
                    driver.switchTo().window(windowHandle).close();
                    driver.switchTo().window(parentWindow);
                    break;
                }
            }
            ExtentReportUtil.INSTANCE.stepPass(driver, "Informes y programaciones");
            generateWord.sendText("Informes y programaciones");
            generateWord.addImageToWord(driver);

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void imprimirEtiquetaDeLineaCorporativo() throws Exception {
        try {
            driver.findElement(O_Corporate.LNK_IMPR_ETIQUETA_LINEA).click();
            wait(driver,O_Corporate.RBTN_INMEDIATO,60);
            driver.findElement(O_Corporate.RBTN_INMEDIATO).click();
            wait(driver,O_Corporate.BTN_ENVIAR2,60);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Imprimir etiqueta de linea");
            generateWord.sendText("Imprimir etiqueta de linea");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ENVIAR2).click();
            Thread.sleep(1000);
            String parentWindow = driver.getWindowHandle();
            Set<String> handles = driver.getWindowHandles();
            for (String windowHandle : handles) {
                if (!windowHandle.equals(parentWindow)) {
                    wait(driver,O_Corporate.TXT_SERVICIO,60);
                    ExtentReportUtil.INSTANCE.stepPass(driver.switchTo().window(windowHandle), "Reporte de etiqueta en linea");
                    generateWord.sendText("Reporte de etiqueta en linea");
                    generateWord.addImageToWord(driver.switchTo().window(windowHandle));
                    driver.switchTo().window(windowHandle).close();
                    driver.switchTo().window(parentWindow);
                    break;
                }
            }
            ExtentReportUtil.INSTANCE.stepPass(driver, "Informes y programaciones");
            generateWord.sendText("Informes y programaciones");
            generateWord.addImageToWord(driver);

        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);

        }
    }
    public void imprimirEtiquetaCorporativo() throws Exception {
        try {
            driver.findElement(O_Corporate.LNK_IMPR_ETIQUETA).click();
            wait(driver,O_Corporate.RBTN_INMEDIATO2,60);
            driver.findElement(O_Corporate.RBTN_INMEDIATO2).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Imprimir etiqueta");
            generateWord.sendText("Imprimir etiqueta");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ENVIAR3).click();
            Thread.sleep(1000);
            String parentWindow = driver.getWindowHandle();
            Set<String> handles = driver.getWindowHandles();
            for (String windowHandle : handles) {
                if (!windowHandle.equals(parentWindow)) {
                    wait(driver,O_Corporate.TXT_NUMRUC2,60);
                    ExtentReportUtil.INSTANCE.stepPass(driver.switchTo().window(windowHandle), "Reporte de etiqueta en linea");
                    generateWord.sendText("Reporte de etiqueta en linea");
                    generateWord.addImageToWord(driver.switchTo().window(windowHandle));
                    driver.switchTo().window(windowHandle).close();
                    driver.switchTo().window(parentWindow);
                    break;
                }
            }
            wait(driver,O_Corporate.BTN_CANCELAR,60);
            driver.findElement(O_Corporate.BTN_CANCELAR).click();
            Thread.sleep(2000);
        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);

        }
    }
    public void observamosElCambioDeEstado() throws Exception {
        wait(driver,O_Corporate.TXT_RUTA_FACTURA,60);
        driver.findElement(O_Corporate.TXT_RUTA_FACTURA).click();
        ExtentReportUtil.INSTANCE.stepPass(driver, "Impresión de documentos terminado");
        generateWord.sendText("Impresión de documentos terminado");
        generateWord.addImageToWord(driver);
        Thread.sleep(2000);
    }
    public void seleccionamosDespachoDeHubCorporativo() throws Exception {
        Thread.sleep(2000);
        driver.findElement(O_Corporate.LST_IR_A).click();
        MoveToElement(driver,O_Corporate.LNK_GEST_PED_CORP);
        driver.findElement(O_Corporate.LNK_DESPACHO_HUB).click();
        wait(driver,O_Corporate.TXT_ID_RESERVA2,60);
        ExtentReportUtil.INSTANCE.stepPass(driver, "Despacho al HUB");
        generateWord.sendText("Despacho al HUB");
        generateWord.addImageToWord(driver);
    }
    public void generaciónDeMasterBOXCorporativo() throws Exception {
        try {
            driver.findElement(O_Corporate.LNK_GENERACION_MASTERBOX).click();
            wait(driver,O_Corporate.TXT_MONTO_CAMION,60);
            driver.findElement(O_Corporate.TXT_MONTO_CAMION).clear();
            driver.findElement(O_Corporate.TXT_MONTO_CAMION).sendKeys("354445");
            Thread.sleep(1000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Generación de Master box");
            generateWord.sendText("Generación de Master box");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ENVIAR4).click();
            Thread.sleep(2000);
            String parentWindow = driver.getWindowHandle();
            Set<String> handles = driver.getWindowHandles();
            for (String windowHandle : handles) {
                if (!windowHandle.equals(parentWindow)) {
                    WebDriverWait wait2 = new WebDriverWait(driver.switchTo().window(windowHandle), 60);
                    wait2.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_MASTER));
                    ExtentReportUtil.INSTANCE.stepPass(driver.switchTo().window(windowHandle), "Informe generación de master box");
                    generateWord.sendText("Informe generación de master box");
                    generateWord.addImageToWord(driver.switchTo().window(windowHandle));
                    driver.switchTo().window(windowHandle).close();
                    driver.switchTo().window(parentWindow);
                    break;
                }
            }
            wait(driver,O_Corporate.BTN_CANCELAR2,60);
            driver.findElement(O_Corporate.BTN_CANCELAR2).click();
            Thread.sleep(2000);

        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);

        }
    }
    public void despacharPedidoCorporativo() throws Exception {
        try {

            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.LBL_COD_MASTERBOX));
            String Cod_Master_Box = driver.findElement(O_Corporate.LBL_COD_MASTERBOX).getText();
            Fecha = driver.findElement(O_Corporate.LBL_FECHA_DESPACHO).getText();
            Thread.sleep(1000);
            driver.findElement(O_Corporate.BTN_DESPACHAR_PEDIDO).click();
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_MASTERBOX));
            driver.findElement(O_Corporate.TXT_MASTERBOX).clear();
            driver.findElement(O_Corporate.TXT_MASTERBOX).sendKeys(Cod_Master_Box);
            /*SimpleDateFormat df = new SimpleDateFormat("dd/MM/YYYY");
            Date dt = new Date();
            Calendar cl = Calendar.getInstance();
            cl.setTime(dt);;
            cl.add(Calendar.DAY_OF_MONTH, 0);
            dt=cl.getTime();
            String str = df.format(dt);*/
            Thread.sleep(1000);
            driver.findElement(O_Corporate.TXT_FECHA_PEDIDO).clear();
            driver.findElement(O_Corporate.TXT_FECHA_PEDIDO).sendKeys(Fecha);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Despachar Pedido");
            generateWord.sendText("Despachar Pedido");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_SI4).click();
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ACEPTAR_MENS_SIST));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje de sistema");
            generateWord.sendText("Mensaje de sistema");
            generateWord.addImageToWord(driver);
            String txt = driver.findElement(By.id("mb_msg")).getText();
            if (txt.contains("La fecha de despacho es requerida. ")){
                driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
                Thread.sleep(1000);
                driver.findElement(O_Corporate.TXT_FECHA_PEDIDO).clear();
                driver.findElement(O_Corporate.TXT_FECHA_PEDIDO).sendKeys(Fecha);
                ExtentReportUtil.INSTANCE.stepPass(driver, "Despachar Pedido");
                generateWord.sendText("Despachar Pedido");
                generateWord.addImageToWord(driver);
                driver.findElement(O_Corporate.BTN_SI4).click();
                wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ACEPTAR_MENS_SIST));
                if (txt.contains("La fecha de despacho es requerida. ")){
                    ExtentReportUtil.INSTANCE.stepFail(driver, "Despachar Pedido Fallido");
                    generateWord.sendText("Despachar Pedido Fallido");
                    generateWord.addImageToWord(driver);
                    driver.quit();
                }

            }
            driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
            Thread.sleep(1000);

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void seleccionamosRecepciónDePedidosCorporativo() throws Exception {

        Thread.sleep(2000);
        driver.findElement(O_Corporate.LST_IR_A).click();
        MoveToElement(driver,O_Corporate.LNK_GEST_PED_CORP);
        driver.findElement(O_Corporate.LNK_RECEPCION_PEDIDO).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_ID_RESERVA2));
        ExtentReportUtil.INSTANCE.stepPass(driver, "Recepción de pedido");
        generateWord.sendText("Recepción de pedido");
        generateWord.addImageToWord(driver);
    }
    public void recepcionarDePedidosCorporativo() throws Exception {
        try {

            String Cod_Master_Box2 = driver.findElement(O_Corporate.LBL_COD_MASTERBOX2).getText();
            Fecha = driver.findElement(O_Corporate.LBL_FECHA_DESPACHO).getText();
            Thread.sleep(1000);
            driver.findElement(O_Corporate.BTN_RECEP_PEDIDO).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_MASTERBOX2));
            driver.findElement(O_Corporate.TXT_MASTERBOX2).sendKeys(Cod_Master_Box2);
            /*SimpleDateFormat df = new SimpleDateFormat("dd/MM/YYYY");
            Date dt = new Date();
            Calendar cl = Calendar.getInstance();
            cl.setTime(dt);;
            cl.add(Calendar.DAY_OF_MONTH, 0);
            dt=cl.getTime();
            String str = df.format(dt);*/
            driver.findElement(O_Corporate.TXT_FECHA_PEDIDO2).sendKeys(Fecha);
            Thread.sleep(1000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Recepcionar pedido");
            generateWord.sendText("Recepcionar pedido");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_SI5).click();
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ACEPTAR_MENS_SIST));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje del sistema");
            generateWord.sendText("Mensaje del sistema");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
            Thread.sleep(2000);

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void procesoDeLógicaDeRuteoCorporativo() throws Exception {
        try {

            driver.findElement(O_Corporate.LNK_PROC_LOGICARUTEO).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_FECHA_DESPACHO));
            /*SimpleDateFormat df = new SimpleDateFormat("dd/MM/YYYY");
            Date dt = new Date();
            Calendar cl = Calendar.getInstance();
            cl.setTime(dt);;
            cl.add(Calendar.DAY_OF_MONTH, 0);
            dt=cl.getTime();
            String str = df.format(dt);*/
            driver.findElement(O_Corporate.TXT_FECHA_DESPACHO).clear();
            driver.findElement(O_Corporate.TXT_FECHA_DESPACHO).sendKeys(Fecha);
            Thread.sleep(1000);
            driver.findElement(O_Corporate.TXT_ALMACEN).clear();
            driver.findElement(O_Corporate.TXT_ALMACEN).sendKeys("=PE10API7");
            Thread.sleep(1000);
            driver.findElement(O_Corporate.TXT_TURNO).clear();
            driver.findElement(O_Corporate.TXT_TURNO).sendKeys("=TARDE_API7");
            Thread.sleep(1000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Proceso de logica de ruteo");
            generateWord.sendText("Proceso de logica de ruteo");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ENVIAR5).click();
            Thread.sleep(1000);
            Boolean isPresent = driver.findElements(O_Corporate.TXT_IMAGEN).size() > 0;
            if (isPresent.equals(true)){
                ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje error");
                generateWord.sendText("Mensaje error");
                generateWord.addImageToWord(driver);
                String f = driver.findElement(O_Corporate.TXT_IMAGEN).getText();
                if (f.contains("Fecha de despacho es un campo necesario") || f.contains("Almacén es un campo necesario") || f.contains("Turno es un campo necesario")){
                    driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
                    driver.findElement(O_Corporate.TXT_FECHA_DESPACHO).clear();
                    driver.findElement(O_Corporate.TXT_FECHA_DESPACHO).sendKeys(Fecha);
                    Thread.sleep(1000);
                    driver.findElement(O_Corporate.TXT_ALMACEN).clear();
                    driver.findElement(O_Corporate.TXT_ALMACEN).sendKeys("=PE10API7");
                    Thread.sleep(1000);
                    driver.findElement(O_Corporate.TXT_TURNO).clear();
                    driver.findElement(O_Corporate.TXT_TURNO).sendKeys("=TARDE_API7");
                    Thread.sleep(1000);
                    ExtentReportUtil.INSTANCE.stepPass(driver, "Proceso de logica de ruteo");
                    generateWord.sendText("Proceso de logica de ruteo");
                    generateWord.addImageToWord(driver);
                    driver.findElement(O_Corporate.BTN_ENVIAR5).click();
                    Boolean isPresent2 = driver.findElements(O_Corporate.TXT_IMAGEN).size() > 0;
                    if (isPresent2.equals(true)){
                        driver.quit();
                    }
                }

            }

            String parentWindow = driver.getWindowHandle();
            Set<String> handles = driver.getWindowHandles();
            for (String windowHandle : handles) {
                if (!windowHandle.equals(parentWindow)) {
                    WebDriverWait wait2 = new WebDriverWait(driver.switchTo().window(windowHandle), 60);
                    wait2.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_FECHA));
                    ExtentReportUtil.INSTANCE.stepPass(driver.switchTo().window(windowHandle), "Proceso de lógica de ruteo");
                    generateWord.sendText("Proceso de lógica de ruteo");
                    generateWord.addImageToWord(driver.switchTo().window(windowHandle));
                    driver.switchTo().window(windowHandle).close();
                    driver.switchTo().window(parentWindow);
                    break;
                }
            }
            ExtentReportUtil.INSTANCE.stepPass(driver, "Informes y programaciones");
            generateWord.sendText("Informes y programaciones");
            generateWord.addImageToWord(driver);
            Thread.sleep(2000);


        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);

        }
    }
    public void reporteDeLogicaDeRuteoCorporativo() throws Exception {
        try {

            driver.findElement(O_Corporate.LNK_REPORTE_LOGICARUTEO).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_FECHA_DESPACHO2));
            /*SimpleDateFormat df = new SimpleDateFormat("dd/MM/YYYY");
            Date dt = new Date();
            Calendar cl = Calendar.getInstance();
            cl.setTime(dt);;
            cl.add(Calendar.DAY_OF_MONTH, 0);
            dt=cl.getTime();
            String str = df.format(dt);*/
            driver.findElement(O_Corporate.TXT_FECHA_DESPACHO2).clear();
            driver.findElement(O_Corporate.TXT_FECHA_DESPACHO2).sendKeys(Fecha);
            Thread.sleep(1000);
            driver.findElement(O_Corporate.TXT_ALMACEN2).clear();
            driver.findElement(O_Corporate.TXT_ALMACEN2).sendKeys("=PE10API7");
            Thread.sleep(1000);
            driver.findElement(O_Corporate.TXT_TURNO2).clear();
            driver.findElement(O_Corporate.TXT_TURNO2).sendKeys("=TARDE_API7");
            Thread.sleep(1000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Reporte de logica de ruteo");
            generateWord.sendText("Reporte de logica de ruteo");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ENVIAR6).click();
            Thread.sleep(10000);
            String parentWindow = driver.getWindowHandle();
            Set<String> handles = driver.getWindowHandles();
            for (String windowHandle : handles) {
                if (!windowHandle.equals(parentWindow)) {

                    ExtentReportUtil.INSTANCE.stepPass(driver.switchTo().window(windowHandle), "Proceso de lógica de ruteo");
                    generateWord.sendText("Proceso de lógica de ruteo");
                    generateWord.addImageToWord(driver.switchTo().window(windowHandle));
                    driver.switchTo().window(windowHandle).close();
                    driver.switchTo().window(parentWindow);
                    break;
                }
            }
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_CANCELAR3));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Informes y programaciones");
            generateWord.sendText("Informes y programaciones");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_CANCELAR3).click();
            Thread.sleep(2000);

        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);

        }
    }
    public void seleccionamosRecepciónDePedidosCorporativos() throws Exception {

        Thread.sleep(2000);
        driver.findElement(O_Corporate.LST_IR_A).click();
        MoveToElement(driver,O_Corporate.LNK_GEST_PED_CORP);
        driver.findElement(O_Corporate.LNK_RECEPCION_PEDIDO).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_ID_RESERVA2));
        ExtentReportUtil.INSTANCE.stepPass(driver, "Recepción de pedido");
        generateWord.sendText("Recepción de pedido");
        generateWord.addImageToWord(driver);
    }
    public void ejecutarCargaDeRuteoCorporativo() throws Exception {
        Thread.sleep(2000);
        driver.findElement(O_Corporate.BTN_EJECUTAR_CARGA_RUTEO).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_SI6));
        ExtentReportUtil.INSTANCE.stepPass(driver, "Carga de logica de ruteo");
        generateWord.sendText("Carga de logica de ruteo");
        generateWord.addImageToWord(driver);
        driver.findElement(O_Corporate.BTN_SI6).click();
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ACEPTAR_MENS_SIST));
        /*String f = driver.findElement(By.id("mb_msg")).getText();
        if(f.contains("No se han cargado Datos para procesar, verificar el archivo adjunto y cargarlo nuevamente")){
            ExtentReportUtil.INSTANCE.stepPass(driver, "Error al momento de cargar el archivo");
            generateWord.sendText("Error al momento de cargar el archivo");
            generateWord.addImageToWord(driver);
            System.out.println("Error al momento de cargar el archivo");
            driver.quit();
        }*/
        ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje del sistema");
        generateWord.sendText("Mensaje del sistema");
        generateWord.addImageToWord(driver);
        driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
        Thread.sleep(2000);
    }
    public void guardamosNúmeroDeEnvioCorporativo() throws Exception {
        Thread.sleep(2000);
        driver.findElement(O_Corporate.TXT_ORDEN_ENVIO).click();
        Thread.sleep(1000);
        NUM_ENVIO = driver.findElement(O_Corporate.TXT_NUMERO_ENVIO).getText();
        System.out.println(NUM_ENVIO);
        if (NUM_ENVIO.equals("")){
            Thread.sleep(1000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Número de envío no generado");
            generateWord.sendText("Número de envío no generado");
            generateWord.addImageToWord(driver);
            driver.quit();
        }
        Thread.sleep(1000);
        ExtentReportUtil.INSTANCE.stepPass(driver, "Número de envío generado");
        generateWord.sendText("Número de envío generado");
        generateWord.addImageToWord(driver);
    }
    public void despachoAMotorizadoCorporativo() throws Exception {

        driver.findElement(O_Corporate.LST_IR_A).click();
        MoveToElement(driver,O_Corporate.LNK_GEST_PED_CORP);
        driver.findElement(O_Corporate.LNK_DESPACHO_MOTORIZADO).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_ENVIO));
        ExtentReportUtil.INSTANCE.stepPass(driver, "Despacho a motorizado");
        generateWord.sendText("Despacho a motorizado");
        generateWord.addImageToWord(driver);
    }
    public void buscamosNumeroDeEnvioCorporativo() throws Exception {
        Thread.sleep(2000);
        driver.findElement(O_Corporate.TXT_ENVIO).clear();
        driver.findElement(O_Corporate.TXT_ENVIO).sendKeys(NUM_ENVIO);
        driver.findElement(O_Corporate.TXT_ENVIO).sendKeys(Keys.ENTER);
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.LNK_NUM_ENVIO));
        ExtentReportUtil.INSTANCE.stepPass(driver, "Pedido encontrado");
        generateWord.sendText("Pedido encontrado");
        generateWord.addImageToWord(driver);
        driver.findElement(O_Corporate.LNK_NUM_ENVIO).click();
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_LUPA11));
    }
    public void buscamosMotorizadoCorporativo() throws Exception {

        driver.findElement(O_Corporate.BTN_LUPA11).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_MOTORIZADO));
        driver.findElement(O_Corporate.TXT_MOTORIZADO).sendKeys("VMONTOVELI");
        driver.findElement(O_Corporate.TXT_MOTORIZADO).sendKeys(Keys.ENTER);
        Thread.sleep(3000);
        ExtentReportUtil.INSTANCE.stepPass(driver, "Seleccionamos motorizado");
        generateWord.sendText("Seleccionamos motorizado");
        generateWord.addImageToWord(driver);
        driver.findElement(O_Corporate.LNK_MOTORIZADO).click();
        Thread.sleep(3000);

    }
    public void despachamosPedidoCorporativo() throws Throwable {
        Thread.sleep(2000);

        driver.findElement(O_Corporate.BTN_PEDIDOS).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_CAMBIAR_ESTADO));
        String filas;
        String d;
        filas = driver.findElement(O_Corporate.TABLE).getAttribute("displayrows");
        int num = Integer.parseInt(filas);
        for (int  i =0; (i<num); i++){
            String orden = driver.findElement(By.id("m187f4d3c_tdrow_[C:4]-c[R:"+i+"]")).getText();
            System.out.println(tipo);
            if (tipo.equals("SI")){
                System.out.println(orden);
                if (orden.equals(pedido)){
                    driver.findElement(By.id("m187f4d3c_tdrow_[C:0]_checkbox-cb[R:"+i+"]_img")).click();
                    break;
                }
            }
            if (tipo.equals("NO")){
                if (orden.equals(reserva)){
                    driver.findElement(By.id("m187f4d3c_tdrow_[C:0]_checkbox-cb[R:"+i+"]_img")).click();
                    break;
                }
            }

            if (i==num-1){
                ExtentReportUtil.INSTANCE.stepFail(driver, "No se encontro la orden");
                generateWord.sendText("No se encontro la orden");
                generateWord.addImageToWord(driver);
                driver.quit();
                break;
            }
        }

        //driver.findElement(BTN_CAMBIAR_ESTADO).click();
        Thread.sleep(1000);
        ExtentReportUtil.INSTANCE.stepPass(driver, "Pedido listo");
        generateWord.sendText("Pedido listo");
        generateWord.addImageToWord(driver);
        driver.findElement(O_Corporate.BTN_DESPACHAR_PEDIDO2).click();
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ACEPTAR_DESPACHAR_PEDIDO));
        ExtentReportUtil.INSTANCE.stepPass(driver, "Despachar pedido");
        generateWord.sendText("Despachar pedido");
        generateWord.addImageToWord(driver);
        driver.findElement(O_Corporate.BTN_ACEPTAR_DESPACHAR_PEDIDO).click();
        Thread.sleep(8000);
        ExtentReportUtil.INSTANCE.stepPass(driver, "Despacho motorizado LISTO");
        generateWord.sendText("Despacho motorizado LISTO");
        generateWord.addImageToWord(driver);

    }
    public void maestroDePedidosCorporativo() throws Exception {

        driver.findElement(O_Corporate.LST_IR_A).click();
        MoveToElement(driver,O_Corporate.LNK_GEST_PED_CORP);
        driver.findElement(O_Corporate.LNK_MAESTRO_PEDIDO).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_ID_RESERVA3));
        ExtentReportUtil.INSTANCE.stepPass(driver, "Maestro de Pedido");
        generateWord.sendText("Maestro de Pedido");
        generateWord.addImageToWord(driver);
    }
    public void buscamosElIDReservaMaestroDePedidoCorporativo(String casoDePrueba) throws Throwable {

        try {
            Thread.sleep(2000);
            WebDriverWait wait = new WebDriverWait(driver, 60);
            int pedido1 = Integer.parseInt(casoDePrueba) - 1;
            String user1 = getData().get(pedido1).get(ExcelCorporativo.IDRESERVA);
            driver.findElement(O_Corporate.TXT_ID_RESERVA3).clear();
            driver.findElement(O_Corporate.TXT_ID_RESERVA3).sendKeys(user1);
            driver.findElement(O_Corporate.TXT_ID_RESERVA3).sendKeys(Keys.ENTER);
            Thread.sleep(2000);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.LNK_PEDIDO));
            ExtentReportUtil.INSTANCE.stepPass(driver, "ID de reserva encontrado");
            generateWord.sendText("ID de reserva encontrado");
            generateWord.addImageToWord(driver);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void verificamosElEstadoDelPedidoCorporativo() throws Exception {

        driver.findElement(O_Corporate.LNK_PEDIDO).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_ESTADO_PEDIDO));
        String estado;

        estado = driver.findElement(O_Corporate.TXT_ESTADO_PEDIDO).getText();
        if  (estado.equals("DESPACHADO")){
            ExtentReportUtil.INSTANCE.stepPass(driver, "Pedido despachado");
            generateWord.sendText("Pedido despachado");
            generateWord.addImageToWord(driver);

        }

    }
    public void entregarPedidoCorporativo() throws Exception {
        try {

            driver.findElement(O_Corporate.BTN_ENTREGA_PEDIDO).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ENTREGA));
            driver.findElement(O_Corporate.BTN_ENTREGA).click();
            Thread.sleep(1000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Entregar Pedido");
            generateWord.sendText("Entregar Pedido");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_SI_4).click();
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ACEPTAR_MENS_SIST));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje del sistema");
            generateWord.sendText("Mensaje del sistema");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
            Thread.sleep(7000);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void validarEstadoDelPedidoCorporativo() throws Exception {
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_ESTADO_PEDIDO));
        String estado;
        estado = driver.findElement(O_Corporate.TXT_ESTADO_PEDIDO).getText();
        if  (estado.equals("ENTREGADO")){
            ExtentReportUtil.INSTANCE.stepPass(driver, "Pedido entregado");
            generateWord.sendText("Pedido entregado");
            generateWord.addImageToWord(driver);
        }
    }
    public void seleccionamosAsignaciónDeSeriesCorporativos() throws Exception {
        driver.findElement(O_Corporate.LST_IR_A).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.LNK_GEST_PED_CORP));
        MoveToElement(driver,O_Corporate.LNK_GEST_PED_CORP);
        MoveToElement(driver,O_Corporate.LNK_PREP_PEDIDO);
        driver.findElement(O_Corporate.LNK_ASIG_SERIES).click();
        Thread.sleep(3000);
        ExtentReportUtil.INSTANCE.stepPass(driver, "Asignación de series.");
        generateWord.sendText("Asignación de series.");
        generateWord.addImageToWord(driver);
    }
    public void buscamosElPedidoCorporativo(String casoDePrueba) throws Throwable {
        try {
            Thread.sleep(2000);
            int pedido1 = Integer.parseInt(casoDePrueba) - 1;
            String user1 = getData().get(pedido1).get(ExcelCorporativo.IDPEDIDO);
            int tipo = Integer.parseInt(casoDePrueba) - 1;
            flujo = getData().get(tipo).get(ExcelCorporativo.TIPO_FLUJO);
            driver.findElement(O_Corporate.TXT_INGRESAR_PEDIDO).clear();
            driver.findElement(O_Corporate.TXT_INGRESAR_PEDIDO).sendKeys(user1);
            driver.findElement(O_Corporate.TXT_INGRESAR_PEDIDO).sendKeys(Keys.ENTER);
            String f = driver.findElement(O_Corporate.TXT_PEDIDO).getText();
            while (!f.equals(user1)) {
                Thread.sleep(1000);
                String g;
                g = driver.findElement(O_Corporate.TXT_VACIO).getText();
                if (g.equals("0 - 0 de 0")){
                    ExtentReportUtil.INSTANCE.stepPass(driver, "ID Reserva no encontrado");
                    generateWord.sendText("ID Reserva no encontrado");
                    generateWord.addImageToWord(driver);
                    driver.quit();
                }else{
                    f = driver.findElement(O_Corporate.TXT_PEDIDO).getText();
                }
            }

            ExtentReportUtil.INSTANCE.stepPass(driver, "Pedido encontrado");
            generateWord.sendText("Pedido encontrado");
            generateWord.addImageToWord(driver);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void ingresamosMaterialesIMEIYSIMCARDS(String casoDePrueba) throws Throwable {
        try {
            int CantFilas = Integer.parseInt(casoDePrueba) - 1;
            String filas;
            filas = driver.findElement(O_Corporate.TABLE2).getAttribute("displayrows");
            int num = Integer.parseInt(filas);
            int s = CantFilas;
            int f = CantFilas;
            for (int  i =0; (i<num); i++){
                String TipoMaterial = driver.findElement(By.id("md6723283_tdrow_[C:6]-c[R:"+i+"]")).getText();
                if (TipoMaterial.equals("IMEI")){
                    String imei = getData().get(s).get(ExcelCorporativo.NUMERO_IMEI);
                    driver.findElement(By.id("md6723283_tdrow_[C:9]_txt-tb[R:"+i+"]")).clear();
                    driver.findElement(By.id("md6723283_tdrow_[C:9]_txt-tb[R:"+i+"]")).sendKeys(imei);
                    Thread.sleep(1000);
                    s++;
                }
                if (TipoMaterial.equals("ICCID")){
                    String sim = getData().get(f).get(ExcelCorporativo.NUMERO_SIMCARD);
                    driver.findElement(By.id("md6723283_tdrow_[C:9]_txt-tb[R:"+i+"]")).clear();
                    driver.findElement(By.id("md6723283_tdrow_[C:9]_txt-tb[R:"+i+"]")).sendKeys(sim);
                    Thread.sleep(1000);
                    f++;
                }
                if (i==num-1){
                    ExtentReportUtil.INSTANCE.stepFail(driver, "Materiales ingresados");
                    generateWord.sendText("Materiales ingresados");
                    generateWord.addImageToWord(driver);
                    break;
                }
            }
            Thread.sleep(2000);

/*
            int CantFilas = Integer.parseInt(casoDePrueba) - 1;
            String Cant_Filas = getData().get(CantFilas).get(CANT_FILAS);
            int Pedido = Integer.parseInt(casoDePrueba) - 1;
            String Pedidos = getData().get(Pedido).get(TIPO_PEDIDO);
            if (Cant_Filas.equals("6")&& Pedidos.equals("EQUIPO+SIM")){
                int imei1 = Integer.parseInt(casoDePrueba) - 1;
                String imei_1 = getData().get(imei1).get(NUMERO_IMEI);
                driver.findElement(TXT_MATERIAL_1).clear();
                driver.findElement(TXT_MATERIAL_1).sendKeys(imei_1);
                int sim1 = Integer.parseInt(casoDePrueba) - 1;
                String sim_1 = getData().get(sim1).get(NUMERO_SIMCARD);
                driver.findElement(TXT_MATERIAL_2).clear();
                driver.findElement(TXT_MATERIAL_2).sendKeys(sim_1);
                String imei_2 = getData().get(1).get(NUMERO_IMEI);
                driver.findElement(TXT_MATERIAL_3).clear();
                driver.findElement(TXT_MATERIAL_3).sendKeys(imei_2);
                String sim_2 = getData().get(1).get(NUMERO_SIMCARD);
                driver.findElement(TXT_MATERIAL_4).clear();
                driver.findElement(TXT_MATERIAL_4).sendKeys(sim_2);
                String imei_3 = getData().get(2).get(NUMERO_IMEI);
                driver.findElement(TXT_MATERIAL_5).clear();
                driver.findElement(TXT_MATERIAL_5).sendKeys(imei_3);
                String sim_3 = getData().get(2).get(NUMERO_SIMCARD);
                driver.findElement(TXT_MATERIAL_6).clear();
                driver.findElement(TXT_MATERIAL_6).sendKeys(sim_3);
                String imei_4 = getData().get(3).get(NUMERO_IMEI);
                driver.findElement(TXT_MATERIAL_7).clear();
                driver.findElement(TXT_MATERIAL_7).sendKeys(imei_4);
                String sim_4 = getData().get(3).get(NUMERO_SIMCARD);
                driver.findElement(TXT_MATERIAL_8).clear();
                driver.findElement(TXT_MATERIAL_8).sendKeys(sim_4);
                String imei_5 = getData().get(4).get(NUMERO_IMEI);
                driver.findElement(TXT_MATERIAL_9).clear();
                driver.findElement(TXT_MATERIAL_9).sendKeys(imei_5);
                String sim_5 = getData().get(4).get(NUMERO_SIMCARD);
                driver.findElement(TXT_MATERIAL_10).clear();
                driver.findElement(TXT_MATERIAL_10).sendKeys(sim_5);
                String imei_6 = getData().get(5).get(NUMERO_IMEI);
                driver.findElement(TXT_MATERIAL_11).clear();
                driver.findElement(TXT_MATERIAL_11).sendKeys(imei_6);
                String sim_6 = getData().get(5).get(NUMERO_SIMCARD);
                driver.findElement(TXT_MATERIAL_12).clear();
                driver.findElement(TXT_MATERIAL_12).sendKeys(sim_6);
            }
            if (Cant_Filas.equals("1")&& Pedidos.equals("EQUIPO+SIM")) {
                int imei1 = Integer.parseInt(casoDePrueba) - 1;
                String imei_1 = getData().get(imei1).get(NUMERO_IMEI);
                driver.findElement(TXT_MATERIAL_1).clear();
                driver.findElement(TXT_MATERIAL_1).sendKeys(imei_1);
                int sim1 = Integer.parseInt(casoDePrueba) - 1;
                String sim_1 = getData().get(sim1).get(NUMERO_SIMCARD);
                driver.findElement(TXT_MATERIAL_2).clear();
                driver.findElement(TXT_MATERIAL_2).sendKeys(sim_1);
            }
            if  (Cant_Filas.equals("1")&& Pedidos.equals("SOLO SIM")) {
                int sim1 = Integer.parseInt(casoDePrueba) - 1;
                String sim_1 = getData().get(sim1).get(NUMERO_SIMCARD);
                driver.findElement(TXT_MATERIAL_1).clear();
                driver.findElement(TXT_MATERIAL_1).sendKeys(sim_1);
            }
*/
        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);

        }

    }
    public void validamosSeriesCorporativos() throws Exception {
        try {
            driver.findElement(O_Corporate.BTN_VALIDAR_SERIE).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_SI));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Validar Serie");
            generateWord.sendText("Validar Serie");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_SI).click();
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ACEPTAR_MENS_SIST));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje de Validación");
            generateWord.sendText("Mensaje de Validación");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
            Thread.sleep(2000);

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void buscamosElPedidoCorporativos(String casoDePrueba) throws Throwable {

        try {
            Thread.sleep(2000);
            int pedido1 = Integer.parseInt(casoDePrueba) - 1;
            String user1 = getData().get(pedido1).get(ExcelCorporativo.IDPEDIDO);
            int tipo = Integer.parseInt(casoDePrueba) - 1;
            flujo = getData().get(tipo).get(ExcelCorporativo.TIPO_FLUJO);
            driver.findElement(O_Corporate.TXT_INGRESAR_PEDIDO).clear();
            driver.findElement(O_Corporate.TXT_INGRESAR_PEDIDO).sendKeys(user1);
            driver.findElement(O_Corporate.TXT_INGRESAR_PEDIDO).sendKeys(Keys.ENTER);
            String f = driver.findElement(O_Corporate.TXT_PEDIDO).getText();
            while (!f.equals(user1)) {
                Thread.sleep(1000);
                String g;
                g = driver.findElement(O_Corporate.TXT_VACIO).getText();
                if (g.equals("0 - 0 de 0")){
                    ExtentReportUtil.INSTANCE.stepPass(driver, "ID Reserva no encontrado");
                    generateWord.sendText("ID Reserva no encontrado");
                    generateWord.addImageToWord(driver);
                    driver.quit();
                }else{
                    f = driver.findElement(O_Corporate.TXT_PEDIDO).getText();
                }
            }

            ExtentReportUtil.INSTANCE.stepPass(driver, "Pedido encontrado");
            generateWord.sendText("Pedido encontrado");
            generateWord.addImageToWord(driver);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void verificarElEstadoDeAsignaciónDeSeriesCorporativos(String casoDePrueba) throws Throwable {

        try {
            String filas;
            int error = 0;
            int bien = 0;
            filas = driver.findElement(O_Corporate.TABLE2).getAttribute("displayrows");
            int num = Integer.parseInt(filas);
            for(int  i =0; (i<num); i++){
                String estado = driver.findElement(By.id("md6723283_tdrow_[C:10]-c[R:"+i+"]")).getText();
                if (estado.equals("ERROR")){
                    error++;
                    String valor = driver.findElement(By.id("md6723283_tdrow_[C:9]_txt-tb[R:"+i+"]")).getAttribute("value");
                    System.out.println(valor +" --> " + estado);
                }
                if (estado.equals("VALIDADO")){
                    bien++;
                    String valor = driver.findElement(By.id("md6723283_tdrow_[C:9]_txt-tb[R:"+i+"]")).getAttribute("value");
                    System.out.println(valor +" --> " + estado);
                }
                if (i == num-1){
                    if (error>0){
                        System.out.println("Se obtuvo un total de: " +error+ " materiales con error");
                        System.out.println("Se obtuvo un total de: " +bien+ " materiales validados");
                        ExtentReportUtil.INSTANCE.stepFail(driver, "Materiales no validados");
                        generateWord.sendText("Materiales no validados");
                        generateWord.addImageToWord(driver);
                        driver.quit();
                        break;
                    }else{
                        System.out.println("Se obtuvo un total de: " +bien+ " materiales validados");
                        ExtentReportUtil.INSTANCE.stepFail(driver, "Materiales validados");
                        generateWord.sendText("Materiales validados");
                        generateWord.addImageToWord(driver);
                        break;
                    }
                }
            }

            /*
            int CantFilas = Integer.parseInt(casoDePrueba) - 1;
            String Cant_Filas = getData().get(CantFilas).get(CANT_FILAS);
            int Pedido = Integer.parseInt(casoDePrueba) - 1;
            String Pedidos = getData().get(Pedido).get(TIPO_PEDIDO);
            if (Cant_Filas.equals("6")&& Pedidos.equals("EQUIPO+SIM")){
                String Estado1 = driver.findElement(LBL_ESTADO_VAL_SERIE1).getText();
                String Estado2 = driver.findElement(LBL_ESTADO_VAL_SERIE2).getText();
                String Estado3 = driver.findElement(LBL_ESTADO_VAL_SERIE3).getText();
                String Estado4 = driver.findElement(LBL_ESTADO_VAL_SERIE4).getText();
                String Estado5 = driver.findElement(LBL_ESTADO_VAL_SERIE5).getText();
                String Estado6 = driver.findElement(LBL_ESTADO_VAL_SERIE6).getText();
                String Estado7 = driver.findElement(LBL_ESTADO_VAL_SERIE7).getText();
                String Estado8 = driver.findElement(LBL_ESTADO_VAL_SERIE8).getText();
                String Estado9 = driver.findElement(LBL_ESTADO_VAL_SERIE9).getText();
                String Estado10 = driver.findElement(LBL_ESTADO_VAL_SERIE10).getText();
                String Estado11 = driver.findElement(LBL_ESTADO_VAL_SERIE11).getText();
                String Estado12 = driver.findElement(LBL_ESTADO_VAL_SERIE12).getText();
                if (Estado1.equals("VALIDADO") && Estado2.equals("VALIDADO") && Estado3.equals("VALIDADO") && Estado4.equals("VALIDADO")&& Estado5.equals("VALIDADO") && Estado6.equals("VALIDADO") && Estado7.equals("VALIDADO") && Estado8.equals("VALIDADO") && Estado9.equals("VALIDADO") && Estado10.equals("VALIDADO") && Estado11.equals("VALIDADO") && Estado12.equals("VALIDADO")){
                    ExtentReportUtil.INSTANCE.stepPass(driver, "Estado de Validación de series: VALIDADO");
                    generateWord.sendText("Estado de Validación de series: VALIDADO");
                    generateWord.addImageToWord(driver);
                }else {
                    ExtentReportUtil.INSTANCE.stepFail(driver, "IMEI Y SIMCARD no validados");
                    generateWord.sendText("IMEI Y SIMCARD no validados");
                    generateWord.addImageToWord(driver);
                }
            }
            if (Cant_Filas.equals("1")&& Pedidos.equals("EQUIPO+SIM")){
                String Estado1 = driver.findElement(LBL_ESTADO_VAL_SERIE1).getText();
                String Estado2 = driver.findElement(LBL_ESTADO_VAL_SERIE2).getText();
                if (Estado1.equals("VALIDADO") && Estado2.equals("VALIDADO")){
                    ExtentReportUtil.INSTANCE.stepPass(driver, "Estado de Validación de series: VALIDADO");
                    generateWord.sendText("Estado de Validación de series: VALIDADO");
                    generateWord.addImageToWord(driver);
                }else {

                        ExtentReportUtil.INSTANCE.stepFail(driver, "IMEI Y SIMCARD no validados");
                        generateWord.sendText("IMEI Y SIMCARD no validados");
                        generateWord.addImageToWord(driver);
                        driver.quit();
                }
            }
            if (Cant_Filas.equals("1")&& Pedidos.equals("SOLO SIM")){
                String Estado1 = driver.findElement(LBL_ESTADO_VAL_SERIE1).getText();
                if (Estado1.equals("VALIDADO")){
                    ExtentReportUtil.INSTANCE.stepPass(driver, "Estado de Validación de series: VALIDADO" );
                    generateWord.sendText("Estado de Validación de series: VALIDADO");
                    generateWord.addImageToWord(driver);
                }else {
                    ExtentReportUtil.INSTANCE.stepFail(driver, "Materiales no validados");
                    generateWord.sendText("Materiales no validados");
                    generateWord.addImageToWord(driver);
                }
            }*/

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void seleccionamosImpresionDeDocumentosCorporativos() throws Exception {

        driver.findElement(O_Corporate.LST_IR_A).click();
        MoveToElement(driver,O_Corporate.LNK_GEST_PED_CORP);
        MoveToElement(driver,O_Corporate.LNK_PREP_PEDIDO);
        driver.findElement(O_Corporate.LNK_IMPR_DOC).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_ID_RESERVA2));
        ExtentReportUtil.INSTANCE.stepPass(driver, "Impresión de documentos");
        generateWord.sendText("Impresión de documentos");
        generateWord.addImageToWord(driver);
    }
    public void ContratoDigitalCorporativo() throws Exception {
        try {

            driver.findElement(O_Corporate.BTN_CARGA_CONTRATO_DIGITAL).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_SI_CARGA_CONTRATO_DIGITAL));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Contrato Digital");
            generateWord.sendText("Contrato Digital");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_SI_CARGA_CONTRATO_DIGITAL).click();
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ACEPTAR_MENS_SIST));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje del sistema");
            generateWord.sendText("Mensaje del sistema");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
            Thread.sleep(2000);

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }

    } public void preparaciónDeLaFacturaCorporativos() throws Exception {
        try {
            driver.findElement(O_Corporate.CMB_SELECT_ACTION).click();
            Thread.sleep(1000);
            driver.findElement(O_Corporate.CMB_SELECT_ACTION_PREP_FACTURA).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ACEPTAR_MENS_SIST));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje del sistema");
            generateWord.sendText("Mensaje del sistema");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_SI_2));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Preparación de la factura");
            generateWord.sendText("Preparación de la factura");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_SI_2).click();
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ACEPTAR_MENS_SIST));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje del sistema");
            generateWord.sendText("Mensaje del sistema");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
            Thread.sleep(2000);

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void impresiónDeLaFacturaCorporativos() throws Exception {
        try {

            driver.findElement(O_Corporate.CMB_SELECT_ACTION).click();
            Thread.sleep(1000);
            driver.findElement(O_Corporate.CMB_SELECT_ACTION_IMPR_FACTURA).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_SI_3));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Impresión de la factura");
            generateWord.sendText("Impresión de la factura");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_SI_3).click();
            Thread.sleep(5000);
            String parentWindow = driver.getWindowHandle();
            Set<String> handles = driver.getWindowHandles();
            for (String windowHandle : handles) {
                if (!windowHandle.equals(parentWindow)) {
                    ExtentReportUtil.INSTANCE.stepPass(driver.switchTo().window(windowHandle), "Reporte de impresion de guia de remision");
                    generateWord.sendText("Reporte de impresion de guia de remision");
                    generateWord.addImageToWord(driver.switchTo().window(windowHandle));
                    driver.switchTo().window(windowHandle).close();
                    driver.switchTo().window(parentWindow);
                    break;
                }
            }


        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void ejecutarInformesCorporativos() throws Exception {
        try {
            driver.findElement(O_Corporate.CMB_SELECT_ACTION).click();
            Thread.sleep(1000);
            driver.findElement(O_Corporate.CMB_SELECT_ACTION_EJEC_INFORMES).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.LNK_IMPR_GUIA_REMISION));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Ejecutar informes");
            generateWord.sendText("Ejecutar informes");
            generateWord.addImageToWord(driver);

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void imprimirGuiaDeRemisiónCorporativos() throws Exception {
        try {

            driver.findElement(O_Corporate.LNK_IMPR_GUIA_REMISION).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_SERIE_GUIA_REMISION));
            driver.findElement(O_Corporate.TXT_SERIE_GUIA_REMISION).clear();
            driver.findElement(O_Corporate.TXT_SERIE_GUIA_REMISION).sendKeys("12345");
            Thread.sleep(100);
            driver.findElement(O_Corporate.TXT_CORRELATIVO_GUIA).clear();
            driver.findElement(O_Corporate.TXT_CORRELATIVO_GUIA).sendKeys("12345");
            Thread.sleep(100);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Imprimir guia de remisión");
            generateWord.sendText("Imprimir guia de remisión");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ACEPTAR).click();
            Thread.sleep(3000);
            Boolean isPresent = driver.findElements(O_Corporate.TXT_IMAGEN).size() > 0;
            if (isPresent.equals(true)){
                ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje error");
                generateWord.sendText("Mensaje error");
                generateWord.addImageToWord(driver);
                driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
                driver.findElement(O_Corporate.TXT_SERIE_GUIA_REMISION).clear();
                driver.findElement(O_Corporate.TXT_SERIE_GUIA_REMISION).sendKeys("12345");
                Thread.sleep(100);
                driver.findElement(O_Corporate.TXT_CORRELATIVO_GUIA).clear();
                driver.findElement(O_Corporate.TXT_CORRELATIVO_GUIA).sendKeys("12345");
                Thread.sleep(100);
                ExtentReportUtil.INSTANCE.stepPass(driver, "Imprimir guia de remisión");
                generateWord.sendText("Imprimir guia de remisión");
                generateWord.addImageToWord(driver);
                driver.findElement(O_Corporate.BTN_ACEPTAR).click();
                Thread.sleep(3000);
                Boolean isPresent2 = driver.findElements(O_Corporate.TXT_IMAGEN).size() > 0;
                if (isPresent2.equals(true)){
                    driver.quit();
                }
            }
            String parentWindow = driver.getWindowHandle();
            Set<String> handles = driver.getWindowHandles();
            for (String windowHandle : handles) {
                if (!windowHandle.equals(parentWindow)) {
                    WebDriverWait wait2 = new WebDriverWait(driver.switchTo().window(windowHandle), 60);
                    wait2.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_NUMRUC));
                    ExtentReportUtil.INSTANCE.stepPass(driver.switchTo().window(windowHandle), "Reporte de impresion de guia de remision");
                    generateWord.sendText("Reporte de impresion de guia de remision");
                    generateWord.addImageToWord(driver.switchTo().window(windowHandle));
                    driver.switchTo().window(windowHandle).close();
                    driver.switchTo().window(parentWindow);
                    break;
                }
            }
            ExtentReportUtil.INSTANCE.stepPass(driver, "Informes y programaciones");
            generateWord.sendText("Informes y programaciones");
            generateWord.addImageToWord(driver);

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void imprimirEtiquetaDeLineaCorporativos() throws Exception {
        try {
            driver.findElement(O_Corporate.LNK_IMPR_ETIQUETA_LINEA).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.RBTN_INMEDIATO));
            driver.findElement(O_Corporate.RBTN_INMEDIATO).click();
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ENVIAR2));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Imprimir etiqueta de linea");
            generateWord.sendText("Imprimir etiqueta de linea");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ENVIAR2).click();
            Thread.sleep(1000);
            String parentWindow = driver.getWindowHandle();
            Set<String> handles = driver.getWindowHandles();
            for (String windowHandle : handles) {
                if (!windowHandle.equals(parentWindow)) {
                    WebDriverWait wait2 = new WebDriverWait(driver.switchTo().window(windowHandle), 60);
                    wait2.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_SERVICIO));
                    ExtentReportUtil.INSTANCE.stepPass(driver.switchTo().window(windowHandle), "Reporte de etiqueta en linea");
                    generateWord.sendText("Reporte de etiqueta en linea");
                    generateWord.addImageToWord(driver.switchTo().window(windowHandle));
                    driver.switchTo().window(windowHandle).close();
                    driver.switchTo().window(parentWindow);
                    break;
                }
            }
            ExtentReportUtil.INSTANCE.stepPass(driver, "Informes y programaciones");
            generateWord.sendText("Informes y programaciones");
            generateWord.addImageToWord(driver);

        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);

        }
    }
    public void imprimirEtiquetaCorporativos() throws Exception {
        try {
            driver.findElement(O_Corporate.LNK_IMPR_ETIQUETA).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.RBTN_INMEDIATO2));
            driver.findElement(O_Corporate.RBTN_INMEDIATO2).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Imprimir etiqueta");
            generateWord.sendText("Imprimir etiqueta");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ENVIAR3).click();
            Thread.sleep(1000);
            String parentWindow = driver.getWindowHandle();
            Set<String> handles = driver.getWindowHandles();
            for (String windowHandle : handles) {
                if (!windowHandle.equals(parentWindow)) {
                    WebDriverWait wait2 = new WebDriverWait(driver.switchTo().window(windowHandle), 60);
                    wait2.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_NUMRUC2));
                    ExtentReportUtil.INSTANCE.stepPass(driver.switchTo().window(windowHandle), "Reporte de etiqueta en linea");
                    generateWord.sendText("Reporte de etiqueta en linea");
                    generateWord.addImageToWord(driver.switchTo().window(windowHandle));
                    driver.switchTo().window(windowHandle).close();
                    driver.switchTo().window(parentWindow);
                    break;
                }
            }
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_CANCELAR));
            driver.findElement(O_Corporate.BTN_CANCELAR).click();
            Thread.sleep(2000);

        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);

        }
    }
    public void observamosElCambioDeEstadoC() throws Exception {
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_RUTA_FACTURA));
        driver.findElement(O_Corporate.TXT_RUTA_FACTURA).click();
        ExtentReportUtil.INSTANCE.stepPass(driver, "Impresión de documentos terminado");
        generateWord.sendText("Impresión de documentos terminado");
        generateWord.addImageToWord(driver);
        Thread.sleep(2000);
    }
    public void seleccionamosDespachoDeHubCorporativos() throws Exception {
        Thread.sleep(2000);
        driver.findElement(O_Corporate.LST_IR_A).click();
        MoveToElement(driver,O_Corporate.LNK_GEST_PED_CORP);
        driver.findElement(O_Corporate.LNK_DESPACHO_HUB).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_ID_RESERVA2));
        ExtentReportUtil.INSTANCE.stepPass(driver, "Despacho al HUB");
        generateWord.sendText("Despacho al HUB");
        generateWord.addImageToWord(driver);
    }
    public void ejecuciónDeInformesCorporativos() throws Exception {
        try {

            driver.findElement(O_Corporate.CMB_SELECT_ACTION).click();
            Thread.sleep(1000);
            driver.findElement(O_Corporate.CMB_SELECT_ACTION_EJEC_INFORMES).click();
            Thread.sleep(5000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Informes y programaciones");
            generateWord.sendText("Informes y programaciones");
            generateWord.addImageToWord(driver);

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void generaciónDeMasterBOXCorporativos() throws Exception {
        try {
            driver.findElement(O_Corporate.LNK_GENERACION_MASTERBOX).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_MONTO_CAMION));
            driver.findElement(O_Corporate.TXT_MONTO_CAMION).clear();
            driver.findElement(O_Corporate.TXT_MONTO_CAMION).sendKeys("354445");
            Thread.sleep(1000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Generación de Master box");
            generateWord.sendText("Generación de Master box");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ENVIAR4).click();
            Thread.sleep(2000);
            String parentWindow = driver.getWindowHandle();
            Set<String> handles = driver.getWindowHandles();
            for (String windowHandle : handles) {
                if (!windowHandle.equals(parentWindow)) {
                    WebDriverWait wait2 = new WebDriverWait(driver.switchTo().window(windowHandle), 60);
                    wait2.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_MASTER));
                    ExtentReportUtil.INSTANCE.stepPass(driver.switchTo().window(windowHandle), "Informe generación de master box");
                    generateWord.sendText("Informe generación de master box");
                    generateWord.addImageToWord(driver.switchTo().window(windowHandle));
                    driver.switchTo().window(windowHandle).close();
                    driver.switchTo().window(parentWindow);
                    break;
                }
            }

            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_CANCELAR2));
            driver.findElement(O_Corporate.BTN_CANCELAR2).click();
            Thread.sleep(2000);

        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);

        }
    }
    public void despacharPedidoCorporativos() throws Exception {
        try {

            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.LBL_COD_MASTERBOX));
            String Cod_Master_Box = driver.findElement(O_Corporate.LBL_COD_MASTERBOX).getText();
            Fecha = driver.findElement(O_Corporate.LBL_FECHA_DESPACHO).getText();
            Thread.sleep(1000);
            driver.findElement(O_Corporate.BTN_DESPACHAR_PEDIDO).click();
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_MASTERBOX));
            driver.findElement(O_Corporate.TXT_MASTERBOX).clear();
            driver.findElement(O_Corporate.TXT_MASTERBOX).sendKeys(Cod_Master_Box);
            /*SimpleDateFormat df = new SimpleDateFormat("dd/MM/YYYY");
            Date dt = new Date();
            Calendar cl = Calendar.getInstance();
            cl.setTime(dt);;
            cl.add(Calendar.DAY_OF_MONTH, 0);
            dt=cl.getTime();
            String str = df.format(dt);*/
            Thread.sleep(1000);
            driver.findElement(O_Corporate.TXT_FECHA_PEDIDO).clear();
            driver.findElement(O_Corporate.TXT_FECHA_PEDIDO).sendKeys(Fecha);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Despachar Pedido");
            generateWord.sendText("Despachar Pedido");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_SI4).click();
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ACEPTAR_MENS_SIST));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje de sistema");
            generateWord.sendText("Mensaje de sistema");
            generateWord.addImageToWord(driver);
            String txt = driver.findElement(By.id("mb_msg")).getText();
            if (txt.contains("La fecha de despacho es requerida. ")){
                driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
                Thread.sleep(1000);
                driver.findElement(O_Corporate.TXT_FECHA_PEDIDO).clear();
                driver.findElement(O_Corporate.TXT_FECHA_PEDIDO).sendKeys(Fecha);
                ExtentReportUtil.INSTANCE.stepPass(driver, "Despachar Pedido");
                generateWord.sendText("Despachar Pedido");
                generateWord.addImageToWord(driver);
                driver.findElement(O_Corporate.BTN_SI4).click();
                wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ACEPTAR_MENS_SIST));
                if (txt.contains("La fecha de despacho es requerida. ")){
                    ExtentReportUtil.INSTANCE.stepFail(driver, "Despachar Pedido Fallido");
                    generateWord.sendText("Despachar Pedido Fallido");
                    generateWord.addImageToWord(driver);
                    driver.quit();
                }

            }
            driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
            Thread.sleep(1000);

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void seleccionRecepciónDePedidosCorporativos() throws Exception {

        Thread.sleep(2000);
        driver.findElement(O_Corporate.LST_IR_A).click();
        MoveToElement(driver,O_Corporate.LNK_GEST_PED_CORP);
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_ID_RESERVA2));
        ExtentReportUtil.INSTANCE.stepPass(driver, "Recepción de pedido");
        generateWord.sendText("Recepción de pedido");
        generateWord.addImageToWord(driver);
    }
    public void recepcionarDePedidosCorporativos() throws Exception {
        try {

            String Cod_Master_Box2 = driver.findElement(O_Corporate.LBL_COD_MASTERBOX2).getText();
            Fecha = driver.findElement(O_Corporate.LBL_FECHA_DESPACHO).getText();
            Thread.sleep(1000);
            driver.findElement(O_Corporate.BTN_RECEP_PEDIDO).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_MASTERBOX2));
            driver.findElement(O_Corporate.TXT_MASTERBOX2).sendKeys(Cod_Master_Box2);
            /*SimpleDateFormat df = new SimpleDateFormat("dd/MM/YYYY");
            Date dt = new Date();
            Calendar cl = Calendar.getInstance();
            cl.setTime(dt);;
            cl.add(Calendar.DAY_OF_MONTH, 0);
            dt=cl.getTime();
            String str = df.format(dt);*/
            driver.findElement(O_Corporate.TXT_FECHA_PEDIDO2).sendKeys(Fecha);
            Thread.sleep(1000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Recepcionar pedido");
            generateWord.sendText("Recepcionar pedido");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_SI5).click();
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ACEPTAR_MENS_SIST));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje del sistema");
            generateWord.sendText("Mensaje del sistema");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
            Thread.sleep(2000);


        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void procesoDeLógicaDeRuteoCorporativos() throws Exception {
        try {

            driver.findElement(O_Corporate.LNK_PROC_LOGICARUTEO).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_FECHA_DESPACHO));
            /*SimpleDateFormat df = new SimpleDateFormat("dd/MM/YYYY");
            Date dt = new Date();
            Calendar cl = Calendar.getInstance();
            cl.setTime(dt);;
            cl.add(Calendar.DAY_OF_MONTH, 0);
            dt=cl.getTime();
            String str = df.format(dt);*/
            driver.findElement(O_Corporate.TXT_FECHA_DESPACHO).clear();
            driver.findElement(O_Corporate.TXT_FECHA_DESPACHO).sendKeys(Fecha);
            Thread.sleep(1000);
            driver.findElement(O_Corporate.TXT_ALMACEN).clear();
            driver.findElement(O_Corporate.TXT_ALMACEN).sendKeys("=PE10API7");
            Thread.sleep(1000);
            driver.findElement(O_Corporate.TXT_TURNO).clear();
            driver.findElement(O_Corporate.TXT_TURNO).sendKeys("=TARDE_API7");
            Thread.sleep(1000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Proceso de logica de ruteo");
            generateWord.sendText("Proceso de logica de ruteo");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ENVIAR5).click();
            Thread.sleep(1000);
            Boolean isPresent = driver.findElements(O_Corporate.TXT_IMAGEN).size() > 0;
            if (isPresent.equals(true)){
                ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje error");
                generateWord.sendText("Mensaje error");
                generateWord.addImageToWord(driver);
                String f = driver.findElement(O_Corporate.TXT_IMAGEN).getText();
                if (f.contains("Fecha de despacho es un campo necesario") || f.contains("Almacén es un campo necesario") || f.contains("Turno es un campo necesario")){
                    driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
                    driver.findElement(O_Corporate.TXT_FECHA_DESPACHO).clear();
                    driver.findElement(O_Corporate.TXT_FECHA_DESPACHO).sendKeys(Fecha);
                    Thread.sleep(1000);
                    driver.findElement(O_Corporate.TXT_ALMACEN).clear();
                    driver.findElement(O_Corporate.TXT_ALMACEN).sendKeys("=PE10API7");
                    Thread.sleep(1000);
                    driver.findElement(O_Corporate.TXT_TURNO).clear();
                    driver.findElement(O_Corporate.TXT_TURNO).sendKeys("=TARDE_API7");
                    Thread.sleep(1000);
                    ExtentReportUtil.INSTANCE.stepPass(driver, "Proceso de logica de ruteo");
                    generateWord.sendText("Proceso de logica de ruteo");
                    generateWord.addImageToWord(driver);
                    driver.findElement(O_Corporate.BTN_ENVIAR5).click();
                    Boolean isPresent2 = driver.findElements(O_Corporate.TXT_IMAGEN).size() > 0;
                    if (isPresent2.equals(true)){
                        driver.quit();
                    }
                }

            }

            String parentWindow = driver.getWindowHandle();
            Set<String> handles = driver.getWindowHandles();
            for (String windowHandle : handles) {
                if (!windowHandle.equals(parentWindow)) {
                    WebDriverWait wait2 = new WebDriverWait(driver.switchTo().window(windowHandle), 60);
                    wait2.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_FECHA));
                    ExtentReportUtil.INSTANCE.stepPass(driver.switchTo().window(windowHandle), "Proceso de lógica de ruteo");
                    generateWord.sendText("Proceso de lógica de ruteo");
                    generateWord.addImageToWord(driver.switchTo().window(windowHandle));
                    driver.switchTo().window(windowHandle).close();
                    driver.switchTo().window(parentWindow);
                    break;
                }
            }
            ExtentReportUtil.INSTANCE.stepPass(driver, "Informes y programaciones");
            generateWord.sendText("Informes y programaciones");
            generateWord.addImageToWord(driver);
            Thread.sleep(2000);


        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);

        }
    }
    public void ReporteDeLogicaDeRuteoCorporativos() throws Exception {
        try {

            driver.findElement(O_Corporate.LNK_REPORTE_LOGICARUTEO).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_FECHA_DESPACHO2));
            /*SimpleDateFormat df = new SimpleDateFormat("dd/MM/YYYY");
            Date dt = new Date();
            Calendar cl = Calendar.getInstance();
            cl.setTime(dt);;
            cl.add(Calendar.DAY_OF_MONTH, 0);
            dt=cl.getTime();
            String str = df.format(dt);*/
            driver.findElement(O_Corporate.TXT_FECHA_DESPACHO2).clear();
            driver.findElement(O_Corporate.TXT_FECHA_DESPACHO2).sendKeys(Fecha);
            Thread.sleep(1000);
            driver.findElement(O_Corporate.TXT_ALMACEN2).clear();
            driver.findElement(O_Corporate.TXT_ALMACEN2).sendKeys("=PE10API7");
            Thread.sleep(1000);
            driver.findElement(O_Corporate.TXT_TURNO2).clear();
            driver.findElement(O_Corporate.TXT_TURNO2).sendKeys("=TARDE_API7");
            Thread.sleep(1000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Reporte de logica de ruteo");
            generateWord.sendText("Reporte de logica de ruteo");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ENVIAR6).click();
            Thread.sleep(10000);
            String parentWindow = driver.getWindowHandle();
            Set<String> handles = driver.getWindowHandles();
            for (String windowHandle : handles) {
                if (!windowHandle.equals(parentWindow)) {

                    ExtentReportUtil.INSTANCE.stepPass(driver.switchTo().window(windowHandle), "Proceso de lógica de ruteo");
                    generateWord.sendText("Proceso de lógica de ruteo");
                    generateWord.addImageToWord(driver.switchTo().window(windowHandle));
                    driver.switchTo().window(windowHandle).close();
                    driver.switchTo().window(parentWindow);
                    break;
                }
            }
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_CANCELAR3));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Informes y programaciones");
            generateWord.sendText("Informes y programaciones");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_CANCELAR3).click();
            Thread.sleep(2000);

        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);

        }
    }
    public void SeleccionamosRecepciónDePedidosCorporativos() throws Exception {

        Thread.sleep(2000);
        driver.findElement(O_Corporate.LST_IR_A).click();
        MoveToElement(driver,O_Corporate.LNK_GEST_PED_CORP);
        driver.findElement(O_Corporate.LNK_RECEPCION_PEDIDO).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_ID_RESERVA2));
        ExtentReportUtil.INSTANCE.stepPass(driver, "Recepción de pedido");
        generateWord.sendText("Recepción de pedido");
        generateWord.addImageToWord(driver);
    }
    public void EjecutarCargaDeRuteoCorporativos() throws Exception {
        Thread.sleep(2000);
        driver.findElement(O_Corporate.BTN_EJECUTAR_CARGA_RUTEO).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_SI6));
        ExtentReportUtil.INSTANCE.stepPass(driver, "Carga de logica de ruteo");
        generateWord.sendText("Carga de logica de ruteo");
        generateWord.addImageToWord(driver);
        driver.findElement(O_Corporate.BTN_SI6).click();
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ACEPTAR_MENS_SIST));
        /*String f = driver.findElement(By.id("mb_msg")).getText();
        if(f.contains("No se han cargado Datos para procesar, verificar el archivo adjunto y cargarlo nuevamente")){
            ExtentReportUtil.INSTANCE.stepPass(driver, "Error al momento de cargar el archivo");
            generateWord.sendText("Error al momento de cargar el archivo");
            generateWord.addImageToWord(driver);
            System.out.println("Error al momento de cargar el archivo");
            driver.quit();
        }*/
        ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje del sistema");
        generateWord.sendText("Mensaje del sistema");
        generateWord.addImageToWord(driver);
        driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
        Thread.sleep(2000);
    }
    public void GuardamosNúmeroDeEnvioCorporativos() throws Exception {
        Thread.sleep(2000);
        driver.findElement(O_Corporate.TXT_ORDEN_ENVIO).click();
        Thread.sleep(1000);
        NUM_ENVIO = driver.findElement(O_Corporate.TXT_NUMERO_ENVIO).getText();
        System.out.println(NUM_ENVIO);
        if (NUM_ENVIO.equals("")){
            Thread.sleep(1000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Número de envío no generado");
            generateWord.sendText("Número de envío no generado");
            generateWord.addImageToWord(driver);
            driver.quit();
        }
        Thread.sleep(1000);
        ExtentReportUtil.INSTANCE.stepPass(driver, "Número de envío generado");
        generateWord.sendText("Número de envío generado");
        generateWord.addImageToWord(driver);
    }
    public void DespachoAMotorizadoCorporativos() throws Exception {

        driver.findElement(O_Corporate.LST_IR_A).click();
        MoveToElement(driver,O_Corporate.LNK_GEST_PED_CORP);
        driver.findElement(O_Corporate.LNK_DESPACHO_MOTORIZADO).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_ENVIO));
        ExtentReportUtil.INSTANCE.stepPass(driver, "Despacho a motorizado");
        generateWord.sendText("Despacho a motorizado");
        generateWord.addImageToWord(driver);

    }
    public void BuscamosNumeroDeEnvioCorporativos() throws Exception {
        Thread.sleep(2000);
        driver.findElement(O_Corporate.TXT_ENVIO).clear();
        driver.findElement(O_Corporate.TXT_ENVIO).sendKeys(NUM_ENVIO);
        driver.findElement(O_Corporate.TXT_ENVIO).sendKeys(Keys.ENTER);
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.LNK_NUM_ENVIO));
        ExtentReportUtil.INSTANCE.stepPass(driver, "Pedido encontrado");
        generateWord.sendText("Pedido encontrado");
        generateWord.addImageToWord(driver);
        driver.findElement(O_Corporate.LNK_NUM_ENVIO).click();
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_LUPA11));
    }
    public void BuscamosMotorizadoCorporativos() throws Exception {

        driver.findElement(O_Corporate.BTN_LUPA11).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_MOTORIZADO));
        driver.findElement(O_Corporate.TXT_MOTORIZADO).sendKeys("VMONTOVELI");
        driver.findElement(O_Corporate.TXT_MOTORIZADO).sendKeys(Keys.ENTER);
        Thread.sleep(3000);
        ExtentReportUtil.INSTANCE.stepPass(driver, "Seleccionamos motorizado");
        generateWord.sendText("Seleccionamos motorizado");
        generateWord.addImageToWord(driver);
        driver.findElement(O_Corporate.LNK_MOTORIZADO).click();
        Thread.sleep(3000);

    }
    public void DespachamosPedidoCorporativos() throws Throwable {
        Thread.sleep(2000);

        driver.findElement(O_Corporate.BTN_PEDIDOS).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_CAMBIAR_ESTADO));
        String filas;
        String d;
        filas = driver.findElement(O_Corporate.TABLE).getAttribute("displayrows");
        int num = Integer.parseInt(filas);
        for (int  i =0; (i<num); i++){
            String orden = driver.findElement(By.id("m187f4d3c_tdrow_[C:4]-c[R:"+i+"]")).getText();
            System.out.println(tipo);
            if (tipo.equals("SI")){
                System.out.println(orden);
                if (orden.equals(pedido)){
                    driver.findElement(By.id("m187f4d3c_tdrow_[C:0]_checkbox-cb[R:"+i+"]_img")).click();
                    break;
                }
            }
            if (tipo.equals("NO")){
                if (orden.equals(reserva)){
                    driver.findElement(By.id("m187f4d3c_tdrow_[C:0]_checkbox-cb[R:"+i+"]_img")).click();
                    break;
                }
            }

            if (i==num-1){
                ExtentReportUtil.INSTANCE.stepFail(driver, "No se encontro la orden");
                generateWord.sendText("No se encontro la orden");
                generateWord.addImageToWord(driver);
                driver.quit();
                break;
            }
        }

        //driver.findElement(BTN_CAMBIAR_ESTADO).click();
        Thread.sleep(1000);
        ExtentReportUtil.INSTANCE.stepPass(driver, "Pedido listo");
        generateWord.sendText("Pedido listo");
        generateWord.addImageToWord(driver);
        driver.findElement(O_Corporate.BTN_DESPACHAR_PEDIDO2).click();
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ACEPTAR_DESPACHAR_PEDIDO));
        ExtentReportUtil.INSTANCE.stepPass(driver, "Despachar pedido");
        generateWord.sendText("Despachar pedido");
        generateWord.addImageToWord(driver);
        driver.findElement(O_Corporate.BTN_ACEPTAR_DESPACHAR_PEDIDO).click();
        Thread.sleep(8000);
        ExtentReportUtil.INSTANCE.stepPass(driver, "Despacho motorizado LISTO");
        generateWord.sendText("Despacho motorizado LISTO");
        generateWord.addImageToWord(driver);

    }
    public void MaestroDePedidosCorporativos() throws Exception {

        driver.findElement(O_Corporate.LST_IR_A).click();
        MoveToElement(driver,O_Corporate.LNK_GEST_PED_CORP);
        driver.findElement(O_Corporate.LNK_MAESTRO_PEDIDO).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_ID_RESERVA3));
        ExtentReportUtil.INSTANCE.stepPass(driver, "Maestro de Pedido");
        generateWord.sendText("Maestro de Pedido");
        generateWord.addImageToWord(driver);
    }
    public void BuscamosElNumeroDePedidoC(String casoDePrueba) throws Throwable {
        // Write code here that turns the phrase above into concrete actions
        try {
            Thread.sleep(2000);
            int pedido1 = Integer.parseInt(casoDePrueba) - 1;
            String user1 = getData().get(pedido1).get(ExcelCorporativo.IDPEDIDO);
            driver.findElement(O_Corporate.TXT_INGRESAR_PEDIDO3).clear();
            driver.findElement(O_Corporate.TXT_INGRESAR_PEDIDO3).sendKeys(user1);
            driver.findElement(O_Corporate.TXT_INGRESAR_PEDIDO3).sendKeys(Keys.ENTER);
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("m6a7dfd2f_tdrow_[C:0]-c[R:0]")));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Ingresamos el número de pedido.");
            generateWord.sendText("Ingresamos el número de pedido.");
            generateWord.addImageToWord(driver);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }

    }
    public void VerificamosElEstadoDelPedidoCorporativos() throws Exception {

        driver.findElement(O_Corporate.LNK_PEDIDO).click();
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_ESTADO_PEDIDO));
        String estado;

        estado = driver.findElement(O_Corporate.TXT_ESTADO_PEDIDO).getText();
        if  (estado.equals("DESPACHADO")){
            ExtentReportUtil.INSTANCE.stepPass(driver, "Pedido despachado");
            generateWord.sendText("Pedido despachado");
            generateWord.addImageToWord(driver);

        }

    }
    public void EntregarPedidoCorporativos() throws Exception {
        try {

            driver.findElement(O_Corporate.BTN_ENTREGA_PEDIDO).click();
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ENTREGA));
            driver.findElement(O_Corporate.BTN_ENTREGA).click();
            Thread.sleep(1000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Entregar Pedido");
            generateWord.sendText("Entregar Pedido");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_SI_4).click();
            wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.BTN_ACEPTAR_MENS_SIST));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje del sistema");
            generateWord.sendText("Mensaje del sistema");
            generateWord.addImageToWord(driver);
            driver.findElement(O_Corporate.BTN_ACEPTAR_MENS_SIST).click();
            Thread.sleep(7000);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void ValidarEstadoDelPedidoCorporativos() throws Exception {
        WebDriverWait wait = new WebDriverWait(driver, 60);
        wait.until(ExpectedConditions.visibilityOfElementLocated(O_Corporate.TXT_ESTADO_PEDIDO));
        String estado;
        estado = driver.findElement(O_Corporate.TXT_ESTADO_PEDIDO).getText();
        if  (estado.equals("ENTREGADO")){
            ExtentReportUtil.INSTANCE.stepPass(driver, "Pedido entregado");
            generateWord.sendText("Pedido entregado");
            generateWord.addImageToWord(driver);
        }
    }


}
