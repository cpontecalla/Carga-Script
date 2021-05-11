package com.tsoft.bot.frontend.pages.pages;

import com.tsoft.bot.frontend.Base.BaseClass;
import com.tsoft.bot.frontend.helpers.Hook;
import com.tsoft.bot.frontend.pages.objects.ExcelResidencial;
import com.tsoft.bot.frontend.pages.objects.O_CargaMateriales;
import com.tsoft.bot.frontend.pages.objects.O_Residential;
import com.tsoft.bot.frontend.utility.ExcelReader;
import com.tsoft.bot.frontend.utility.ExtentReportUtil;
import com.tsoft.bot.frontend.utility.GenerateWord;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.sikuli.script.Region;
import org.sikuli.script.Screen;
import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.util.HashMap;
import java.util.List;
import java.util.Set;
import java.util.concurrent.ThreadLocalRandom;

public class P_Residential extends BaseClass {
    String fecha;
    String NUM_ENVIO;
    String user;
    public WebDriver driver;
    static GenerateWord generateWord = new GenerateWord();
    public P_Residential( WebDriver driver) {
        super(driver);
        this.driver = Hook.getDriver();
    }
    private List<HashMap<String, String>> getData() throws Throwable {
        return ExcelReader.data(ExcelResidencial.EXCEL_WEB, ExcelResidencial.ORDEN);
    }
        public void ingresoALaUrlDeWEBDELIVERY(String casoDePrueba) throws Throwable {
            try {
                int LoginWD = Integer.parseInt(casoDePrueba) - 1;
                String url = getData().get(LoginWD).get(ExcelResidencial.COLUMNA_URL);
                driver.get(url);
                stepPass(driver,"Se cargó correctamente la página");
                generateWord.sendText("Carga correcta de la página");
                generateWord.addImageToWord(driver);
                println("[LOG] Se cargó correctamente la página");
                generateWord.sendBreak();
            }catch (Exception e){
                ExcelReader.writeCellValue(ExcelResidencial.EXCEL_WEB, ExcelResidencial.ORDEN, 1, 19, "FAIL");
                ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
                generateWord.sendText("Tiempo de espera ha excedido");
                generateWord.addImageToWord(driver);
            }
        }

        public void ingresoElUsuarioDeWEBDELIVERY(String casoDePrueba) throws Throwable {

            try {
                int user = Integer.parseInt(casoDePrueba) - 1;
                String usuario = getData().get(user).get(ExcelResidencial.COLUMNA_USUARIO);
                wait(driver, O_CargaMateriales.TXT_USER,60);
                if (isDisplayed(driver, O_CargaMateriales.TXT_USER)){
                    clear(driver, O_CargaMateriales.TXT_USER);
                    sendKeys(driver, O_CargaMateriales.TXT_USER,usuario);
                }
                stepPass(driver,"Ingresamos el usuario");
                generateWord.sendText("Ingresamos el usuario");
                generateWord.addImageToWord(driver);
                println("[LOG] Ingresamos usuario");

            }catch (Exception e){
                ExcelReader.writeCellValue(ExcelResidencial.EXCEL_WEB, ExcelResidencial.ORDEN, 1, 19, "FAIL");
                ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
                generateWord.sendText("Tiempo de espera ha excedido");
                generateWord.addImageToWord(driver);
            }
        }
        public void laContraseñaDeWEBDELIVERY(String casoDePrueba) throws Throwable {
            try {
                int PASS = Integer.parseInt(casoDePrueba) - 1;
                wait(driver, O_CargaMateriales.TXT_PASSWORD,60);
                clear(driver, O_CargaMateriales.TXT_PASSWORD);
                String contra = getData().get(PASS).get(ExcelResidencial.COLUMNA_CONTRASENIA);
                sendKeys(driver, O_CargaMateriales.TXT_PASSWORD,contra);
                stepPass(driver,"Ingresamos la contraseña");
                generateWord.sendText("Ingresamos la contraseña");
                generateWord.addImageToWord(driver);
                println("[LOG] Ingresamos contraseña");
            }catch (Exception e){
                ExcelReader.writeCellValue(ExcelResidencial.EXCEL_WEB, ExcelResidencial.ORDEN, 1, 19, "FAIL");
                ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
                generateWord.sendText("Tiempo de espera ha excedido");
                generateWord.addImageToWord(driver);
            }
        }
        public void seDaClicEnElBotonLoginDeWEBDELIVERYIngresandoCorrectamente() throws Throwable {
            try {
                click(driver, O_CargaMateriales.BTN_LOGIN);
                sleep(2000);
                wait(driver,O_CargaMateriales.LNK_CREAR_PEDIDO,60);
                stepPass(driver,"Se ingresa correctamente a la pagina");
                generateWord.sendText("Se ingresa correctamente a la pagina");
                generateWord.addImageToWord(driver);
                println("[LOG] Logueo exitoso");
            }catch (Exception e){
                ExcelReader.writeCellValue(ExcelResidencial.EXCEL_WEB, ExcelResidencial.ORDEN, 1, 19, "FAIL");
                ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
                generateWord.sendText("Tiempo de espera ha excedido");
                generateWord.addImageToWord(driver);
            }
        }
    public void seDaClickEnElBotonIRAEnWEBDELIVERY(String arg0) throws Throwable {
        try {
            click(driver, O_CargaMateriales.LST_IR_A);
            ExtentReportUtil.INSTANCE.stepPass(driver, "IR A lista de pedidos");
        }catch (Exception e){
            ExcelReader.writeCellValue(ExcelResidencial.EXCEL_WEB, ExcelResidencial.ORDEN, 1, 19, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void seleccionarAjusteDeInventario() throws Throwable {
        try {
            Actions act = new Actions(driver);
            act.moveToElement(driver.findElement(O_CargaMateriales.LNK_GESTION_PEDIDOS)).build().perform();
            Actions act2 = new Actions(driver);
            act2.moveToElement(driver.findElement(O_CargaMateriales.LNK_GESTION_INVENTARIOS)).build().perform();
            click(driver,O_CargaMateriales.LNK_AJUSTE_INVENTARIO);
            sleep(2000);
            stepPass(driver,"Ajuste de inventario");
            generateWord.sendText("Ajuste de inventario");
            generateWord.addImageToWord(driver);
            println("[LOG] Ajuste de inventario");
        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void clickEnElBotonNuevoRegistro() throws Throwable {
        try {
            click(driver,O_CargaMateriales.BTN_NUEVO_REGISTRO);
            sleep(2000);
            stepPass(driver,"Seleccionamos nuevo registro");
            generateWord.sendText("Seleccionamos nuevo registro");
            generateWord.addImageToWord(driver);
            println("[LOG] Seleccionamos nuevo registro");
        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void seleccionamosElTipoABASTECIMIENTO(String arg0) throws Throwable {
        try {
            click(driver,O_CargaMateriales.BTN_TIPO);
            sleep(1000);
            click(driver,O_CargaMateriales.LNK_ABASTECIMIENTO);
            sleep(2000);
            String estado = driver.findElement(O_CargaMateriales.TXT_TIPO).getAttribute("value");
            if (estado.equals("ABASTECIMIENTO")){
                stepPass(driver,"Seleccionamos tipo: ABASTECIMIENTO");
                generateWord.sendText("Seleccionamos tipo: ABASTECIMIENTO");
                generateWord.addImageToWord(driver);
                println("[LOG] Seleccionamos tipo: ABASTECIMIENTO");
            }else {
                stepFail(driver,"No seleccionó ABASTECIMIENTO");
                generateWord.sendText("No seleccionó ABASTECIMIENTO");
                generateWord.addImageToWord(driver);
            }
        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void ingresamosUnComentario(String casoDePrueba) throws Throwable {

        try {
            clear(driver,O_CargaMateriales.TXT_COMENTARIO);
            sendKeys(driver,O_CargaMateriales.TXT_COMENTARIO,"PRUEBAS-QA");
            stepPass(driver,"Ingresamos comentario");
            generateWord.sendText("Ingresamos comentario");
            generateWord.addImageToWord(driver);
            println("[LOG] Ingresamos comentario: PRUEBAS-QA");
        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void ingresamosGuiaDeRemision(String casoDePrueba) throws Throwable {

        try {
            clear(driver,O_CargaMateriales.TXT_GUIA_REMISION);
            int random = ThreadLocalRandom.current().nextInt(10, 99);
            int random2 = ThreadLocalRandom.current().nextInt(10, 99);
            int random3 = ThreadLocalRandom.current().nextInt(10, 99);
            int random4 = ThreadLocalRandom.current().nextInt(1, 9);
            int random5 = ThreadLocalRandom.current().nextInt(1, 9);
            int random6 = ThreadLocalRandom.current().nextInt(1, 9);
            String numero = "12"+random6+random5+"-"+ random + random2 + random3+random4;
            sendKeys(driver,O_CargaMateriales.TXT_GUIA_REMISION,numero);
            stepPass(driver,"Ingresamos guia de remision");
            generateWord.sendText("Ingresamos guia de remision");
            generateWord.addImageToWord(driver);
            println("[LOG] Guia de remision ingresada: "+numero);
        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void ingresamosElArchivo() throws Throwable {
        try {
            click(driver,O_CargaMateriales.BTN_ADJUNTAR_ARCHIVOS);
            Actions act = new Actions(driver);
            act.moveToElement(driver.findElement(O_CargaMateriales.LNK_ADJUNTAR_NUEVO_ARCHIVO)).build().perform();
            click(driver,O_CargaMateriales.LNK_ARCHIVO_NUEVO);
            sleep(2000);
            stepPass(driver,"Ingresamos nuevo archivo");
            generateWord.sendText("Ingresamos nuevo archivo");
            generateWord.addImageToWord(driver);
            println("[LOG] Cargando archivo CSV");
            driver.switchTo().frame(0);
            driver.findElement(O_CargaMateriales.BTN_SELECCIONAR_ARCHIVO).click();
            Thread.sleep(1000);
            Robot robot = new Robot();
            String ruta = "D:\\ASIGNACIONES\\AsignacionSeries_3.csv";
            String text = ruta;
            StringSelection stringSelection = new StringSelection(text);
            Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
            clipboard.setContents(stringSelection, stringSelection);
            robot.keyPress(KeyEvent.VK_CONTROL);
            robot.keyPress(KeyEvent.VK_V);
            robot.keyRelease(KeyEvent.VK_V);
            robot.keyRelease(KeyEvent.VK_CONTROL);
            sleep(2000);
            robot.keyPress(KeyEvent.VK_ENTER);
            sleep(4000);
            Screen screen = new Screen();
            screen.wait(O_CargaMateriales.BTN_ACEPTAR_ARCHIVO);
            Region valBtn = screen.find(O_CargaMateriales.BTN_ACEPTAR_ARCHIVO).highlight(1,"green");
            screen.click(O_CargaMateriales.BTN_ACEPTAR_ARCHIVO);
            println("[LOG] Archivo CSV cargado");
        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void clickEnEjecutarAjusteYAceptarMensaje() throws Throwable {
        try {
            sleep(5000);
            click(driver,O_CargaMateriales.BTN_EJECUTAR_AJUSTE);
            stepPass(driver,"Ejecutar ajuste");
            generateWord.sendText("Ejecutar ajuste");
            generateWord.addImageToWord(driver);
            println("[LOG] Ejecutamos ajuste");
            sleep(2000);
            click(driver,O_CargaMateriales.BTN_ACEPTAR_AJUSTE);
            sleep(7000);
            stepPass(driver,"Mensaje del sistema");
            generateWord.sendText("Mensaje del sistema");
            generateWord.addImageToWord(driver);
            wait(driver,O_CargaMateriales.BTN_ACEPTAR_SISTEMA,60);
            String text;
            text = driver.findElement(O_CargaMateriales.TXT_IMAGEN).getText();
            text = text.substring(13);
            if (text.equals("Error en el proceso, verificar el campo de error")){
               stepFail(driver,"Error al cargar materiales");
                generateWord.sendText("Error al cargar materiales");
                generateWord.addImageToWord(driver);
                println("[LOG] Error en la carga de materiales");
            }
            if (text.equals("Ajuste ejecutado con éxito") || text.equals(" Ajuste ejecutado con éxito") ){
                stepPass(driver,"Carga de materiales exitoso");
                generateWord.sendText("Carga de materiales exitoso");
                generateWord.addImageToWord(driver);
                println("[LOG] Materiales cargados correctamente");
            }
            click(driver,O_CargaMateriales.BTN_ACEPTAR_SISTEMA);
            sleep(4000);

        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void validarQueLosArchivosHayanCargado() throws Throwable {
        try {

            String filas;
            filas = driver.findElement(O_CargaMateriales.TABLE).getAttribute("displayrows");
            println("[LOG] ------ Detalle ------");
            int num = Integer.parseInt(filas);
            for (int  i =0; (i<num); i++){
                String valor = driver.findElement(By.id("me7037f0c_tdrow_[C:10]-c[R:"+i+"]")).getText();
                String material = driver.findElement(By.id("me7037f0c_tdrow_[C:7]-c[R:"+i+"]")).getText();
                println(material + "  ->  " + valor);
            }
            stepPass(driver,"Detalle de carga");
            generateWord.sendText("Detalle de carga");
            generateWord.addImageToWord(driver);
        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void seleccionamosAUDITORIADEPEDIDO(String arg0) throws Throwable {
        try {
            wait(driver, O_Residential.LST_IrA,60);
            click(driver, O_Residential.LST_IrA);
            //action
            Actions act = new Actions(driver);
            act.moveToElement(driver.findElement(O_Residential.LST_GestPedido)).build().perform();
            click(driver, O_Residential.LST_AuditoriaPedido);
            wait(driver, O_Residential.TXT_IdOrden, 60);
            stepPass(driver,"Auditoria de pedido");
            generateWord.sendText("Auditoria de pedido");
            generateWord.addImageToWord(driver);
            sleep(2000);
            println("Se ingresa correctamente a la pagina");
        }catch (Exception e){
            ExcelReader.writeCellValue(ExcelResidencial.EXCEL_WEB, ExcelResidencial.ORDEN, 1, 19, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void buscamosIDDEORDEN(String casoDePrueba) throws Throwable {
        try {
            int orden = Integer.parseInt(casoDePrueba) - 1;
            clear(driver, O_Residential.TXT_IdOrden);
            String user = getData().get(orden).get(ExcelResidencial.ID_ORDEN);
            sendKeys(driver,O_Residential.TXT_IdOrden,user);
            sendKeysRobot(driver, O_Residential.TXT_IdOrden, Keys.ENTER);
            sleep(3000);
            wait(driver, O_Residential.BTN_Pedido, 30);
            Boolean isPresent = driver.findElements(O_Residential.BTN_Pedido).size() > 0;
            System.out.println(isPresent);
            if (isPresent.equals("false")) {
                ExtentReportUtil.INSTANCE.stepFail(driver, "Orden no encontrada");
                generateWord.sendText("Orden no encontrada");
                generateWord.addImageToWord(driver);
                driver.quit();
            }
            stepPass(driver,"Id de orden encontrada");
            generateWord.sendText("Id de orden encontrada");
            generateWord.addImageToWord(driver);
            println("Se buscó el Id de la orden");
        }catch (Exception e){
            ExcelReader.writeCellValue(ExcelResidencial.EXCEL_WEB, ExcelResidencial.ORDEN, 1, 19, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void seleccionamosPEDIDO() throws Throwable {
        try {
            wait(driver, O_Residential.BTN_Pedido, 60);
            click(driver, O_Residential.BTN_Pedido);
            wait(driver, O_Residential.TXT_EstadoAgendado, 60);
            stepPass(driver,"Información del pedido");
            generateWord.sendText("Información del pedido");
            generateWord.addImageToWord(driver);
            println("Información del pedido");
        }catch (Exception e){
            ExcelReader.writeCellValue(ExcelResidencial.EXCEL_WEB, ExcelResidencial.ORDEN, 1, 19, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void agendamosELPEDIDO(String arg0) throws Throwable {
        try{
            click(driver, O_Residential.BTN_AgendarPedido);
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.or(ExpectedConditions.presenceOfElementLocated(O_Residential.LNK_FechaPedido), ExpectedConditions.presenceOfElementLocated(O_Residential.MENSAJE)));
            if(driver.findElement(O_Residential.MENSAJE).isDisplayed()){
                String text;
                text = driver.findElement(O_Residential.MENSAJE).getText();
                println(text);
                if (text.contains("La integración con el servicio Retrieve Billing Info cayó en error.")){
                    stepFail(driver,"Mensaje de error: "+text);
                    generateWord.sendText("Mensaje de error: "+text);
                    generateWord.addImageToWord(driver);
                    println("Mensaje de error: "+text);
                    driver.quit();
                }
            }
            stepPass(driver,"Agendar Pedido");
            generateWord.sendText("Agendar Pedido");
            generateWord.addImageToWord(driver);
            println("Agendar Pedido");
        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);

        }
    }

    public void seleccionamosFECHADEPEDIDO(String arg0) throws Throwable {
        try{
            click(driver, O_Residential.LNK_FechaPedido);
            click(driver, O_Residential.BTN_Buscar);
            wait(driver, O_Residential.TXT_EstadoAgendado, 60);
            stepPass(driver,"Seleccionamos fecha de pedido");
            generateWord.sendText("Seleccionamos fecha de pedido");
            generateWord.addImageToWord(driver);
            println("Seleccionamos fecha de pedido");
        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    public void validamosCAMBIODEESTADODELPEDIDOAGENDADO(String arg0) throws Throwable {
        try {
            sleep(3000);
            String estado = driver.findElement(O_Residential.TXT_EstadoAgendado).getAttribute("value");
            if (estado.equals("AGENDADO")) {
                stepPass(driver,"Estado de pedido: " + estado);
                generateWord.sendText("Estado de pedido: " + estado);
                generateWord.addImageToWord(driver);
                println("Estado de pedido: " + estado);
            } else {
                stepPass(driver, "Estado de pedido: " + estado + " -- es incorrecto--");
                generateWord.sendText("Estado de pedido: " + estado + " -- es incorrecto--");
                generateWord.addImageToWord(driver);
                println("Estado de pedido: " + estado + " -- es incorrecto--");
                driver.quit();
            }
        }catch (Exception e){
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    public void seleccionamosASIGNACIONDESERIES(String arg0) throws Throwable {
        try {
            click(driver, O_Residential.LST_Menu);
            Actions act = new Actions(driver);
            act.moveToElement(driver.findElement(O_Residential.LNK_Gestion_Pedido)).build().perform();
            act.moveToElement(driver.findElement(O_Residential.LNK_Preparacion_Pedido)).build().perform();
            click(driver, O_Residential.LNK_Asignacion_Serie);
            wait(driver, O_Residential.TXT_IdOrden2, 60);
            stepPass(driver,"Asignación de series");
            generateWord.sendText("Asignación de series");
            generateWord.addImageToWord(driver);
            println("Asignación de series");
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void buscamosELORDERID(String casoDePrueba) throws Throwable {
        try {
            int orden = Integer.parseInt(casoDePrueba) - 1;
            clear(driver, O_Residential.TXT_IdOrden2);
            String user = getData().get(orden).get(ExcelResidencial.ID_ORDEN);
            sendKeys(driver, O_Residential.TXT_IdOrden2, user);
            sendKeysRobot(driver, O_Residential.TXT_IdOrden2,Keys.ENTER);
            wait(driver, O_Residential.TXT_ORDEN, 60);

            String f = driver.findElement(O_Residential.TXT_ORDEN).getText();
            while (!f.equals(user)) {
                Thread.sleep(1000);
                String g;
                g = driver.findElement(O_Residential.TXT_VACIO2).getText();
                if (g.equals("0 - 0 de 0")) {
                    stepPass(driver, "ID Reserva no encontrado");
                    generateWord.sendText("ID Reserva no encontrado");
                    generateWord.addImageToWord(driver);
                    println("ID Reserva no encontrado");
                    driver.quit();
                } else {
                    f = driver.findElement(O_Residential.TXT_ORDEN).getText();
                }
            }
            stepPass(driver,"Orden encontrada");
            generateWord.sendText("Orden encontrada");
            generateWord.addImageToWord(driver);
            println("Orden encontrada");
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void ingresamosMATERIALES(String casoDePrueba) throws Throwable {
        try {
            int CantFilas = Integer.parseInt(casoDePrueba) - 1;
            String filas;
            filas = driver.findElement(O_Residential.TABLA).getAttribute("displayrows");
            int num = Integer.parseInt(filas);
            int s = CantFilas;
            int f = CantFilas;
            for (int i = 0; (i < num); i++) {
                String TipoMaterial = driver.findElement(By.id("md6723283_tdrow_[C:6]-c[R:" + i + "]")).getText();
                if (TipoMaterial.equals("IMEI")) {
                    String imei = getData().get(s).get(ExcelResidencial.IMEI);
                    clear(driver, By.id("md6723283_tdrow_[C:9]_txt-tb[R:" + i + "]"));
                    sendKeys(driver,By.id("md6723283_tdrow_[C:9]_txt-tb[R:" + i + "]"),imei);
                    sleep(1500);
                    s++;
                }
                if (TipoMaterial.equals("ICCID")) {
                    String sim = getData().get(f).get(ExcelResidencial.SIMCARD);
                    clear(driver, By.id("md6723283_tdrow_[C:9]_txt-tb[R:" + i + "]"));
                    sendKeys(driver,By.id("md6723283_tdrow_[C:9]_txt-tb[R:" + i + "]"),sim);
                    sleep(1500);
                    f++;
                }
                if (i == num - 1) {
                    ExtentReportUtil.INSTANCE.stepFail(driver, "Materiales ingresados");
                    generateWord.sendText("Materiales ingresados");
                    generateWord.addImageToWord(driver);
                    break;
                }
            }
            sleep(2000);


            /*
            int Tipo_Trans = Integer.parseInt(casoDePrueba) - 1;
            String TipoTrans = getData().get(Tipo_Trans).get(TIPO_TRANSACCION);
            String TipoAlta = getData().get(Tipo_Trans).get(TIPO_ALTA);
            String Imei = getData().get(Tipo_Trans).get(IMEI);
            String SimCard = getData().get(Tipo_Trans).get(SIMCARD);
            if (TipoTrans.equals("ALTA")){
                if (TipoAlta.equals("EQUIPO + SIM")){
                    driver.findElement(TXT_Imei).clear();
                    driver.findElement(TXT_Imei).sendKeys(Imei);
                    driver.findElement(TXT_SimCard).clear();
                    driver.findElement(TXT_SimCard).sendKeys(SimCard);
                }
                if(TipoAlta.equals("SOLO SIM")){
                    driver.findElement(TXT_Imei).clear();
                    driver.findElement(TXT_Imei).sendKeys(SimCard);
                }
            }
            if (TipoTrans.equals("CAMBIO SIMCARD")){
                driver.findElement(TXT_Imei).clear();
                driver.findElement(TXT_Imei).sendKeys(SimCard);
            }
            if (TipoTrans.equals("CAMBIO EQUIPO")){
                driver.findElement(TXT_Imei).clear();
                driver.findElement(TXT_Imei).sendKeys(Imei);
            }

            ExtentReportUtil.INSTANCE.stepPass(driver, "Ingresamos materiales");
            generateWord.sendText("Ingresamos materiales");
            generateWord.addImageToWord(driver);*/
        } catch (Exception e) {
            ExcelReader.writeCellValue(ExcelResidencial.EXCEL_WEB, ExcelResidencial.ORDEN, 1, 19, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);

        }
    }
    public void validamosSERIES() throws Throwable {
        try{
            click(driver, O_Residential.BTN_Validar_Serie);
            wait(driver, O_Residential.BTN_Aceptar_1, 60);
            stepPass(driver,"Validar serie");
            generateWord.sendText("Validar serie");
            generateWord.addImageToWord(driver);
            println("Validar serie");

            click(driver, O_Residential.BTN_Aceptar_1);
            wait(driver, O_Residential.BTN_Aceptar_2, 60);
            stepPass(driver,"Aceptamos mensaje para refrescar la pagina");
            generateWord.sendText("Aceptamos mensaje para refrescar la pagina");
            generateWord.addImageToWord(driver);
            println("Aceptamos mensaje para refrescar la pagina");
            click(driver, O_Residential.BTN_Aceptar_2);
            sleep(3000);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void verificamosESTADODEVALIDACIONDESERIES(String casoDePrueba) throws Throwable {
        try {
            String filas;
            int error = 0;
            int bien = 0;
            filas = driver.findElement(O_Residential.TABLA).getAttribute("displayrows");
            int num = Integer.parseInt(filas);
            for (int i = 0; (i < num); i++) {
                String estado = driver.findElement(By.id("md6723283_tdrow_[C:10]-c[R:" + i + "]")).getText();
                if (estado.equals("ERROR") || estado.equals("PENDIENTE")) {
                    error++;
                    String valor = driver.findElement(By.id("md6723283_tdrow_[C:9]_txt-tb[R:" + i + "]")).getAttribute("value");
                    System.out.println(valor + " --> " + estado);
                }
                if (estado.equals("VALIDADO")) {
                    bien++;
                    String valor = driver.findElement(By.id("md6723283_tdrow_[C:9]_txt-tb[R:" + i + "]")).getAttribute("value");
                    System.out.println(valor + " --> VALIDADO");
                }
                if (i == num - 1) {
                    if (error > 0) {
                        System.out.println("Se obtuvo un total de: " + error + " materiales con error");
                        System.out.println("Se obtuvo un total de: " + bien + " materiales validados");
                        ExtentReportUtil.INSTANCE.stepFail(driver, "Materiales no validados");
                        generateWord.sendText("Materiales no validados");
                        generateWord.addImageToWord(driver);
                        driver.quit();
                        break;
                    } else {
                        System.out.println("Se obtuvo un total de: " + bien + " materiales validados");
                        ExtentReportUtil.INSTANCE.stepFail(driver, "Materiales validados");
                        generateWord.sendText("Materiales validados");
                        generateWord.addImageToWord(driver);
                        break;
                    }
                }
            }

            /*
            int Tipo_Trans = Integer.parseInt(casoDePrueba) - 1;
            String TipoTrans = getData().get(Tipo_Trans).get(TIPO_TRANSACCION);
            String TipoAlta = getData().get(Tipo_Trans).get(TIPO_ALTA);
            if (TipoTrans.equals("ALTA")){
                if (TipoAlta.equals("EQUIPO + SIM")){
                    String IMEI = driver.findElement(TXT_Estado_Imei).getText();
                    while(IMEI.equals("PENDIENTE")){
                        Thread.sleep(1000);
                        IMEI = driver.findElement(TXT_Estado_Imei).getText();
                    }
                    String SIMCARD = driver.findElement(TXT_Estado_SimCard).getText();
                    if (IMEI.equals("VALIDADO") && SIMCARD.equals("VALIDADO")){
                        ExtentReportUtil.INSTANCE.stepPass(driver, "Materiales validados");
                        generateWord.sendText("Materiales validados");
                        generateWord.addImageToWord(driver);

                    }else {
                        ExtentReportUtil.INSTANCE.stepFail(driver, "Error al validar las series");
                        generateWord.sendText("Error al validar las series");
                        generateWord.addImageToWord(driver);

                    }
                }
                if (TipoAlta.equals("SOLO SIM")){
                    String IMEI = driver.findElement(TXT_Estado_Imei).getText();
                    while(IMEI.equals("PENDIENTE")){
                        Thread.sleep(1000);
                        IMEI = driver.findElement(TXT_Estado_Imei).getText();
                    }
                    if (IMEI.equals("VALIDADO")){
                        ExtentReportUtil.INSTANCE.stepPass(driver, "Materiales validados");
                        generateWord.sendText("Materiales validados");
                        generateWord.addImageToWord(driver);

                    }else {
                        ExtentReportUtil.INSTANCE.stepFail(driver, "Error al validar las series");
                        generateWord.sendText("Error al validar las series");
                        generateWord.addImageToWord(driver);

                    }
                }
            }*/

        } catch (Exception e) {
            ExcelReader.writeCellValue(ExcelResidencial.EXCEL_WEB, ExcelResidencial.ORDEN, 1, 19, "FAIL");
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);

        }
    }
    public void seleccionamosIMPRESIONDEDOCUMENTOS(String arg0) throws Throwable {
        try {
            click(driver, O_Residential.LST_Menu);
            Actions act = new Actions(driver);
            act.moveToElement(driver.findElement(O_Residential.LNK_Gestion_Pedido)).build().perform();
            act.moveToElement(driver.findElement(O_Residential.LNK_Preparacion_Pedido)).build().perform();
            click(driver, O_Residential.LNK_Impresion_Documento);
            wait(driver, O_Residential.LST_Seleccionar_Accion, 60);
            stepPass(driver, "Impresión de documentos");
            generateWord.sendText("Impresión de documentos");
            generateWord.addImageToWord(driver);
            println("Impresión de documentos");
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void buscamosELORDER_ID(String casoDePrueba) throws Throwable {
        try {
            int orden = Integer.parseInt(casoDePrueba) - 1;
            wait(driver, O_Residential.TXT_IdOrden3, 60);
            clear(driver, O_Residential.TXT_IdOrden3);
            String user = getData().get(orden).get(ExcelResidencial.ID_ORDEN);
            sendKeys(driver, O_Residential.TXT_IdOrden3,user);
            sendKeysRobot(driver, O_Residential.TXT_IdOrden3, Keys.ENTER);
            wait(driver, O_Residential.TXT_RESULTADO, 60);
            //
            if (driver.findElement(O_Residential.TXT_RESULTADO).isDisplayed()) {
                sleep(110);
            } else {
                ExtentReportUtil.INSTANCE.stepPass(driver, "ID Reserva no encontrado");
                generateWord.sendText("ID Reserva no encontrado");
                generateWord.addImageToWord(driver);
                driver.quit();
            }
            String f = driver.findElement(O_Residential.TXT_RESULTADO).getText();
            while (!f.equals(user)) {
                sleep(1000);
                String g;
                g = driver.findElement(O_Residential.TXT_VACIO).getText();
                if (g.equals("0 - 0 de 0")) {
                    ExtentReportUtil.INSTANCE.stepPass(driver, "ID Reserva no encontrado");
                    generateWord.sendText("ID Reserva no encontrado");
                    generateWord.addImageToWord(driver);
                    driver.quit();
                } else {
                    f = driver.findElement(O_Residential.TXT_ORDEN).getText();

                }
            }
            ExtentReportUtil.INSTANCE.stepPass(driver, "Orden encontrada");
            generateWord.sendText("Orden encontrada");
            generateWord.addImageToWord(driver);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void seleccionamosEJECUTARINFORMES(String arg0) throws Throwable {
        try{
            click(driver, O_Residential.LST_Seleccionar_Accion);
            sleep(1000);
            click(driver, O_Residential.LNK_Ejecutar_Informes);
            wait(driver, O_Residential.LNK_Guia_Remision,60);
            stepPass(driver, "Informes y programaciones");
            generateWord.sendText("Informes y programaciones");
            generateWord.addImageToWord(driver);
            println("Informes y programaciones");
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void impresionDEGUIADEREMISION() throws Throwable {
        try {
            click(driver, O_Residential.LNK_Guia_Remision);
            wait(driver, O_Residential.TXT_Numero_Serie, 60);
            clear(driver, O_Residential.TXT_Numero_Serie);
            sendKeys(driver, O_Residential.TXT_Numero_Serie,"12345");
            sleep(1000);
            clear(driver, O_Residential.TXT_Correlativo_Guia);
            sendKeys(driver, O_Residential.TXT_Correlativo_Guia,"12345");
            sleep(150);
            stepPass(driver, "Impresión de guia de remision");
            generateWord.sendText("Impresión de guia de remision");
            generateWord.addImageToWord(driver);
            click(driver, O_Residential.BTN_Enviar);
            sleep(1000);
            String f = driver.findElement(O_Residential.MENSAJE).getText();
            if (f.contains("Correlativo de la Guia de Remision es un campo necesario.")) {
                click(driver, O_Residential.BTN_Aceptar_2);
                sleep(1000);
                clear(driver, O_Residential.TXT_Numero_Serie);
                sendKeys(driver, O_Residential.TXT_Numero_Serie,"12345");
                sleep(100);
                clear(driver, O_Residential.TXT_Correlativo_Guia);
                sendKeys(driver, O_Residential.TXT_Correlativo_Guia,"12345");
                sleep(100);
                stepPass(driver, "Impresión de guia de remision");
                generateWord.sendText("Impresión de guia de remision");
                generateWord.addImageToWord(driver);
                click(driver, O_Residential.BTN_Enviar);
                sleep(1000);

                /*String g = driver.findElement(MENSAJE).getText();
                if (g.contains("Correlativo de la Guia de Remision es un campo necesario.")) {
                    ExtentReportUtil.INSTANCE.stepFail(driver, "Mensaje recurrente");
                    generateWord.sendText("Mensaje recurrente");
                    generateWord.addImageToWord(driver);
                    driver.quit();
                }*/
            }
            sleep(17000);
            String parentWindow = driver.getWindowHandle();
            Set<String> handles = driver.getWindowHandles();
            for (String windowHandle : handles) {
                if (!windowHandle.equals(parentWindow)) {
                    //WebDriverWait wait2 = new WebDriverWait(driver.switchTo().window(windowHandle), 60);
                    //wait2.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"__bookmark_1\"]/tbody/tr[2]/td[2]")));
                    ExtentReportUtil.INSTANCE.stepPass(driver.switchTo().window(windowHandle), "Informe guia de remision");
                    generateWord.sendText("Informe guia de remision");
                    generateWord.addImageToWord(driver.switchTo().window(windowHandle));
                    driver.switchTo().window(windowHandle).close();
                    driver.switchTo().window(parentWindow);
                    break;
                }
            }
            /*ExtentReportUtil.INSTANCE.stepPass(driver, "Descarga completa");
            generateWord.sendText("Descarga completa");
            generateWord.addImageToWord(driver);
            Thread.sleep(5000);*/


        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void impresionDEETIQUETA(String arg0) throws Throwable {
        try {
            wait(driver, O_Residential.LNK_Etiqueta, 60);
            click(driver, O_Residential.LNK_Etiqueta);
            wait(driver, O_Residential.BTN_Enviar_Etiq, 60);
            stepPass(driver, "Impresión de etiqueta");
            generateWord.sendText("Impresión de etiqueta");
            generateWord.addImageToWord(driver);
            click(driver, O_Residential.BTN_Enviar_Etiq);
            sleep(5000);
            String parentWindow = driver.getWindowHandle();
            Set<String> handles = driver.getWindowHandles();
            for (String windowHandle : handles) {
                if (!windowHandle.equals(parentWindow)) {
                    WebDriverWait wait2 = new WebDriverWait(driver.switchTo().window(windowHandle), 60);
                    wait2.until(ExpectedConditions.visibilityOfElementLocated(O_Residential.TXT_NOMBRE));
                    ExtentReportUtil.INSTANCE.stepPass(driver.switchTo().window(windowHandle), "Informe impresión de etiqueta");
                    generateWord.sendText("Informe impresión de etiqueta");
                    generateWord.addImageToWord(driver.switchTo().window(windowHandle));
                    driver.switchTo().window(windowHandle).close();
                    driver.switchTo().window(parentWindow);
                    break;
                }
            }
            sleep(2000);
            click(driver, O_Residential.BTN_Cancelar);
            sleep(2000);

            /*ExtentReportUtil.INSTANCE.stepPass(driver, "Descarga completa");
            generateWord.sendText("Descarga completa");
            generateWord.addImageToWord(driver);
            Thread.sleep(5000);*/

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void verificamosQUEELESTADODELAORDENSEAELCORRECTOREALIZADO() throws Exception {
        try {
            sleep(2000);
            String estado1 = driver.findElement(O_Residential.TBC_IMPRIMIR_FACT).getText();
            String estado2 = driver.findElement(O_Residential.TBC_IMPRIMIR_GUIA).getText();
            String estado3 = driver.findElement(O_Residential.TBC_IMPRIMIR_ETIQ).getText();

            if (estado1.equals("REALIZADO") || estado2.equals("REALIZADO") || estado3.equals("REALIZADO")) {
                ExtentReportUtil.INSTANCE.stepPass(driver, "Estado Correcto: REALIZADO");
                generateWord.sendText("Estado Correcto: REALIZADO");
                generateWord.addImageToWord(driver);
            } else {
                ExtentReportUtil.INSTANCE.stepFail(driver, "Estado Incorrecto");
                generateWord.sendText("Estado Incorrecto");
                generateWord.addImageToWord(driver);
            }

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void seleccionamosDESPACHODEHUB(String arg0) throws Throwable {
        try{
            click(driver, O_Residential.LST_IrA);
            Actions act = new Actions(driver);
            act.moveToElement(driver.findElement(O_Residential.LST_GestPedido)).build().perform();
            click(driver, O_Residential.LNK_DESPACHO_HUB);
            wait(driver, O_Residential.TXT_IdOrden3, 60);
            stepPass(driver, "Despacho de HUB");
            generateWord.sendText("Despacho de HUB");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void asignamosMASTERBOX(String arg0) throws Throwable {
        try{
            click(driver, O_Residential.BTN_MASTER_BOX);
            wait(driver, O_Residential.BTN_ACEPTAR_MB, 60);
            stepPass(driver, "Asignar Master Box");
            generateWord.sendText("Asignar Master Box");
            generateWord.addImageToWord(driver);
            click(driver, O_Residential.BTN_ACEPTAR_MB);
            wait(driver, O_Residential.BTN_ACEPTAR_MS, 60);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Mensaje del sistema");
            generateWord.sendText("Mensaje del sistema");
            generateWord.addImageToWord(driver);
            click(driver, O_Residential.BTN_ACEPTAR_MS);
            wait(driver, O_Residential.TBC_ESTADO,60);
            sleep(2000);
        }catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void verificamosLAASIGNACIONCORRECTADELCODIGO() throws Throwable {
        try{
            String f = driver.findElement(By.id("m6a7dfd2f_tdrow_[C:9]-c[R:0]")).getText();
            if (f.contains("Error 1002 XML Body not well-formed")) {
                stepFail(driver,"Mensaje de error: " + f);
                generateWord.sendText("Mensaje de error: " + f);
                generateWord.addImageToWord(driver);
                driver.quit();
            }
            if (isDisplayed(driver,O_Residential.TBC_ESTADO)) {
                stepPass(driver,"Se asignó código de Master Hub");
                generateWord.sendText("Se asignó código de Master Hub");
                generateWord.addImageToWord(driver);
            } else {
                stepFail(driver,"No se asignó código de Master Hub");
                generateWord.sendText("No se asignó código de Master Hub");
                generateWord.addImageToWord(driver);
                driver.quit();
            }
        }catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void despachoDEPEDIDO() throws Throwable {
        try{
            String CodMastHub = driver.findElement(O_Residential.TBC_ESTADO).getText();
            click(driver, O_Residential.BTN_DESPACHAR_PEDIDO);
            wait(driver, O_Residential.TXT_COD_MASTBOX, 60);
            clear(driver, O_Residential.TXT_COD_MASTBOX);
            sendKeys(driver, O_Residential.TXT_COD_MASTBOX, CodMastHub);
            sleep(100);
            clear(driver, O_Residential.TXT_HUB_DISTRIB);
            sendKeys(driver, O_Residential.TXT_HUB_DISTRIB,"HUBLIMA");
            sleep(100);
            stepPass(driver, "Despachar pedido");
            generateWord.sendText("Despachar pedido");
            generateWord.addImageToWord(driver);
            click(driver, O_Residential.BTN_DESPACHAR_PED);
            wait(driver, O_Residential.BTN_ACEPTAR_MS, 60);
            stepPass(driver,"Mensaje de sistema");
            generateWord.sendText("Mensaje de sistema");
            generateWord.addImageToWord(driver);
            click(driver, O_Residential.BTN_ACEPTAR_MS);
        }catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void recepcionarPEDIDOS() throws Throwable {
        try{
            click(driver, O_Residential.LST_IrA);
            Actions act = new Actions(driver);
            act.moveToElement(driver.findElement(O_Residential.LST_GestPedido)).build().perform();
            click(driver, O_Residential.LST_Recepcio_Pedidos);
            sleep(2000);
            wait(driver, O_Residential.TXT_IdOrden3, 60);
            stepPass(driver,"Recepción de pedidos");
            generateWord.sendText("Recepción de pedidos");
            generateWord.addImageToWord(driver);
        }catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void recepcionarPEDIDORESI() throws Throwable {
        try{
            wait(driver, O_Residential.TXT_COD_MASTER_BOX, 60);
            String Cod_Master_Box2 = driver.findElement(O_Residential.TXT_COD_MASTER_BOX).getText();
            println(Cod_Master_Box2);
            sleep(100);
            click(driver, O_Residential.BTN_RECEP_PEDIDO);
            println(Cod_Master_Box2);
            wait(driver, O_Residential.TXT_MASTERBOX2, 60);
            clear(driver, O_Residential.TXT_MASTERBOX2);
            sendKeys(driver, O_Residential.TXT_MASTERBOX2,Cod_Master_Box2);
            println(Cod_Master_Box2);
                /*SimpleDateFormat df = new SimpleDateFormat("dd/MM/YYYY");
                Date dt = new Date();
                Calendar cl = Calendar.getInstance();
                cl.setTime(dt);;
                cl.add(Calendar.DAY_OF_MONTH, 0);
                dt=cl.getTime();
                String str = df.format(dt);*/
            stepPass(driver,"Recepcionar pedido");
            generateWord.sendText("Recepcionar pedido");
            generateWord.addImageToWord(driver);
            click(driver, O_Residential.BTN_Aceptar_1);
            wait(driver, O_Residential.BTN_Aceptar_2, 60);
            String f = driver.findElement(O_Residential.MENSAJE).getText();
            if (f.contains("La recepción se ha finalizado con éxito")) {
                stepPass(driver,"Mensaje del sistema");
                generateWord.sendText("Mensaje del sistema");
                generateWord.addImageToWord(driver);
                click(driver, O_Residential.BTN_Aceptar_2);
                sleep(2000);
            }
            if (f.contains("Ingrese un valor de masterbox")) {
                stepFail(driver,"Mensaje sistema");
                generateWord.sendText("Mensaje sistema");
                generateWord.addImageToWord(driver);
                click(driver,O_Residential.BTN_Aceptar_2);
                sleep(2000);
                click(driver,O_Residential.BTN_RECEP_PEDIDO);
                wait(driver, O_Residential.TXT_MASTERBOX2, 60);
                clear(driver, O_Residential.TXT_MASTERBOX2);
                sendKeys(driver, O_Residential.TXT_MASTERBOX2, Cod_Master_Box2);
                stepPass(driver,"Recepcionar pedido");
                generateWord.sendText("Recepcionar pedido");
                generateWord.addImageToWord(driver);
                click(driver, O_Residential.BTN_Aceptar_1);
                wait(driver, O_Residential.BTN_Aceptar_2, 60);
                String g = driver.findElement(O_Residential.MENSAJE).getText();
                if (g.contains("Ingrese un valor de masterbox")) {
                    stepFail(driver,"Mensaje recurrente");
                    generateWord.sendText("Mensaje recurrente");
                    generateWord.addImageToWord(driver);
                    driver.quit();
                }
            }
            if (f.contains("No hay pedidos tengan el masterbox ingresado")) {
                stepFail(driver,"El Master box ingresado es incorrecto");
                generateWord.sendText("El Master box ingresado es incorrecto");
                generateWord.addImageToWord(driver);
                driver.quit();
            }
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void ejecutarCARGADELOGICADERUTEO() throws Throwable {
        try{
            sleep(2000);
            click(driver, O_Residential.BTN_EJECUTAR_CARGA_RUTEO);
            wait(driver, O_Residential.BTN_SI6, 60);
            stepPass(driver,"Carga lógica de ruteo");
            generateWord.sendText("Carga lógica de ruteo");
            generateWord.addImageToWord(driver);
            click(driver, O_Residential.BTN_SI6);
            String f = driver.findElement(O_Residential.MENSAJE).getText();
            if (f.contains("No se han cargado Datos para procesar, verificar el archivo adjunto y cargarlo nuevamente")) {
                stepFail(driver,"Error al momento de cargar el archivo");
                generateWord.sendText("Error al momento de cargar el archivo");
                generateWord.addImageToWord(driver);
                println("Error al momento de cargar el archivo");
                driver.quit();
            }
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void guardarNUMERODEENVIO() throws Throwable {
        try{
            sleep(2000);
            click(driver, O_Residential.TXT_ORDEN_ENVIO);
            sleep(1000);
            String NUM_ENVIO = driver.findElement(O_Residential.TXT_NUMERO_ENVIO).getText();
            println(NUM_ENVIO);
            if (NUM_ENVIO.equals("")) {
                sleep(1000);
                stepFail(driver,"Número de envío no generado");
                generateWord.sendText("Número de envío no generado");
                generateWord.addImageToWord(driver);
                driver.quit();
            }
            sleep(1000);
            stepPass(driver,"Número de envío generado");
            generateWord.sendText("Número de envío generado");
            generateWord.addImageToWord(driver);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void despachoDePedidoDeEnvio() throws Throwable {
        try{
            click(driver, O_Residential.LST_IrA);
            Actions act = new Actions(driver);
            act.moveToElement(driver.findElement(O_Residential.LST_GestPedido)).build().perform();
            click(driver, O_Residential.LNK_DESPACHO_PEDIDO);
            wait(driver, O_Residential.TXT_ENVIO, 60);
            stepPass(driver,"Despacho a motorizado");
            generateWord.sendText("Despacho a motorizado");
            generateWord.addImageToWord(driver);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void buscamosNumeroDeEnvio() throws Throwable {
        try{
            sleep(2000);
            clear(driver, O_Residential.TXT_ENVIO);
            String NUM_ENVIO = driver.findElement(O_Residential.TXT_NUMERO_ENVIO).getText();
            sendKeys(driver, O_Residential.TXT_ENVIO,NUM_ENVIO);
            sendKeysRobot(driver, O_Residential.TXT_ENVIO, Keys.ENTER);
            wait(driver, O_Residential.LNK_NUM_ENVIO, 60);
            stepPass(driver,"Pedido encontrado");
            generateWord.sendText("Pedido encontrado");
            generateWord.addImageToWord(driver);
            click(driver, O_Residential.LNK_NUM_ENVIO);
            wait(driver, O_Residential.BTN_LUPA11, 60);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void buscamosMotorizado() throws Throwable {
        try{
            click(driver, O_Residential.BTN_LUPA11);
            wait(driver, O_Residential.TXT_MOTORIZADO, 60);
            sendKeys(driver, O_Residential.TXT_MOTORIZADO, "MESCALGAMI");
            sendKeysRobot(driver, O_Residential.TXT_MOTORIZADO, Keys.ENTER);
            sleep(3000);
            stepPass(driver,"Seleccionamos motorizado");
            generateWord.sendText("Seleccionamos motorizado");
            generateWord.addImageToWord(driver);
            click(driver, O_Residential.LNK_MOTORIZADO);
            sleep(2000);
            Boolean isPresent = driver.findElements(O_Residential.MENSAJE).size() > 0;
            if (isPresent.equals(true)){
                String f = driver.findElement(O_Residential.MENSAJE).getText();
                if (f.contains("El campo Motorizado es de sólo lectura.")) {
                    stepPass(driver,"Mensaje");
                    generateWord.sendText("Mensaje");
                    generateWord.addImageToWord(driver);
                    click(driver, O_Residential.BTN_Aceptar_2);
                    sleep(2000);
                    click(driver,By.id("m507211d4-pb"));
                    sleep(2000);
                }
            }
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void despachamosPedido() throws Throwable {
        try{
            sleep(1000);
            click(driver, O_Residential.BTN_PEDIDOS);
            wait(driver, O_Residential.BTN_CAMBIAR_ESTADO, 60);
            String filas;
            String d;
            filas = driver.findElement(O_Residential.TABLE).getAttribute("displayrows");
            int num = Integer.parseInt(filas);
            System.out.println(num);
            for (int i = 0; (i < num); i++) {
                String o = driver.findElement(By.id("m187f4d3c_tdrow_[C:0]_checkbox-cb[R:" + i + "]_img")).getAttribute("src");
                System.out.println(o);
                if (o.contains("unchecked")) {
                    sleep(1000);
                }else{
                    //driver.findElement(By.id("m187f4d3c_tdrow_[C:0]_checkbox-cb[R:" + i + "]_img")).click();
                    click(driver,By.id("m187f4d3c_tdrow_[C:0]_checkbox-cb[R:" + i + "]_img"));
                }
            }
            sleep(2000);
            for (int i = 0; (i < num); i++) {
                String orden = driver.findElement(By.id("m187f4d3c_tdrow_[C:3]-c[R:" + i + "]")).getText();
                System.out.println(orden);
                System.out.println(user);
                sleep(1000);
                if (orden.equals(user)) {
                    sleep(1000);
                    String z = driver.findElement(By.id("m187f4d3c_tdrow_[C:0]_checkbox-cb[R:" + i + "]_img")).getAttribute("src");
                    System.out.println(z);
                    if (z.contains("unchecked")) {
                        System.out.println(z);
                        driver.findElement(By.id("m187f4d3c_tdrow_[C:0]_checkbox-cb[R:" + i + "]_img")).click();
                        break;
                    }
                }

                if (i == num - 1) {
                    ExtentReportUtil.INSTANCE.stepFail(driver, "No se encontro la orden");
                    generateWord.sendText("No se encontro la orden");
                    generateWord.addImageToWord(driver);
                    driver.quit();
                    break;
                }
            }

            sleep(1000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Pedido listo");
            generateWord.sendText("Pedido listo");
            generateWord.addImageToWord(driver);

            click(driver, O_Residential.BTN_DESPACHAR_PEDIDO2);
            wait(driver,O_Residential.BTN_ACEPTAR_DESPACHAR_PEDIDO, 60);
            //driver.findElement(BTN_DESPACHAR_PEDIDO2).click();
            //wait.until(ExpectedConditions.visibilityOfElementLocated(BTN_ACEPTAR_DESPACHAR_PEDIDO));
            ExtentReportUtil.INSTANCE.stepPass(driver, "Despachar pedido");
            generateWord.sendText("Despachar pedido");
            generateWord.addImageToWord(driver);
            click(driver, O_Residential.BTN_ACEPTAR_DESPACHAR_PEDIDO);
            //driver.findElement(BTN_ACEPTAR_DESPACHAR_PEDIDO).click();
            sleep(6000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Despacho de pedido listo");
            generateWord.sendText("Despacho de pedido listo");
            generateWord.addImageToWord(driver);

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
    public void validamosElEstadoDePedidoDESPACHADO() throws Exception {
        try {
            String text = driver.findElement(O_Residential.TXT_ESTADO_ENVIO).getAttribute("value");
            if (text.equals("DESPACHADO")){
                ExtentReportUtil.INSTANCE.stepPass(driver, "Estado de pedido: "+text);
                generateWord.sendText("Estado de pedido: "+text);
                generateWord.addImageToWord(driver);
            }else{
                ExtentReportUtil.INSTANCE.stepFail(driver, "Estado de pedido: "+text);
                generateWord.sendText("Estado de pedido: "+text);
                generateWord.addImageToWord(driver);
            }

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

}
