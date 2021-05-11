package com.tsoft.bot.frontend.steps.APP_USSD;

import com.tsoft.bot.frontend.helpers.Hook;
import com.tsoft.bot.frontend.utility.ExtentReportUtil;
import com.tsoft.bot.frontend.utility.GenerateWord;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import io.appium.java_client.AppiumDriver;
import io.appium.java_client.MobileElement;
import org.openqa.selenium.By;
import org.testng.Assert;

import java.util.List;

import static com.tsoft.bot.frontend.pageobject.APP_USSD.PageObject_USSD.*;

public class steps_USSD {

    private static GenerateWord generateWord = new GenerateWord();
    private AppiumDriver<MobileElement> driver;

    public steps_USSD() {
        this.driver = Hook.getDriver();
    }


    @Given("^Se ingresa a USSD mediante \"([^\"]*)\"$")
    public void seIngresaAUSSDMediante(String arg0) throws Throwable {
        try {
            driver.findElement(By.id(TXT_SEARCH)).clear();
            driver.findElement(By.id(TXT_SEARCH)).sendKeys("USSD");
            driver.findElement(By.id(SELECT_USSD)).click();
            driver.findElement(By.id(TAB_USSD)).click();
            driver.findElement(By.id(TXT_FIELD)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se inició correctamente USSD");
            generateWord.sendText("Se inició correctamente USSD");
            generateWord.addImageToWord(driver);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @When("^se ingresa a la opcion Consultas$")
    public void seIngresaALaOpcionConsultas() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(3000);
            driver.findElement(By.id(TXT_FIELD)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Consultas");
            generateWord.sendText("Se ingresó correctamente a la opción : Consultas");
            generateWord.addImageToWord(driver);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se ingresa a la opcion Mas$")
    public void seIngresaALaOpcionMas() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("5");
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(3000);
            driver.findElement(By.id(TXT_FIELD)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Mas");
            generateWord.sendText("Se ingresó correctamente a la opción : Mas");
            generateWord.addImageToWord(driver);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se ingresa a la opcion Tarifas$")
    public void seIngresaALaOpcionTarifas() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(5000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Tarifas");
            generateWord.sendText("Se ingresó correctamente a la opción : Tarifas");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se ingresa a la opcion Perdida/robo de tu equipo$")
    public void seIngresaALaOpcionPerdidaRoboDeTuEquipo() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("2");
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(5000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Perdida/robo de tu equipo");
            generateWord.sendText("Se ingresó correctamente a la opción : Perdida/robo de tu equipo");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se ingresa a la opcion Llamadas Internacionales$")
    public void seIngresaALaOpcionLlamadasInternacionales() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("3");
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(5000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Llamadas Internacionales");
            generateWord.sendText("Se ingresó correctamente a la opción : Llamadas Internacionales");
            generateWord.addImageToWord(driver);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se verifica para otros operadores$")
    public void seVerificaParaOtrosOperadores() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(5000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Llamadas Internacionales para otros Operadores");
            generateWord.sendText("Se ingresó correctamente a la opción : Llamadas Internacionales para otros Operadores");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @When("^se ingresa a la opcion Duplicar MB / FB Gratis$")
    public void seIngresaALaOpcionDuplicarMBFBGratis() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("2");
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(5000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Duplicar MB + FB GRATIS");
            generateWord.sendText("Se ingresó correctamente a la opción : Duplicar MB + FB GRATIS");
            generateWord.addImageToWord(driver);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se ingresa la opcion de duplicar (\\d+)MB$")
    public void seIngresaLaOpcionDeDuplicarMB(int arg0) throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(3000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Duplicar, Procesando solicitud");
            generateWord.sendText("Se ingresó correctamente a la opción : Duplicar, Procesando solicitud");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Given("^Se ingresa a USSD mediante \\*(\\d+)# Testing$")
    public void seIngresaAUSSDMedianteTesting(int arg0) throws Exception {
        try {
            driver.findElement(By.id(TXT_SEARCH)).clear();
            driver.findElement(By.id(TXT_SEARCH)).sendKeys("USSD");
            driver.findElement(By.id(SELECT_USSD)).click();
            driver.findElement(By.id(TAB_USSD)).click();
            driver.findElement(By.id(TXT_FIELD)).click();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se inició correctamente USSD");
            generateWord.sendText("Se inició correctamente USSD");
            generateWord.addImageToWord(driver);
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @When("^se ingresa a la opcion Consultar mi prepago$")
    public void seIngresaALaOpcionConsultarMiPrepago() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Consultar mi prepago");
            generateWord.sendText("Se ingresó correctamente a la opción : Consultar mi prepago");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(2000);
            driver.findElement(By.id(TXT_FIELD)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se ingresa la opcion Recibir SMS para descargar APP$")
    public void seIngresaLaOpcionRecibirSMSParaDescargarAPP() throws Exception {
        try {
            Thread.sleep(2000);
            driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Recibir SMS para descargar la app Mi Movistar");
            generateWord.sendText("Se ingresó correctamente a la opción : Recibir SMS para descargar la app Mi Movistar");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se verifica el envio de SMS$")
    public void seVerificaElEnvioDeSMS() throws Exception {
        try {
            Thread.sleep(5000);
            String textoToSearch = "Te enviaremos un SMS para descargar el app Mi Movistar";
            MobileElement element = driver.findElement(By.id(LBL_MESSAGE));
            String elText = element.getText();
            if(elText.equals(textoToSearch)){
                ExtentReportUtil.INSTANCE.stepPass(driver, elText);
                generateWord.sendText(elText);
                generateWord.addImageToWord(driver);
                driver.findElement(By.id(BTN_ENVIAR)).click();
            }else{
                ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : No se presentó el mensaje esperado");
                generateWord.sendText("Fallo el caso de prueba : No se presentó el mensaje esperado");
                generateWord.addImageToWord(driver);
            }

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @When("^se ingresa a la opcion Comprar Paquetes$")
    public void seIngresaALaOpcionComprarPaquetes() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("3");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Comprar Paquetes");
            generateWord.sendText("Se ingresó correctamente a la opción : Comprar Paquetes");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(2000);
            driver.findElement(By.id(TXT_FIELD)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se ingresa la opcion LLamadas\\+Apps\\+Datos\\+SMS$")
    public void seIngresaLaOpcionLLamadasAppsDatosSMS() throws Exception {
        try {
            Thread.sleep(2000);
            driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Elige tu paquete");
            generateWord.sendText("Se ingresó correctamente a la opción : Elige tu paquete");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se elige el monto a comprar \\((\\d+) soles\\)$")
    public void seEligeElMontoAComprarSoles(int arg0) throws Exception {
        try {
            Thread.sleep(2000);
            driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : 3 soles");
            generateWord.sendText("Se ingresó correctamente a la opción : 3 soles");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se verifica el envio de SMS al comprar (\\d+) soles$")
    public void seVerificaElEnvioDeSMSAlComprarSoles(int arg0) throws Exception {
        try {
            Thread.sleep(5000);
            String textoToSearch = "Gracias por comprar tu paquete de S/3 Min+Redes+100MB x2d, en unos instantes recibiras un SMS confirmando tu compra";
            MobileElement element = driver.findElement(By.id(LBL_MESSAGE));
            String elText = element.getText();
            if(elText.equals(textoToSearch)){
                ExtentReportUtil.INSTANCE.stepPass(driver, elText);
                generateWord.sendText(elText);
                generateWord.addImageToWord(driver);
                driver.findElement(By.id(BTN_ENVIAR)).click();
            }else{
                ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : No se presentó el mensaje esperado");
                generateWord.sendText("Fallo el caso de prueba : No se presentó el mensaje esperado");
                generateWord.addImageToWord(driver);
            }

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @When("^se ingresa a la opcion Beneficios Prepago$")
    public void seIngresaALaOpcionBeneficiosPrepago() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("4");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Beneficios Prepago");
            generateWord.sendText("Se ingresó correctamente a la opción : Beneficios Prepago");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(2000);
            driver.findElement(By.id(TXT_FIELD)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se ingresa la opcion Bono por antiguedad$")
    public void seIngresaLaOpcionBonoPorAntiguedad() throws Exception {
        try {
            Thread.sleep(2000);
            driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Bono por Antiguedad");
            generateWord.sendText("Se ingresó correctamente a la opción : Bono por Antiguedad");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se verifica el mensaje validador$")
    public void seVerificaElMensajeValidador() throws Exception {
        try {
            Thread.sleep(5000);
            String textoToSearch = "Gracias por ser parte de la familia Movistar. Recuerda buscar aqui tu premio por recargar cada mes.";
            MobileElement element = driver.findElement(By.id(LBL_MESSAGE));
            String elText = element.getText();
            if(elText.equals(textoToSearch)){
                ExtentReportUtil.INSTANCE.stepPass(driver, elText);
                generateWord.sendText(elText);
                generateWord.addImageToWord(driver);
                driver.findElement(By.id(BTN_ENVIAR)).click();
            }else{
                ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : No se presentó el mensaje esperado");
                generateWord.sendText("Fallo el caso de prueba : No se presentó el mensaje esperado");
                generateWord.addImageToWord(driver);
            }

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @When("^se ingresa a la opcion Servicios$")
    public void seIngresaALaOpcionServicios() throws Exception {
        try {
            Thread.sleep(2000);
            driver.findElement(By.id(TXT_FIELD)).sendKeys("5");
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Servicios");
            generateWord.sendText("Se ingresó correctamente a la opción : Consultas");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(3000);
            driver.findElement(By.id(TXT_FIELD)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se ingresa la opcion Comparte tu Saldo$")
    public void seIngresaLaOpcionComparteTuSaldo() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se elige la opción Comparte tu saldo");
            generateWord.sendText("Se elige la opción Comparte tu saldo");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(3000);
            driver.findElement(By.id(TXT_FIELD)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se ingresa el monto a compartir \\((\\d+) sol\\)$")
    public void seIngresaElMontoACompartirSol(int arg0) throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresa el monto a compartir: 1 sol");
            generateWord.sendText("Se ingresa el monto a compartir: 1 sol");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(3000);
            driver.findElement(By.id(TXT_FIELD)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se ingresa el telefono de destino$")
    public void seIngresaElTelefonoDeDestino() throws Exception {
        try {
            String numero = "958025001";
            driver.findElement(By.id(TXT_FIELD)).sendKeys(numero);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresa el número de destino: "+numero);
            generateWord.sendText("Se ingresa el número de destino: "+numero);
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(3000);
            driver.findElement(By.id(TXT_FIELD)).click();

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se confirma el envío$")
    public void seConfirmaElEnvío() throws Exception {
        try {
            Thread.sleep(2000);
            driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se Confirma el envío");
            generateWord.sendText("Se Confirma el envío");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se verifica el mensaje de saldo insuficiente$")
    public void seVerificaElMensajeDeSaldoInsuficiente() throws Exception {
        try {
            Thread.sleep(5000);
            String textoToSearch = "No tienes saldo suficiente.";
            MobileElement element = driver.findElement(By.id(LBL_MESSAGE));
            String elText = element.getText();
            if(elText.equals(textoToSearch)){
                ExtentReportUtil.INSTANCE.stepPass(driver, elText);
                generateWord.sendText(elText);
                generateWord.addImageToWord(driver);
                driver.findElement(By.id(BTN_ENVIAR)).click();
            }else{
                ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : No se presentó el mensaje esperado");
                generateWord.sendText("Fallo el caso de prueba : No se presentó el mensaje esperado");
                generateWord.addImageToWord(driver);
            }

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @When("^se ingresa a la opcion Roaming$")
    public void seIngresaALaOpcionRoaming() throws Exception {
        try {
            Thread.sleep(2000);
            driver.findElement(By.id(TXT_FIELD)).sendKeys("6");
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Roaming");
            generateWord.sendText("Se ingresó correctamente a la opción : Roaming");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(3000);
            driver.findElement(By.id(TXT_FIELD)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se ingresa la opcion consultar Cobertura$")
    public void seIngresaLaOpcionConsultarCobertura() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("3");
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se elige la opción Cobertura");
            generateWord.sendText("Se elige la opción Cobertura");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(3000);
            driver.findElement(By.id(TXT_FIELD)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se ingresa la opcion ver más$")
    public void seIngresaLaOpcionVerMás() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("9");
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se elige la opción Ver más");
            generateWord.sendText("Se elige la opción Ver más");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(3000);
            driver.findElement(By.id(TXT_FIELD)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se ingresa la opcion Continuar$")
    public void seIngresaLaOpcionContinuar() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se elige la opción Continuar");
            generateWord.sendText("Se elige la opción Continuar");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se verifica el mensaje de Cobertura Pasaporte Movistar$")
    public void seVerificaElMensajeDeCoberturaPasaporteMovistar() throws Exception {

        try {
            Thread.sleep(5000);
            String textoToSearch = "Cobertura Pasaporte Movistar";
            MobileElement element = driver.findElement(By.id(LBL_MESSAGE));
            String elText = element.getText();
            if(elText.equals(textoToSearch)){
                ExtentReportUtil.INSTANCE.stepPass(driver, elText);
                generateWord.sendText(elText);
                generateWord.addImageToWord(driver);
                driver.findElement(By.id(BTN_ENVIAR)).click();
            }else{
                ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : No se presentó el mensaje esperado");
                generateWord.sendText("Fallo el caso de prueba : No se presentó el mensaje esperado");
                generateWord.addImageToWord(driver);
            }

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se ingresa la opcion Continua aqui$")
    public void seIngresaLaOpcionContinuaAqui() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("2");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Continua aqui");
            generateWord.sendText("Se ingresó correctamente a la opción : Continua aqui");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(2000);
            driver.findElement(By.id(TXT_FIELD)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se elige Consultar Saldo$")
    public void seEligeConsultarSaldo() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Consultar Saldo");
            generateWord.sendText("Se ingresó correctamente a la opción : Consultar Saldo");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se obtiene el saldo actual$")
    public void seObtieneElSaldoActual() throws Exception {
        try {
            Thread.sleep(5000);
            //String textoToSearch = "Te enviaremos un SMS para descargar el app Mi Movistar";
            MobileElement element = driver.findElement(By.id(LBL_MESSAGE));
            String elText = element.getText();
                ExtentReportUtil.INSTANCE.stepPass(driver, "El saldo actual del chip es: "+elText);
                generateWord.sendText("El saldo actual del chip es: "+elText);
                generateWord.addImageToWord(driver);
                driver.findElement(By.id(BTN_ENVIAR)).click();


        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se ingresa a Mas sobre mi Prepago$")
    public void seIngresaAMasSobreMiPrepago() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("3");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Mas sobre mi Prepago");
            generateWord.sendText("Se ingresó correctamente a la opción : Mas sobre mi Prepago");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(2000);
            driver.findElement(By.id(TXT_FIELD)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se elige Consultar Contrato Prepago$")
    public void seEligeConsultarContratoPrepago() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Contrato Prepago");
            generateWord.sendText("Se ingresó correctamente a la opción : Contrato Prepago");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(2000);
            driver.findElement(By.id(TXT_FIELD)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se obtiene el contrato actual$")
    public void seObtieneElContratoActual() throws Exception {
        try {
            Thread.sleep(5000);
            //String textoToSearch = "Te enviaremos un SMS para descargar el app Mi Movistar";
            MobileElement element = driver.findElement(By.id(LBL_DIALOG));
            String elText = element.getText();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se optiene el contrato actual : "+elText);
            generateWord.sendText("Se optiene el contrato actual : "+elText);
            generateWord.addImageToWord(driver);

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se elige Tarifas$")
    public void seEligeTarifas() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("2");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Tarifas");
            generateWord.sendText("Se ingresó correctamente a la opción : Tarifas");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(2000);
            driver.findElement(By.id(TXT_FIELD)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se elige la opción LLAMADAS$")
    public void seEligeLaOpciónLLAMADAS() throws Exception {
            try {
                driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
                Thread.sleep(2000);
                ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : LLAMADAS");
                generateWord.sendText("Se ingresó correctamente a la opción : LLAMADAS");
                generateWord.addImageToWord(driver);
                driver.findElement(By.id(BTN_ENVIAR)).click();
            } catch (Exception e) {
                ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
                generateWord.sendText("Tiempo de espera ha excedido");
                generateWord.addImageToWord(driver);
            }

    }

    @And("^se elige la opción DATOS$")
    public void seEligeLaOpciónDATOS() throws Exception {
            try {
                driver.findElement(By.id(TXT_FIELD)).sendKeys("2");
                Thread.sleep(2000);
                ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : DATOS");
                generateWord.sendText("Se ingresó correctamente a la opción : DATOS");
                generateWord.addImageToWord(driver);
                driver.findElement(By.id(BTN_ENVIAR)).click();
            } catch (Exception e) {
                ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
                generateWord.sendText("Tiempo de espera ha excedido");
                generateWord.addImageToWord(driver);
            }

    }

    @Then("^se obtiene la tarifa de LLAMADAS$")
    public void seObtieneLaTarifaDeLLAMADAS() throws Exception {
        try {
            Thread.sleep(5000);
            //String textoToSearch = "Te enviaremos un SMS para descargar el app Mi Movistar";
            MobileElement element = driver.findElement(By.id(LBL_MESSAGE));
            String elText = element.getText();
            ExtentReportUtil.INSTANCE.stepPass(driver, "La Tarifa de LLAMADAS es: "+elText);
            generateWord.sendText("La Tarifa de LLAMADAS es: "+elText);
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se obtiene la tarifa de DATOS$")
    public void seObtieneLaTarifaDeDATOS() throws Exception {
        try {
            Thread.sleep(5000);
            //String textoToSearch = "Te enviaremos un SMS para descargar el app Mi Movistar";
            MobileElement element = driver.findElement(By.id(LBL_MESSAGE));
            String elText = element.getText();
            ExtentReportUtil.INSTANCE.stepPass(driver, "La Tarifa de DATOS es: "+elText);
            generateWord.sendText("La Tarifa de DATOS es: "+elText);
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se elige Perdida o robo de un equipo$")
    public void seEligePerdidaORoboDeUnEquipo() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("3");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Perdida o robo de un equipo");
            generateWord.sendText("Se ingresó correctamente a la opción : Perdida o robo de un equipo");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se obtiene la información  solicitada$")
    public void seObtieneLaInformaciónSolicitada() throws Exception {
        try {
            Thread.sleep(5000);
            //String textoToSearch = "Te enviaremos un SMS para descargar el app Mi Movistar";
            MobileElement element = driver.findElement(By.id(LBL_MESSAGE));
            String elText = element.getText();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se obtiene la información solicitada : "+elText);
            generateWord.sendText("Se obtiene la información solicitada "+elText);
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se elige Llamadas internacionales$")
    public void seEligeLlamadasInternacionales() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("4");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Llamadas Internacionales");
            generateWord.sendText("Se ingresó correctamente a la opción :  Llamadas Internacionales");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
            Thread.sleep(2000);
            driver.findElement(By.id(TXT_FIELD)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se elige la opción Movistar (\\d+)$")
    public void seEligeLaOpciónMovistar(int arg0) throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("1");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Movistar 1911");
            generateWord.sendText("Se ingresó correctamente a la opción : Movistar 1911");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se obtiene la información de Movistar (\\d+)$")
    public void seObtieneLaInformaciónDeMovistar(int arg0) throws Exception {
        try {
            Thread.sleep(5000);
            //String textoToSearch = "Te enviaremos un SMS para descargar el app Mi Movistar";
            MobileElement element = driver.findElement(By.id(LBL_MESSAGE));
            String elText = element.getText();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se obtiene la información solicitada : "+elText);
            generateWord.sendText("Se obtiene la información solicitada "+elText);
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @And("^se elige la opción Otros Operadores$")
    public void seEligeLaOpciónOtrosOperadores() throws Exception {
        try {
            driver.findElement(By.id(TXT_FIELD)).sendKeys("2");
            Thread.sleep(2000);
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se ingresó correctamente a la opción : Otros Operadores");
            generateWord.sendText("Se ingresó correctamente a la opción : Otros Operadores");
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();
        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }

    @Then("^se obtiene la información de Otros Operadores$")
    public void seObtieneLaInformaciónDeOtrosOperadores() throws Exception {
        try {
            Thread.sleep(5000);
            //String textoToSearch = "Te enviaremos un SMS para descargar el app Mi Movistar";
            MobileElement element = driver.findElement(By.id(LBL_MESSAGE));
            String elText = element.getText();
            ExtentReportUtil.INSTANCE.stepPass(driver, "Se obtiene la información solicitada : "+elText);
            generateWord.sendText("Se obtiene la información solicitada "+elText);
            generateWord.addImageToWord(driver);
            driver.findElement(By.id(BTN_ENVIAR)).click();

        } catch (Exception e) {
            ExtentReportUtil.INSTANCE.stepFail(driver, "Fallo el caso de prueba : " + e.getMessage());
            generateWord.sendText("Tiempo de espera ha excedido");
            generateWord.addImageToWord(driver);
        }
    }
}