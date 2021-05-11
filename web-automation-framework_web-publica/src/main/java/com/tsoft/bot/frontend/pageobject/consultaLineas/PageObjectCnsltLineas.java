package com.tsoft.bot.frontend.pageobject.consultaLineas;

import org.openqa.selenium.By;

public class PageObjectCnsltLineas {

    //Login del Portal
    public static By TXT_USER = By.xpath("//*[@id=\"_58_login\"]");
    public static By TXT_PASS = By.xpath("//*[@id=\"_58_password\"]");
    public static By BTN_Acceder = By.xpath("//*[@id=\"_58_fm\"]/div/button");

    //Consulta Lineas
    public static By LBL_CONSULTA = By.xpath("//*[@id=\"p_p_id_consultmobilenumbers_WAR_consultmobilenumbersportlet_\"]/div/div[2]/div/div/div/div[1]/div/h1");
    public static By LST_TIPO_DOC = By.xpath("//input[@value='Tipo de documento']");

    public static By LBL_DOCUMENTO = By.xpath("//label[@for='_consultmobilenumbers_WAR_consultmobilenumbersportlet_documentNumber']");
    public static By TXT_DOCUMENTO = By.xpath("//*[@id=\"_consultmobilenumbers_WAR_consultmobilenumbersportlet_documentNumber\"]");
    public static By BTN_CONSULTAR = By.xpath("//*[@id=\"_consultmobilenumbers_WAR_consultmobilenumbersportlet_btnSubmit\"]");

    public static By LST_TIPO_DOCRL = By.xpath("/html/body/div[4]/div/div[1]/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/form/fieldset[2]/div/div[1]/div/input");


    public static By LBL_DOCUMENTORL = By.xpath("/html/body/div[4]/div/div[1]/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/form/fieldset[3]/div/div/label/label");

    //public static By TXT_DOCUMENTORL = By.xpath("/html/body/div[4]/div/div[1]/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/form/fieldset[3]/div/div/label/input");
    public static By TXT_DOCUMENTORL = By.xpath("//*[@id=\"_consultmobilenumbers_WAR_consultmobilenumbersportlet_documentNumberRpstative\"]");

    // id="_consultmobilenumbers_WAR_consultmobilenumbersportlet_documentNumberRpstative"
//nueva consulta

    //public static By LBL_TABLA_RESULTADO = By.xpath("/html/body/div[4]/div/div[1]/div[2]/div/div/div/div/div[2]/div");
    public static By LBL_TABLA_RESULTADO = By.xpath("/html/body/div[4]/div/div[1]/div[2]/div/div/div/div/div[2]/div/div/div");
}
