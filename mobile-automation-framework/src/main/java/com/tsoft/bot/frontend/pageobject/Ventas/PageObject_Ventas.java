package com.tsoft.bot.frontend.pageobject.Ventas;

import org.omg.CORBA.PUBLIC_MEMBER;
import org.openqa.selenium.By;

public class PageObject_Ventas {

    //AUTHORITIES
    public static final String POPUP_AUTH = "com.android.packageinstaller:id/permission_allow_button";//id
    public static final String POPUP_AUTH2 = "com.android.packageinstaller:id/permission_allow_button";//id
    public static By POPUP_AUTH3 = By.id("com.android.packageinstaller:id/permission_allow_button"); //Después de iniciar sesión

    //LOGIN
    public static String TXT_DNI_VENDEDOR = "pe.vasslatam.movistar.mobile.sales:id/txt_dni";
    //public static String TXT_DNI_VENDEDOR = "pe.vasslatam.movistar.mobile.sales:id/textinput_placeholder";
    public static final String BTN_INGRESAR = "pe.vasslatam.movistar.mobile.sales:id/button";//id

    //CONFIGURACION DE UBICACIÓN
    public static final String FORM_CONFIGURACION = "pe.vasslatam.movistar.mobile.sales:id/tv_title";//id
    public static final String BTN_GUARDAR = "pe.vasslatam.movistar.mobile.sales:id/btn_save";//id
    public static final String BTN_ACEPTAR_CONFIGURACION = "pe.vasslatam.movistar.mobile.sales:id/btn_acept";//id
    public static final String FORM_PRODUCTOS = "pe.vasslatam.movistar.mobile.sales:id/ivBg";//id

    //REUTILIZABLES
    public static final String LBL_GENERAL = "pe.vasslatam.movistar.mobile.sales:id/tvTitle";
    public static final String BTN_NUEVA_LINEA = "pe.vasslatam.movistar.mobile.sales:id/ll_new_line";//id
    public static final String LBL_ESCOGER_MODALIDAD = "pe.vasslatam.movistar.mobile.sales:id/tv_escoge";//id
    public static final String LBL_CODIGO_BARRAS = "pe.vasslatam.movistar.mobile.sales:id/tvTitle";//id

    //PREPAGO
    public static final String BTN_PREPAGO = "pe.vasslatam.movistar.mobile.sales:id/ll_prepaid";//id
    public static final String BTN_CHIP_SOLO = "pe.vasslatam.movistar.mobile.sales:id/ll_SIM"; //id
    public static final String CBX_PLAN = "//android.widget.TextView[@resource-id='android:id/text1']"; //Xpath
    public static final String TXT_SERIE_CHIP = "pe.vasslatam.movistar.mobile.sales:id/txt_iccid";
    public static By TXT_SERIECHIP = By.id("pe.vasslatam.movistar.mobile.sales:id/txt_iccid");
    public static final String BTN_AVISO = "pe.vasslatam.movistar.mobile.sales:id/btn_acept";//id
    public static final String BTN_AVISO2 = "pe.vasslatam.movistar.mobile.sales:id/btn_acept";//id
    public static final String TXT_NUEVO_NUMERO = "pe.vasslatam.movistar.mobile.sales:id/txt_telephone";//id
    public static final String CHBX_PREPLAN_FLEX = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.view.ViewGroup/android.widget.ScrollView/android.view.ViewGroup/android.widget.LinearLayout[3]/android.widget.LinearLayout/android.widget.CheckBox[2]";//XPATH
    public static final String CHBX_TARIFA_UNICA = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.view.ViewGroup/android.widget.ScrollView/android.view.ViewGroup/android.widget.LinearLayout[3]/android.widget.LinearLayout/android.widget.CheckBox[1]";//XPATH
    public static final String BTN_CONTINUAR = "pe.vasslatam.movistar.mobile.sales:id/btn_next";//ID

    //ALTA POSTPAGO
    public static final String BTN_POSTPAGO = "pe.vasslatam.movistar.mobile.sales:id/ll_postpaid";//id
    public static By LBL_DATOS_CLIENTE = By.id("pe.vasslatam.movistar.mobile.sales:id/tv_title");
    public static String SELECT_DOCUMENTO = "//android.widget.TextView[@resource-id='android:id/text1']";
    public static String MSG_ERROR = "/hierarchy/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.ScrollView/android.view.View";

    public static String TXT_DOCUMENTO = "pe.vasslatam.movistar.mobile.sales:id/txt_DNI";//id
    public static final String TXT_EMAIL = "pe.vasslatam.movistar.mobile.sales:id/txt_email";//id
    public static By BTN_CHIP_SOLO_POST = By.id("pe.vasslatam.movistar.mobile.sales:id/ll_SIM"); //----------------
    public static By TXT_SERIE_ALTA_POSTPAGO = By.id("pe.vasslatam.movistar.mobile.sales:id/txt_iccid"); //$$$$$$$$$$$$$$$$$$$4
    public static By SELECT_PLAN = By.id("android:id/text1");
    //public static By BTN_CONTINUAR = By.id("pe.vasslatam.movistar.mobile.sales:id/btn_next");
    public static final String TXT_DIRECCION = "pe.vasslatam.movistar.mobile.sales:id/txt_address";//id
    public static final String BTN_CONTINUAR2 = "pe.vasslatam.movistar.mobile.sales:id/btn_continue";//id
    public static By BTN_ERROR = By.id("pe.vasslatam.movistar.mobile.sales:id/btn_acept");//id
    public static By TITLE_ERROR = By.id("pe.vasslatam.movistar.mobile.sales:id/txt_tittle");//id

    //RENOVACIÓN
    public static final String BTN_RENOVACION = "pe.vasslatam.movistar.mobile.sales:id/ll_renew"; //id
    public static final String BTN_PORTABILIDAD = "pe.vasslatam.movistar.mobile.sales:id/ll_porta";//id
    public static final String BTN_VALIDAR_LINEA = "pe.vasslatam.movistar.mobile.sales:id/ll_PORTA";//id
    public static final String CBX_OPERADOR = "pe.vasslatam.movistar.mobile.sales:id/cmb_operator";//id
    public static final String CBX_OPERADOR_ACTUAL = "pe.vasslatam.movistar.mobile.sales:id/cmb_productType";//id
    public static final String TXT_CONDICIONES = "pe.vasslatam.movistar.mobile.sales:id/txt_leido";//id

}
