package com.tsoft.bot.frontend.pageobject.Fijas;

import io.appium.java_client.AppiumDriver;
import io.appium.java_client.MobileElement;
import org.omg.CORBA.PUBLIC_MEMBER;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

public class PageObject_Fijas {
    private static AppiumDriver<MobileElement> driver;

    //LOGIN
    public static final String TXT_CODIGO = "com.telefonica.ventafija.dev:id/et_cod_atis";//id
    public static final String BTN_PRE_LOGUEO = "com.telefonica.ventafija.dev:id/btn_pre_login";//id
    public static final String BTN_LOGIN = "com.telefonica.ventafija.dev:id/btn_login";//id
    public static final String TXT_PASSWORD = "com.telefonica.ventafija.dev:id/edt_password";//id
    public static final String TXT_TOKEN = "com.telefonica.ventafija.dev:id/edt_token";//id
    public static final String POPUP_SAVE_PASSWORD = "android:id/autofill_save_no";//id
    public static final String POPUP_AGREE = "com.telefonica.ventafija.dev:id/btn_agree";//id

    //MENU PRINCIPAL
    public static final String FORM_MENU = "com.telefonica.ventafija.dev:id/view";//id
    public static final String BTN_ALTA = "com.telefonica.ventafija.dev:id/btn_alta";//id
    public static final String BTN_MIGRACION = "com.telefonica.ventafija.dev:id/btn_migration";//id
    public static final String BTN_SVAS = "com.telefonica.ventafija.dev:id/btn_sva";//id

    //ALTA
    public static final String CBX_DOCUMENTO = "com.telefonica.ventafija.dev:id/txt_document_type";//id
    public static final String TXT_DOCUMENTO = "com.telefonica.ventafija.dev:id/txt_document";//id
    public static final String CBX_DEPARTAMENTO = "com.telefonica.ventafija.dev:id/txt_department";//id
    public static final String CBX_PROVINCIA = "com.telefonica.ventafija.dev:id/txt_province";//id
    public static final String CBX_DISTRITO = "com.telefonica.ventafija.dev:id/txt_district";//id
    public static final String BTN_EVALUAR = "com.telefonica.ventafija.dev:id/btnevaluar";//id
    public static final String LBL_MONTO = "com.telefonica.ventafija.dev:id/txtRentaMaxima2";//id
    public static final String LBL_MONTO_LINEA = "com.telefonica.ventafija.dev:id/textViewProductType";//id
    public static final String BTN_INICIAR_VENTA = "com.telefonica.ventafija.dev:id/btnStartSale";//id
    //DIRECCION DE INSTALACION
    public static final String POP_UP = "com.android.packageinstaller:id/permission_allow_button";//id
    public static final String POP_UP_CERRAR = "com.telefonica.ventafija.dev:id/btn_dialognormalizador";//id
    public static final String TXT_DIRECCION = "com.telefonica.ventafija.dev:id/txt_address";//id
    public static final String BTN_BUSCAR = "com.telefonica.ventafija.dev:id/btn_search";//id
    public static final String TXT_REFERENCIA = "com.telefonica.ventafija.dev:id/txt_address_reference";//id
    public static final String TXT_CELULAR = "com.telefonica.ventafija.dev:id/txt_phone";//id
    public static final String TXT_TELEFONO_ADICIONAL = "com.telefonica.ventafija.dev:id/txt_secondary_phone";//id
    public static final String BTN_CONTINUAR = "com.telefonica.ventafija.dev:id/btn_continue";//id
    //COMPLETA DIRECCION DE INSTALACION
    public static final String CBX_TIPOVIA = "com.telefonica.ventafija.dev:id/txt_via_tipo";//id
    public static final String TXT_NOMBREVIA = "com.telefonica.ventafija.dev:id/edt_via_name";//id
    public static final String TXT_CUADRA = "com.telefonica.ventafija.dev:id/edt_nro_cuadra";//id
    public static final String TXT_NRO_PUERTA = "com.telefonica.ventafija.dev:id/edt_nro_puerta_1";//id
    public static final String CBX_CCHH = "com.telefonica.ventafija.dev:id/txt_cchh_tipo";//id
    public static final String TXT_NOMBRE_CCHH = "com.telefonica.ventafija.dev:id/edt_cchh_name";//id
    public static final String TXT_MANZANA = "com.telefonica.ventafija.dev:id/edt_manzana";//id
    public static final String TXT_LOTE = "com.telefonica.ventafija.dev:id/edt_lote";//id
    public static final String TXT_EDITAR_REFERENCIA = "com.telefonica.ventafija.dev:id/edt_refence";//id

    //LISTA DE CAMPAÑAS
    public static final String BTN_INTENTAR = "com.telefonica.ventafija.dev:id/btn_get_catalog_retry"; //ID
    public static final String BTN_VER_TODOS = "com.telefonica.ventafija.dev:id/btn_show_all_catalog_items"; //ID
    //public static final String OPC_1 = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.view.ViewGroup/android.view.ViewGroup/android.view.ViewGroup/androidx.recyclerview.widget.RecyclerView/android.widget.FrameLayout[1]/android.view.ViewGroup/android.widget.TextView[1]"; //Xpath
    //public static final String OPC_2 = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.view.ViewGroup/android.view.ViewGroup/android.view.ViewGroup/androidx.recyclerview.widget.RecyclerView/android.widget.FrameLayout[2]/android.view.ViewGroup/android.widget.TextView[1]"; //Xpath
    //public static final String OPC_3 = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.view.ViewGroup/android.view.ViewGroup/android.view.ViewGroup/androidx.recyclerview.widget.RecyclerView/android.widget.FrameLayout[3]/android.view.ViewGroup/android.widget.TextView[1]"; //Xpath
    //public static final String OPC_4 = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.view.ViewGroup/android.view.ViewGroup/android.view.ViewGroup/androidx.recyclerview.widget.RecyclerView/android.widget.FrameLayout[4]/android.view.ViewGroup/android.widget.TextView[1]"; //Xpath
    public static final String OPC_1 = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.view.ViewGroup/android.view.ViewGroup/androidx.recyclerview.widget.RecyclerView/android.widget.FrameLayout[1]/android.view.ViewGroup/android.widget.TextView[1]"; //Xpath
    public static final String OPC_2 = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.view.ViewGroup/android.view.ViewGroup/androidx.recyclerview.widget.RecyclerView/android.widget.FrameLayout[2]/android.view.ViewGroup/android.widget.TextView[1]"; //Xpath
    public static final String OPC_3 = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.view.ViewGroup/android.view.ViewGroup/androidx.recyclerview.widget.RecyclerView/android.widget.FrameLayout[3]/android.view.ViewGroup/android.widget.TextView[1]"; //Xpath
    public static final String OPC_4 = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.view.ViewGroup/android.view.ViewGroup/androidx.recyclerview.widget.RecyclerView/android.widget.FrameLayout[4]/android.view.ViewGroup/android.widget.TextView[1]"; //Xpath
    public static final String OPC_5 = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.view.ViewGroup/android.view.ViewGroup/androidx.recyclerview.widget.RecyclerView/android.widget.FrameLayout[5]/android.view.ViewGroup/android.widget.TextView[1]"; //Xpath
    public static final String OPC_SELECT_1 = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.view.ViewGroup/android.view.ViewGroup/androidx.recyclerview.widget.RecyclerView/android.widget.FrameLayout[1]/android.view.ViewGroup"; //Xpath
    public static final String OPC_SELECT_2 = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.view.ViewGroup/android.view.ViewGroup/androidx.recyclerview.widget.RecyclerView/android.widget.FrameLayout[2]/android.view.ViewGroup"; //Xpath
    public static final String OPC_SELECT_3 = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.view.ViewGroup/android.view.ViewGroup/androidx.recyclerview.widget.RecyclerView/android.widget.FrameLayout[3]/android.view.ViewGroup"; //Xpath
    public static final String OPC_SELECT_4 = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.view.ViewGroup/android.view.ViewGroup/androidx.recyclerview.widget.RecyclerView/android.widget.FrameLayout[4]/android.view.ViewGroup"; //Xpath

    //DETALLE PRODUCTO
    public static final String LBL_PRECIO_PRODUCTO = "com.telefonica.ventafija.dev:id/txtPriceNormal"; //ID
    public static final String BTN_DP_SIGUIENTE = "com.telefonica.ventafija.dev:id/btnGoSva"; //ID

    //SVA
    public static final String CBX_BLOQUE_TV = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.view.ViewGroup/android.widget.ScrollView/android.view.ViewGroup/android.widget.LinearLayout[1]/android.widget.CheckBox[2]"; //XPATH
    public static final String BTN_SVA_SIGUIENTE = "com.telefonica.ventafija.dev:id/btn_continue"; //ID

    //CONDICIONES
    public static final String LBL_CONDICIONES = "com.telefonica.ventafija.dev:id/txt_TituloDebitoAutomatico"; //id
    public static final String CBX_DEBITO_AUTOMATICO_SI = "com.telefonica.ventafija.dev:id/cb_si_DebitoAutomatico"; //id
    public static final String CBX_DEBITO_AUTOMATICO_NO = "com.telefonica.ventafija.dev:id/cb_no_DebitoAutomatico"; //id
    public static final String CBX_TRATAMIENTO_DATOS_SI = "com.telefonica.ventafija.dev:id/cb_si_TratamientoDatos"; //id
    public static final String CBX_TRATAMIENTO_DATOS_NO = "com.telefonica.ventafija.dev:id/cb_no_TratamientoDatos"; //id
    public static final String CBX_CONTRATO_DIGITAL_SI = "com.telefonica.ventafija.dev:id/cb_si_PackVerde"; //id
    public static final String CBX_CONTRATO_DIGITAL_NO = "com.telefonica.ventafija.dev:id/cb_no_PackVerde"; //id
    public static final String CBX_WEB_PARENTAL_SI = "com.telefonica.ventafija.dev:id/cb_si_controlParental"; //id
    public static final String CBX_WEB_PARENTAL_NO = "com.telefonica.ventafija.dev:id/cb_no_controlParental"; //id
    public static final String BTN_CONDICIONES_CONTINUAR = "com.telefonica.ventafija.dev:id/btn_next"; //id

    //RESUMEN DE VENTA
    public static final String LBL_PLAN = "com.telefonica.ventafija.dev:id/txt_plan"; //id
    public static final String BTN_RESUMEN_VENTA_ACEPTAR = "com.telefonica.ventafija.dev:id/btn_next"; //id

    //LECTURA DE CONTRATO
    public static final String BTN_PLAY = "com.telefonica.ventafija.dev:id/img_audio_controller"; //id
    public static final String POP_UP_ALLOW = "com.android.packageinstaller:id/permission_allow_button"; //id
    public static final String BTN_LECTURA_CONTINUAR = "com.telefonica.ventafija.dev:id/txt_continue"; //id

    //VALIDACION DE IDENTIDAD
    public static final String LBL_NOMBRE_CLIENTE = "com.telefonica.ventafija.dev:id/txt_client_name_value"; //id
    public static final String TXT_EMAIL = "com.telefonica.ventafija.dev:id/edt_email"; //id
    public static final String CBX_SELECCIONAR_MADRE = "com.telefonica.ventafija.dev:id/text_input_end_icon"; //id
    public static final String CBX_LBL_NOMBRES_MADRE = "com.telefonica.ventafija.dev:id/ll_data"; //id
    public static final String TXT_NOMBRE_MADRE = "com.telefonica.ventafija.dev:id/txt_madre_padre"; //id
    public static final String BTN_VALIDACION_SIGUIENTE = "com.telefonica.ventafija.dev:id/btn_next"; //id

    //TOMAR FOTO
    public static final String IMG_FRONTAL_DNI = "com.telefonica.ventafija.dev:id/img_document_front"; //id
    public static final String IMG_POSTERIOR_DNI = "com.telefonica.ventafija.dev:id/img_document_back"; //id
    public static final String OPC_ELEGIR_GALERIA = "/hierarchy/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.ListView/android.widget.TextView[2]"; //xpath
    public static final String GALERIA = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.RelativeLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.view.ViewGroup/android.widget.LinearLayout/android.widget.FrameLayout/android.support.v7.widget.RecyclerView/android.widget.LinearLayout[1]/android.widget.FrameLayout/android.widget.ImageView[1]"; //xpath
    public static final String GALERIA_FOTO1 = "(//android.widget.FrameLayout[@content-desc='Button'])[1]/android.widget.ImageView"; //xpath
    public static final String BTN_CROP = "com.telefonica.ventafija.dev:id/crop_image_menu_crop"; //ID
    public static final String BTN_FOTO_SIGUIENTE = "com.telefonica.ventafija.dev:id/btn_continue"; //ID

    //VENTA EXITOSA
    public static final String IMG_EXITOSA = "com.telefonica.ventafija.dev:id/check"; //ID

    //MIGRACIONES - PRODUCTOS DEL CLIENTE
    public static final String BTN_BUSCAR_PRODUCTO = "com.telefonica.ventafija.dev:id/btn_find";
    public static final String OPCION_PRODUCTO = "/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/androidx.drawerlayout.widget.DrawerLayout/android.view.ViewGroup/android.view.ViewGroup/android.widget.FrameLayout/android.view.ViewGroup/androidx.viewpager.widget.ViewPager/androidx.recyclerview.widget.RecyclerView/android.view.ViewGroup/androidx.recyclerview.widget.RecyclerView/androidx.cardview.widget.CardView/android.view.ViewGroup"; //XPATH
    public static final String LBL_SERVICIO = "com.telefonica.ventafija.dev:id/txt_code"; //ID
    public static final String LBL_PRECIO= "com.telefonica.ventafija.dev:id/txt_price"; //ID
    public static final String LBL_DIRECCION= "com.telefonica.ventafija.dev:id/txt_address"; //ID
    public static final String BTN_SELECCIONAR = "com.telefonica.ventafija.dev:id/btn_select"; //ID



}
