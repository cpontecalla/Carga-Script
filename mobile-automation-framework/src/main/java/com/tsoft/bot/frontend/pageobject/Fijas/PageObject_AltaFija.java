package com.tsoft.bot.frontend.pageobject.Fijas;

import io.appium.java_client.AppiumDriver;
import io.appium.java_client.MobileElement;

public class PageObject_AltaFija {
    private static AppiumDriver<MobileElement> driver;

    //ALTA
    public static final String CBX_DOCUMENTO = "com.telefonica.ventafija.dev.debug:id/txt_document_type";//id
    public static final String TXT_DOCUMENTO = "com.telefonica.ventafija.dev.debug:id/txt_document";//id
    public static final String CBX_DEPARTAMENTO = "com.telefonica.ventafija.dev.debug:id/txt_department";//id
    public static final String CBX_PROVINCIA = "com.telefonica.ventafija.dev.debug:id/txt_province";//id
    public static final String CBX_DISTRITO = "com.telefonica.ventafija.dev.debug:id/txt_district";//id
    public static final String BTN_EVALUAR = "com.telefonica.ventafija.dev.debug:id/btnevaluar";//id
    public static final String LBL_MONTO = "com.telefonica.ventafija.dev.debug:id/txtRentaMaxima2";//id
    public static final String LBL_MONTO_LINEA = "com.telefonica.ventafija.dev.debug:id/textViewProductType";//id
    public static final String BTN_INICIAR_VENTA = "com.telefonica.ventafija.dev.debug:id/btnStartSale";//id

    //DIRECCION DE INSTALACION
    public static final String POP_UP = "com.android.packageinstaller:id/permission_allow_button";//id
    public static final String POP_UP_CERRAR = "com.telefonica.ventafija.dev.debug:id/btn_dialognormalizador";//id
    public static final String TXT_DIRECCION = "com.telefonica.ventafija.dev.debug:id/txt_address";//id
    public static final String BTN_BUSCAR = "com.telefonica.ventafija.dev.debug:id/btn_search";//id
    public static final String TXT_REFERENCIA = "com.telefonica.ventafija.dev.debug:id/txt_address_reference";//id
    public static final String TXT_CELULAR = "com.telefonica.ventafija.dev.debug:id/txt_phone";//id
    public static final String TXT_TELEFONO_ADICIONAL = "com.telefonica.ventafija.dev.debug:id/txt_secondary_phone";//id
    public static final String BTN_CONTINUAR = "com.telefonica.ventafija.dev.debug:id/btn_continue";//id

    //COMPLETA DIRECCION DE INSTALACION
    public static final String CBX_TIPOVIA = "com.telefonica.ventafija.dev.debug:id/txt_via_tipo";//id
    public static final String TXT_NOMBREVIA = "com.telefonica.ventafija.dev.debug:id/edt_via_name";//id
    public static final String TXT_CUADRA = "com.telefonica.ventafija.dev.debug:id/edt_nro_cuadra";//id
    public static final String TXT_NRO_PUERTA = "com.telefonica.ventafija.dev.debug:id/edt_nro_puerta_1";//id
    public static final String CBX_CCHH = "com.telefonica.ventafija.dev.debug:id/txt_cchh_tipo";//id
    public static final String TXT_NOMBRE_CCHH = "com.telefonica.ventafija.dev.debug:id/edt_cchh_name";//id
    public static final String TXT_MANZANA = "com.telefonica.ventafija.dev.debug:id/edt_manzana";//id
    public static final String TXT_LOTE = "com.telefonica.ventafija.dev.debug:id/edt_lote";//id
    public static final String TXT_EDITAR_REFERENCIA = "com.telefonica.ventafija.dev.debug:id/edt_refence";//id

    //LISTA DE CAMPAÑAS
    public static final String BTN_INTENTAR = "com.telefonica.ventafija.dev.debug:id/btn_get_catalog_retry"; //ID


}
