VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private strCodigo As String
Private strNombre As String
Private strPassword As String
Private blnConectado As Boolean
Private strEstado As String
Private strPerfil As String
Private strTipDocDefault As String
Private strMotivoGeneraGuiaLocal As String
Private rsUsuario As oraDynaset
Private rsMaquina As oraDynaset
Private strCentroCosto As String
Private rsParametros As oraDynaset
Private rsParametrosLocal As oraDynaset
Private rsInicioCaja As oraDynaset
Private rsMoneda As oraDynaset
Private arrEmpresa As Variant
Private strCodTipoVenta As Variant
Private xParametros As New XArrayDB
Private objMaquina As New clsMaquina
Private objFormaPago As New clsFormaPago
Private objMoneda As New clsMoneda
Private strNombreMaquina As String
Private strMayoristaFracciones As Boolean
Private strCodigoLiquidacion As String
Private strMetaValePromedio As Double
Private strMetaNumeroCliente As Double
Dim strTipoMaquina  As String
Dim strIPMaquina As String
Dim strCia As String
Dim strCodLocalCallCenter As String
Dim strDireccionLocal As String
Dim strLocal As String
Dim strFlgDelivery As String
Dim strNombreLocal As String
Dim strUbigeoLocal As String
Dim dbltipocambio As Double
Dim blnEsDelivery As Boolean
Dim blnPrecioOnLine As Boolean
Dim strTipoMaquinaCabina As String
Dim strFlgDeliveryProv As String
Dim strCodMotorizado As String

Public Function Login(ByVal strCodUsuario As String, strPassword As String) As oraDynaset

    If objMaquina.valida = False Then 'evalua en la base de datos que el ip se encuentre registrado
        End 'si la maquina no esta registrada en la base de datos
        'sale de la aplicación
    Else
        'de lo contraro llena las variables para seguir cargado los parametros
        strIPMaquina = Obtener(12)
        strNombreMaquina = objMaquina.NombrePC
        strCia = objMaquina.Cia
        strLocal = objMaquina.CodigoLocal
        strFlgDelivery = objMaquina.Delivery
        strTipoMaquina = objMaquina.TipoMaquina
        strNombreLocal = objMaquina.NombreLocal
        'strCentroCosto = objMaquina.CentroCosto
        strDireccionLocal = objMaquina.DireccionLocal
        strUbigeoLocal = objMaquina.UbigeoLocal
        strMayoristaFracciones = MayoristaFracciones
        objUsuario.CodLocalCallCenter = "" & objMaquina.CodLocalCallCenter
        On Error GoTo Pase
        strTipDocDefault = objMaquina.ListaDocDefecto(strCia, strIPMaquina)
Pase:
        flgDeliveryProv = objMaquina.flgDeliveryProv
        
        'strMotivoGeneraGuiaLocal = MotivoGeneraGuiaLocal
    End If
    
    gclsOracle.Num_Intentos = 1
    On Error GoTo Handle
    
    Set rsUsuario = gclsOracle.FN_Cursor("CMR.PKG_USUARIO.FN_LOGIN", 0, strCodUsuario, strPassword)
    
    If rsUsuario Is Nothing Then blnConectado = False: Exit Function
    
    If rsUsuario.RecordCount = 0 Then MsgBox "No existe ningun usuario con el codigo " & strCodUsuario, vbCritical, App.ProductName
    
    strCodigo = "" & rsUsuario("COD_USUARIO").Value
    strNombre = "" & rsUsuario("NOM_USUARIO").Value
    Password = "" & rsUsuario("PAS_USUARIO").Value
    strEstado = "" & rsUsuario("PAS_USUARIO").Value
    strPerfil = "" & rsUsuario("COD_PERFIL").Value
    strCentroCosto = "" & rsUsuario("DES_CNT_COSTO").Value
    
    If objUsuario.ValidaVersion = False Then
        MsgBox "La version " & gstrVersion & " no esta permitida, comunicarse con el area de Mesa de Ayuda", vbCritical, App.ProductName
        End
    End If
     TRAEMETA

    'cambio por PHERRERA 06/10/08 - Hubo problemas con locales de Delivery en provincia (el sistema cambia el local a DLV)
    If objUsuario.flgDeliveryProv = "0" Then
        If UsuarioEsDelLocal(strLocal, strCodigo, strPerfil) = 0 Then
            blnConectado = False
            MsgBox "El usuario no esta asignado al local", vbCritical, App.ProductName
            Exit Function
        End If
    End If
    
    Set Login = rsUsuario
    
    blnConectado = True
    Set rsParametros = gclsOracle.FN_Cursor("BTLPROD.PKG_EMPRESA.FN_LISTA", 0, strCia)
    
    arrEmpresa(0) = "" & rsParametros("CIA").Value
    arrEmpresa(1) = "" & rsParametros("ESTADO").Value
    arrEmpresa(2) = "" & rsParametros("DESCRIP").Value
    arrEmpresa(3) = "" & rsParametros("SIGLAS").Value
    arrEmpresa(4) = "" & rsParametros("RUC").Value
    arrEmpresa(5) = "" & rsParametros("TIPO").Value
    arrEmpresa(6) = "" & rsParametros("DIRECCION").Value
    arrEmpresa(7) = "" & rsParametros("DISTRITO").Value
    arrEmpresa(8) = "" & rsParametros("URBANIZA").Value
    arrEmpresa(9) = "" & rsParametros("PROVINCIA").Value
    arrEmpresa(10) = "" & rsParametros("DPTO").Value
    arrEmpresa(11) = "" & rsParametros("PAIS").Value
    arrEmpresa(12) = "" & rsParametros("TELEFONO1").Value
    arrEmpresa(13) = "" & rsParametros("TELEFONO2").Value
    arrEmpresa(14) = "" & rsParametros("FAX").Value
    arrEmpresa(19) = "" & rsParametros("FLG_IMPRIME_DIR").Value
    
    arrEmpresa(18) = "" & Valor("COD_MONEDA", CodigoEmpresa)
    
    Set rsMoneda = objFormaPago.ListaMoneda(CodMonedaDefecto)
    arrEmpresa(15) = "" & rsMoneda("SMB").Value
    arrEmpresa(16) = "" & rsMoneda("LG").Value
    
    TipoCambio = objMoneda.TCambio(TipoCambioDefault, TipoCambioMonedaDefault)

    If Me.CodigoLocal = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_LOCAL_DELIVERY") Then
        blnEsDelivery = True
    Else
        blnEsDelivery = False
    End If

    If gclsOracle.Const_Val("CMR.PKG_USUARIO.P_VALIDA_PRECIO_ON_LINE") = "1" Then
        blnPrecioOnLine = True
    Else
        blnPrecioOnLine = False
    End If
    
    strTipoMaquinaCabina = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIPO_MAQUINA_CABINA")
    
    Dim k As Integer
    Set rsParametrosLocal = gclsOracle.FN_Cursor("BTLPROD.PKG_LOCAL.FN_LISTA_PARAMETROS", 0, strCia)
    While Not rsParametrosLocal.EOF
        xParametros.AppendRows
        xParametros(k, 0) = rsParametrosLocal("COD_PARAMETRO").Value
        xParametros(k, 1) = rsParametrosLocal("DES_PARAMETRO").Value
        xParametros(k, 2) = rsParametrosLocal("VAL_PARAMETRO").Value
        rsParametrosLocal.MoveNext
        k = k + 1
    Wend
    
    If Me.CodigoLocal = LocalDelivery And strFlgDelivery = "1" Then
        CodTipoVenta = TipoVentaDlv
    ElseIf Me.CodigoLocal <> LocalDelivery And strFlgDelivery = "0" Then
        CodTipoVenta = TipoVentaLocal
    Else
        CodTipoVenta = TipoVentaWeb
    End If
    
    Set rsInicioCaja = gclsOracle.FN_Cursor("BTLPROD.PKG_USUARIO.FN_LISTA_INICIO_CAJA", 0, strCia, strCodigo, strLocal, strNombreMaquina)
    If rsInicioCaja Is Nothing Then
        blnConectado = False
    Else
        strCodigoLiquidacion = "" & rsInicioCaja("COD_LIQUIDACION")
    End If
    
    Set objMaquina = Nothing
    
    Exit Function
Handle:
    gclsOracle.cerrar
    Err.Raise Err.Number, "clsUsuario.Login", Err.Description
End Function

Public Property Get Parametros(Optional CodigoParametro As String) As XArrayDB
Dim xTemp As New XArrayDB
xTemp.ReDim 0, 0, 0, 2
    If CodigoParametro = "" Then
        Set Parametros = Parametros
    Else
        Dim l As Integer
        l = xParametros.Find(0, 0, CodigoParametro)
        xTemp(0, 0) = xParametros(l, 0)
        xTemp(0, 1) = xParametros(l, 1)
        xTemp(0, 2) = xParametros(l, 2)
        Set Parametros = xTemp
    End If
End Property

Public Property Get Conectado()
    Conectado = blnConectado
End Property

Public Property Get Nombre()
    Nombre = strNombre
End Property

Public Property Get Perfil()
    Perfil = strPerfil
End Property

Private Sub Class_Initialize()
    blnConectado = False
    arrEmpresa = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    xParametros.ReDim 0, -1, 0, 2
End Sub

Public Property Get Aplicacion() As String
    Aplicacion = "14"
End Property

Public Property Get CodigoLocal() As String
    CodigoLocal = strLocal
End Property

Public Property Get CodLocalCallCenter() As String
    CodLocalCallCenter = strCodLocalCallCenter
End Property

Public Property Let CodLocalCallCenter(ByVal vdblMoneda As String)
    strCodLocalCallCenter = vdblMoneda
End Property

Public Property Get UbigeoLocal() As String
    UbigeoLocal = strUbigeoLocal
End Property

Public Property Get NombreLocal() As String
    NombreLocal = strNombreLocal
End Property

Public Property Get CentroCosto() As String
    CentroCosto = strCentroCosto
End Property

''Public Property Get DireccLocal() As String
''    DireccLocal = fnDevDireccLocal(objUsuario.CodigoLocal)
''End Property

Public Property Get Localidad() As String
    Localidad = fnDevLocalidad(objUsuario.CodigoLocal)
End Property

Public Property Get NombrePC() As String
    NombrePC = strNombreMaquina
End Property

Public Property Let NombrePC(ByVal lstrNombreMaquina As String)
    strNombreMaquina = lstrNombreMaquina
End Property

Public Property Get IGV() As Double
    IGV = gclsOracle.FN_Valor("BTLPROD.pkg_impuesto.fn_devuelve_impuesto", Me.CodigoEmpresa, Me.CodigoLocal, "")
    
End Property

Public Property Let TipoCambio(ByVal vdblMoneda As Double)
    dbltipocambio = vdblMoneda
End Property

Public Property Get TipoCambio() As Double
    TipoCambio = dbltipocambio 'arrEmpresa(17)
End Property

Public Function sysdate() As Date
    sysdate = gclsOracle.Fecha_Servidor
    'sysdate = Date
End Function

Public Property Get CodigoEmpresa() As String
    CodigoEmpresa = arrEmpresa(0)
End Property

Public Property Get Empresa() As String
    Empresa = arrEmpresa(2)
End Property

Public Property Get Siglas() As String
    Siglas = arrEmpresa(3)
End Property

Public Property Get Ruc() As String
    Ruc = arrEmpresa(4)
End Property

Public Property Get Tipo() As String
    Tipo = arrEmpresa(5)
End Property

Public Property Get direccion() As String
    direccion = arrEmpresa(6)
End Property

Public Property Get Distrito() As String
    Distrito = arrEmpresa(7)
End Property

Public Property Get Urbanizacion() As String
    Urbanizacion = arrEmpresa(8)
End Property

Public Property Get Provincia() As String
    Provincia = arrEmpresa(9)
End Property

Public Property Get Departamento() As String
    Departamento = arrEmpresa(10)
End Property

Public Property Get Pais() As String
    Pais = arrEmpresa(11)
End Property

Public Property Get Telefono1() As String
    Telefono1 = arrEmpresa(12)
End Property

Public Property Get Telefono2() As String
    Telefono2 = arrEmpresa(13)
End Property

Public Property Get ImprimirDireccion() As Boolean
    ImprimirDireccion = CBool("" & arrEmpresa(19))
End Property

Public Property Get Fax() As String
    Fax = arrEmpresa(14)
End Property

Public Function Lista(Optional CodUsuario As String = "", Optional cntCosto As String = "", Optional Estado As String = "") As oraDynaset
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_USUARIO.FN_LISTA", 0, CodUsuario, cntCosto, Estado)
End Function

Public Function Lista_No_Asign(Optional cntCosto As String, Optional Busqueda As String = "") As oraDynaset
    Set Lista_No_Asign = gclsOracle.FN_Cursor("BTLPROD.PKG_USUARIO.LISTA_NO_ASIGN", 0, cntCosto, Busqueda)
End Function

Public Function ListaCons(Optional CodUsuario As String = "", Optional cntCosto As String = "") As oraDynaset
    Set ListaCons = gclsOracle.FN_Cursor("BTLPROD.PKG_USUARIO.FN_LISTA_CONS", 0, CodUsuario, cntCosto)
End Function

Public Function ListaPermisosAdm(ByVal vstrUsuario As String, _
                                   ByVal vstrAplicacion As String) As oraDynaset
    On Error GoTo CntrlErr
    Set ListaPermisosAdm = gclsOracle.FN_Cursor("NUEVO.PKG_PERMISO.FN_LISTA_MENU", 0, vstrUsuario, vstrAplicacion)

    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsUsuario", Err.Description
End Function

Public Property Get PrecioOnLine() As Boolean
    PrecioOnLine = blnPrecioOnLine
End Property

Public Function ListaInicioCaja(ByVal vCodCia As String, ByVal vCodUsuario As String, ByVal vCodLocal As String, ByVal vCodMaquina As String) As oraDynaset
    On Error GoTo CntrlErr
    Set ListaInicioCaja = gclsOracle.FN_Cursor("BTLPROD.PKG_USUARIO.FN_LISTA_INICIO_CAJA", 0, vCodCia, vCodUsuario, vCodLocal, vCodMaquina)
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsUsuario.ListaInicioCaja", Err.Description
    
End Function

Public Function ListaMantenimientos(ByVal vstrUsuario As String, _
                                    ByVal vstrAplicacion As String) As oraDynaset
    On Error GoTo CntrlErr
    Set ListaMantenimientos = gclsOracle.FN_Cursor("NUEVO.PKG_PERMISO.FN_LISTA_MENU_MANT", 0, vstrUsuario, vstrAplicacion)
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsUsuario.ListaMantenimientos", Err.Description
End Function

'Public Property Get CodTipoVenta() As Variant
'
'    CodTipoVenta = strCodTipoVenta
'
'
'
'
'End Property

Public Property Get TipoMaquina() As String
    TipoMaquina = strTipoMaquina
End Property

Public Property Let TipoMaquina(ByVal vData As String)
    strTipoMaquina = vData
End Property


Public Property Get DireccionLocal() As String
    DireccionLocal = strDireccionLocal
End Property
''''cambiar por nombre local
Public Function fnDevLocalidad(ByVal vstrCodBtl As String) As String
    On Error GoTo CntrlErr
    fnDevLocalidad = gclsOracle.FN_Valor("BTLPROD.PKG_USUARIO.FN_DEV_LOCALIDAD", vstrCodBtl)
    
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsUsuario.ListaMantenimientos", Err.Description
End Function
''cambiar por direccion
''''Public Function fnDevDireccLocal(ByVal vstrCodBtl As String) As String
''''    On Error GoTo CntrlErr
''''    fnDevDireccLocal = gclsOracle.FN_Valor("BTLPROD.PKG_USUARIO.FN_DEV_DIRECC_LOCAL", vstrCodBtl)
''''
''''    Exit Function
''''CntrlErr:
''''    Err.Raise Err.Number, "clsUsuario.ListaMantenimientos", Err.Description
''''End Function

Public Property Get SmbMoneda() As Variant
    SmbMoneda = arrEmpresa(15)
End Property

Public Property Get DesLgMoneda() As Variant
    DesLgMoneda = arrEmpresa(16)
End Property

Public Property Get CodMonedaDefecto() As Variant
    CodMonedaDefecto = arrEmpresa(18)
End Property

Public Property Get LocalDelivery() As Variant
    LocalDelivery = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_LOCAL_DELIVERY")
End Property

Public Property Get EsDelivery() As Boolean
    EsDelivery = blnEsDelivery
End Property

Public Function Valor(ByVal CodParametro As String, ByVal CodCia As String) As String

On Error GoTo CtrlErr

        Valor = gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_CIA", CodParametro, CodCia)

    Exit Function
    
CtrlErr:

    Err.Raise Err.Number, "clsUsuario.ListaMantenimientos", Err.Description
End Function

Public Function MostrarStock(ByVal CodPerfil As String) As String
On Error GoTo CtrlErr

    MostrarStock = gclsOracle.FN_Valor("CMR.PKG_USUARIO.FN_MOSTRAR_STOCK", CodPerfil)

Exit Function
CtrlErr:

    Err.Raise Err.Number, "clsUsuario.MostrarStock", Err.Description
End Function

Public Property Get TipDocDefault() As String
    TipDocDefault = strTipDocDefault
End Property

Public Property Get TipoDocGuia() As Variant
    TipoDocGuia = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIP_DOC_GUIA_LOC")
End Property

Public Property Get TipoDocNC() As Variant
    TipoDocNC = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIP_DOC_NOTA")
End Property

Public Property Get ModalidadVentaConvenio() As Variant
    ModalidadVentaConvenio = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_MODAL_VTA_002_CONVEMP")
End Property

Public Property Get LabelImpConvGuia01() As Variant
    LabelImpConvGuia01 = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_LABEL_IMP_CONV_GUIA_01")
End Property

Public Property Get LabelImpConvGuia02() As Variant
    LabelImpConvGuia02 = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_LABEL_IMP_CONV_GUIA_02")
End Property

Public Property Get MotivoGeneraGuiaLocal() As Variant
    MotivoGeneraGuiaLocal = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_MOTIVO_GENERA_GUIA_LOCAL")
End Property

Public Property Get TipoCambioDefault() As Variant
    TipoCambioDefault = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIPO_CAMBIO_DEFAULT")
End Property

Public Property Get TipoCambioMonedaDefault() As Variant
    TipoCambioMonedaDefault = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIPO_CAMBIO_MON_DEFAULT")
End Property

Public Property Get EstadoAnulado() As Variant
    EstadoAnulado = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_EST_ANULADO")
End Property

Public Property Get EstadoEmitido() As Variant
    EstadoEmitido = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_EST_EMITIDO")
End Property

Public Property Get PedidoAvisado() As Variant
    PedidoAvisado = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PEDIDO_AVISADO")
End Property

Public Property Get PedidoProformado() As Variant
    PedidoProformado = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PEDIDO_PROFORMA")
End Property

Public Property Get PedidoVerificado() As Variant
    PedidoVerificado = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PEDIDO_VERIFICADO")
End Property

Public Property Get TipoDocBol() As Variant
    TipoDocBol = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIP_DOC_BOL")
End Property

Public Property Get TipoDocFac() As Variant
    TipoDocFac = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIP_DOC_FAC")
End Property

Public Property Get ValParametroSinLote() As Variant
    ValParametroSinLote = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_VAL_PARAMETRO_SIN_LOTE")
End Property

Public Property Get ValParametroSinFecha() As Variant
    ValParametroSinFecha = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_VAL_PARAMETRO_SIN_FECHA")
End Property

Public Property Get TipoDocRecibo() As Variant
    TipoDocRecibo = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIP_DOC_RECIBO")
End Property

Public Property Get PrecMinRM() As String
    PrecMinRM = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PREVTA_RM")
End Property

Public Property Get TipoMaquinaAdmin() As String
    TipoMaquinaAdmin = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIPO_MAQUINA_ADMIN")
End Property

Public Property Get TipoMaquinaCajero() As String
    TipoMaquinaCajero = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIPO_MAQUINA_CAJERO")
End Property

Public Property Get TipoMaquinaCabina() As String
    TipoMaquinaCabina = strTipoMaquinaCabina
End Property

Public Property Get TipoMaquinaRuteo() As String
    TipoMaquinaRuteo = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIPO_MAQUINA_RUTEO")
End Property

Public Property Get TipoMaquinaVerif() As String
    TipoMaquinaVerif = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIPO_MAQUINA_VERIF")
End Property

Public Property Get TipoVentaLocal() As String
    TipoVentaLocal = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIPO_VENTA_LOCAL")
End Property

Public Property Get TipoVentaDlv() As String
    TipoVentaDlv = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIPO_VENTA_DLV")
End Property

Public Property Get TipoVentaWeb() As String
    TipoVentaWeb = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIPO_VENTA_WEB")
End Property

Public Property Get FormaPagoEfe() As String
    FormaPagoEfe = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_FORMA_PAGO_EFE")
End Property

Public Property Get CodigoAplicacion() As String
    CodigoAplicacion = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_COD_APLICACION")
End Property

Public Property Get CodigoMenuAsigna() As String
    CodigoMenuAsigna = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_COD_MENU_ASIGNA_META_DLV")
End Property

Public Property Get FormaPagoEfeSoles() As String
    FormaPagoEfeSoles = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_FORMA_PAGO_EFE_SOLES")
End Property

Public Property Get CodTipoVenta() As String
    CodTipoVenta = strCodTipoVenta
    
End Property

Public Property Let CodTipoVenta(ByVal lstrCodTipoVenta As String)
    strCodTipoVenta = lstrCodTipoVenta
    
End Property

Public Property Get CODIGO() As String
    CODIGO = strCodigo
End Property

Public Property Let CODIGO(ByVal lstrCodigo As String)
    strCodigo = lstrCodigo
End Property

Public Function ListaUsuarioDLV() As oraDynaset
    On Error GoTo CtrlErr
    Set ListaUsuarioDLV = gclsOracle.FN_Cursor("BTLPROD.PKG_USUARIO.LISTA_USUARIOS_DLV", 0)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsUsuario.ListaUsuarioDLV", Err.Description
End Function

Public Function AsignaMetaDLV(ByVal vstrCodAplicacion As String, _
                              ByVal vstrCodMenu As String, _
                              ByVal vstrCodUsuario As String) As String
    On Error GoTo CtrlErr
    AsignaMetaDLV = gclsOracle.FN_Valor("BTLPROD.PKG_USUARIO.USUARIO_ASIGNA_META", vstrCodAplicacion, _
                                                                                       vstrCodMenu, _
                                                                                       vstrCodUsuario)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsUsuario.AsignaMetaDLV", Err.Description
End Function

Public Function ListaUsuarioDLVRuta() As oraDynaset
    On Error GoTo CtrlErr
    Set ListaUsuarioDLVRuta = gclsOracle.FN_Cursor("BTLPROD.PKG_USUARIO.LISTA_USUARIOS_DLV_RUTA", 0)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsUsuario.ListaUsuarioDLVRuta", Err.Description
End Function

Public Property Get Password() As String
    Password = strPassword
End Property

Public Property Let Password(ByVal lstrPassword As String)
    strPassword = lstrPassword
End Property

Public Property Get flgDeliveryProv() As String
    flgDeliveryProv = strFlgDeliveryProv
End Property

Public Property Let flgDeliveryProv(ByVal lstrFlgDeliveryProv As String)
    strFlgDeliveryProv = lstrFlgDeliveryProv
End Property

Public Property Get CodConvenioSenati() As String
    CodConvenioSenati = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_CONVENIO_SENATI")
End Property

Public Property Get CodConvenioScotiaBank() As String
    CodConvenioScotiaBank = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_CONVENIO_SCOTIABANK")
End Property

Public Property Get CodigoMotorizado() As String
    CodigoMotorizado = strCodMotorizado
End Property

Public Property Let CodigoMotorizado(ByVal lstrCodMotorizado As String)
    strCodMotorizado = lstrCodMotorizado
End Property

Public Function DevMotorizado(ByVal vstrCia As String, _
                              ByVal vstrCodLocal As String, _
                              ByVal vstrNumProforma As String) As String
    On Error GoTo CtrlErr
    DevMotorizado = "" & gclsOracle.FN_Valor("BTLPROD.PKG_MOTORIZADO.LISTA_MOTORIZADO_PEDIDO", vstrCia, vstrCodLocal, vstrNumProforma)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMotorizado.DevMotorizado"
End Function

Public Function GrabaEstadoMotorizado(ByVal vstrCodMotorizado As String, _
                                      ByVal vstrCodEstadoMotorizado As String, _
                                      ByVal vstrCodUsuario As String)
   On Error GoTo CtrlErr
   
   Dim gvarValores As Variant
   Dim gvarIO  As Variant
   
   gvarValores = Array(vstrCodMotorizado, vstrCodEstadoMotorizado, _
                       vstrCodUsuario)
                          
   gvarIO = Array(entrada, entrada, _
                  entrada)
                  
   GrabaEstadoMotorizado = gclsOracle.SP("BTLPROD.PKG_MOTORIZADO.SP_GRABA_ESTADO_MOTORIZADO", _
                                    gvarValores, _
                                    gvarIO)
                  
   Exit Function
CtrlErr:
   Err.Raise Err.Number, "clsUsuario.GrabaEstadoMotorizado", Err.Description
End Function

Public Function EsQuimico(ByVal CodigoUsuario As String) As Boolean
    On Error GoTo CtrlErr
    EsQuimico = IIf("" & gclsOracle.FN_Valor("BTLPROD.PKG_USUARIO.FN_ES_QUIMICO", CodigoUsuario) = "1", True, False)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsUsuario.EsQuimico"
End Function

Public Function EsTecnico(ByVal CodigoUsuario As String) As Boolean
    On Error GoTo CtrlErr
    EsTecnico = IIf("" & gclsOracle.FN_Valor("BTLPROD.PKG_USUARIO.FN_ES_TECNICO", CodigoUsuario) = "1", True, False)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsUsuario.EsQuimico"
End Function

Public Function PuedeVerStock(ByVal CodigoUsuario As String) As Boolean
    On Error GoTo CtrlErr
    PuedeVerStock = IIf("" & gclsOracle.FN_Valor("BTLPROD.PKG_USUARIO.FN_PUEDE_VER_STOCK", CodigoUsuario) = "1", True, False)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsUsuario.PuedeVerStock"
End Function

Public Function UsuarioEsDelLocal(ByVal CodLocal As String, ByVal usuario As String, ByVal Perfil As String) As Integer
    On Error GoTo CtrlErr
    
    UsuarioEsDelLocal = gclsOracle.FN_Valor("CMR.PKG_USUARIO.FN_USUARIO_ES_DEL_LOCAL", CodLocal, usuario, Perfil)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsUsuario.UsuarioEsDelLocal", Err.Description

End Function

Public Function NumeroDelivery() As String
                       
On Error GoTo CtrlErr
    NumeroDelivery = gclsOracle.FN_Valor("BTLPROD.PKG_LOCAL.FN_DEV_NUMDELIVERY", Me.CodigoLocal)
    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsUsuario.NumeroDelivery", Err.Description
End Function

Private Function MayoristaFracciones() As Boolean
                       
On Error GoTo CtrlErr
    MayoristaFracciones = IIf(gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_VALIDA_FRAC_VTA_MAYORISTA") = 1, True, False)
    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsUsuario.MayoristaFracciones", Err.Description
End Function

Public Property Get MayoristaFracciona() As Boolean
    MayoristaFracciona = strMayoristaFracciones
End Property

Public Property Get CodigoLiquidacion() As String
    CodigoLiquidacion = strCodigoLiquidacion
End Property

Public Function ValidaVersion() As Boolean
                       
On Error GoTo CtrlErr
    ValidaVersion = IIf(gclsOracle.FN_Valor("NUEVO.PKG_APLICACION.FN_VALIDA_VERSION", Me.Aplicacion, gstrVersion) = 1, True, False)
    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsUsuario.ValidaVersion", Err.Description
End Function

Public Function GrabaVisita(ByVal NombreFormularios As String, _
                            ByVal CodigoMenu As String)
   On Error GoTo CtrlErr
   
   Dim gvarValores As Variant
   Dim gvarIO  As Variant
   
   gvarValores = Array(objUsuario.CODIGO, NombreFormularios, Obtener(12), objUsuario.CodigoAplicacion, CodigoMenu)
                          
   gvarIO = Array(entrada, entrada, entrada, entrada, entrada)
                  
   GrabaVisita = gclsOracle.SP("BTLPROD.SP_GRABA_LOG_VISITA", gvarValores, gvarIO)
                  
   Exit Function
CtrlErr:
   Err.Raise Err.Number, "clsUsuario.GrabaVisita", Err.Description
End Function

Public Function EvocarContraseña(ByVal vstrCodUsuario As String, _
                                 ByVal vstrNumDNI As String, _
                                 ByVal vstrFchNacimiento As String) As String
   On Error GoTo CtrlErr

   Dim miconexion As New clsOracle

   miconexion.Conexion gvarTNSNAME, "GENERAL", "GENERAL"
   
   EvocarContraseña = miconexion.FN_Valor("BTLPROD.PKG_LOGUEO.FN_EVOCAR_PASSWORD", _
                                          vstrCodUsuario, _
                                          vstrNumDNI, _
                                          vstrFchNacimiento)
    miconexion.cerrar
    Set miconexion = Nothing
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsUsuario.EvocarContraseña", Err.Description
End Function

Public Function GrabaContraseña(ByVal vstrCodUsuario As String, _
                                ByVal vstrPasswordActual As String, _
                                ByVal vstrPasswordNew As String) As String
   On Error GoTo CtrlErr
    '------------------------------
   Dim miconexion As New clsOracle
    '------------------------------
   Dim gvarValores As Variant
   Dim gvarIO  As Variant

   '----------------------------------------------------
   miconexion.Conexion gvarTNSNAME, "GENERAL", "GENERAL"
   '----------------------------------------------------

   gvarValores = Array(vstrCodUsuario, vstrPasswordActual, vstrPasswordNew)

   gvarIO = Array(entrada, entrada, entrada)

   GrabaContraseña = miconexion.SP("BTLPROD.PKG_LOGUEO.SP_NUEVO_PASSWORD", gvarValores, gvarIO)

   '----------------------------------------------------
   miconexion.cerrar
   Set miconexion = Nothing
   '----------------------------------------------------

   Exit Function
CtrlErr:
   Err.Raise Err.Number, "clsUsuario.GrabaContraseña", Err.Description
End Function

Public Property Get MetaValePromedio() As String
        MetaValePromedio = strMetaValePromedio
End Property

Public Property Get MetaNumeroCliente() As String
        MetaNumeroCliente = strMetaNumeroCliente
End Property

Private Sub TRAEMETA()
On Error GoTo Handle
If "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "MUESMETA") = "1" Then
    strMetaValePromedio = gclsOracle.FN_Valor("BTLPROD.PKG_REPO_GESTION.FN_DEV_META", "015", Format(objUsuario.sysdate, "YYYY"), Format(objUsuario.sysdate, "MM"), objUsuario.CodigoLocal)
    strMetaNumeroCliente = gclsOracle.FN_Valor("BTLPROD.PKG_REPO_GESTION.FN_DEV_META", "016", Format(objUsuario.sysdate, "YYYY"), Format(objUsuario.sysdate, "MM"), objUsuario.CodigoLocal)
End If
Exit Sub
Handle:
strMetaValePromedio = 0
strMetaNumeroCliente = 0
End Sub

Public Function GrabaLogBusqueda(ByVal strFlag As String, ByVal strFiltroBuscado As String) As String
On Error GoTo CtrlErr

Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(objUsuario.CODIGO, _
                        objUsuario.Nombre, _
                        objUsuario.NombrePC, _
                        objUsuario.NombreLocal, _
                        objUsuario.Perfil, _
                        objUsuario.CodigoLocal, _
                        strFlag, _
                        strFiltroBuscado)
    
    gvarIO = Array(entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada)
                   
    GrabaLogBusqueda = gclsOracle.SP("BTLPROD.PKG_LOG_BUSQUEDA.SP_INSERTLOG", _
                              gvarValores, _
                              gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsUsuario", Err.Description
End Function


Public Function fnValidaCotiza(ByVal vstrCodUsuario As String, ByVal vstrPASSWORD As String, ByVal vstrCodAplicacion As String) As oraDynaset
    On Error GoTo CtrlErr
    Set fnValidaCotiza = gclsOracle.FN_Cursor("BTLPROD.PKG_USUARIO.FN_VALIDA_COTIZA", 0, vstrCodUsuario, vstrPASSWORD, vstrCodAplicacion)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsUsuario.fnValidaCotiza", Err.Description
End Function

Public Function ListaEmpresa(ByVal vstrCia As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaEmpresa = gclsOracle.FN_Cursor("BTLPROD.PKG_EMPRESA.FN_LISTA", 0, vstrCia)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsUsuario.ListaEmpresa", Err.Description
End Function


