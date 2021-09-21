Attribute VB_Name = "mdlPrincipal"
Option Explicit
Global gblNumeroTelefonico As String
Global m_objPedido As clsPedido
Global m_tipo_maquina As String
Global callcenter As String
Global codModalidadVentaBK As String
Global bk_usuario As String
Global bk_password As String
'Public Enum eTipoModalidad
'    VentaRegular = 1
'    VentaConvenio = 2
'    VentaMayorista = 3
'    CobroRespon = 4
'    Canje = 5
'    Servicios = 6
'    CobranzaVtaCred = 7
'    Cotizacion = 8
'    Fact_Servicios = 9
'    Guias = 10
'    Magistral = 11
'End Enum

'Public ptmModalidad As eTipoModalidad
Public intGrabadoPedido As Integer 'INDICA SI EXISTEN PEDIDOS DESPACHADOS


'--------------------------------------
'Tipo Precio
'--------------------------------------

Public Enum eTipoPrecio
    regular = 1
    Mayorista = 2
    Convenio = 1
    'Reg_Mag = 3
End Enum
Public ptmTipoPrecio As eTipoPrecio

'--------------------------------------'
'--- Ventana Activa ---'
Public Enum eVentanaCli
    Documento1 = 1
    Recetario_Magistral = 2
End Enum
Public penumVentCli As eVentanaCli
'--------------------------------------'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\'
'--------------------------------------'
'--- Transacciones a correr ---'
Public Enum eTransacc
    GbDocumento = 1
    GbProforma = 2
End Enum
Public peTransacc As eTransacc
'--------------------------------------'

'--- Producto es regalo o Normal ---'
Public Enum eProdRegalo
    Producto_Regalo = 1
    Producto_Normal = 0
    Producto_Regalo_precio = 2
End Enum
Public pProdRegalo As eProdRegalo

'Tipo del valor que se guarda en Valor
Public Enum eTipoValor
        Valor_Porcentaje = 0
        Valor_Importe = 1
End Enum
Public pTipoValor As eTipoValor

'Tipo de documento Monedero
Public Enum ETipoDocumentoMonedero
    eeTarjetaMonedero
    eeDniCliente
End Enum

'API Types
Public Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

#If UNICODE Then
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal uMgs As Long, ByVal wParam As Long, lParam As Any) As Long
#Else
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal uMgs As Long, ByVal wParam As Long, lParam As Any) As Long
#End If
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Any) As Long
'the standard version of this API function uses a Point structure, but we cant pass
'that using VB, so it has been modified to accept Long Integers
'Public Declare Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal lLeft As Long, ByVal lTop As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal HwndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
'////////////////////////////////////
' para abrir con el shell
Private Declare Function ShellExecute Lib "Shell32.Dll" Alias "ShellExecuteA" ( _
ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

'///////////////////////////////////
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
 
Public Const SWW_HPARENT = -8
Public Const HTRIGHT = 11
Public Const WM_NCLBUTTONDOWN = &HA1

'para encriptacion de texto
Private oTest As CRijndael
Private Const pEncryptionKey As String = "thekeymaker"
Public Const pdblColorFondo As Double = &H80000018
Public blnEnviaMensajeDelivery As Boolean
Public Const pstrPuerto As String = "8758"
Public pstrPaginaComunicandonos As String
Dim KeyTime As Double
Dim ACABA As Boolean

Public Sub psub_KeyDownAplicacion(ByVal KeyCode As Integer, ByVal Shift As Integer)
    Dim tempCtrl  As Boolean, tempAlt As Boolean
    tempCtrl = (Shift And vbCtrlMask) > 0
    tempAlt = (Shift And vbAltMask) > 0
    'Controla que las teclas de funcion esten desabilitadas mientras este en load el frm de Cantidad'

    If frm_VTA_CantidadProducto.pBlnModal = True Then Exit Sub
        Select Case KeyCode
            Case vbKeyF5 And mdiPrincipal.cmdDocumento.Enabled
                mdiPrincipal.cmdDocumento_Click
'                 If objVenta.CodigoTipoVenta = Guias_Remision Then
'                    frm_VTA_GuiaRemision.Show
'                    frm_VTA_GuiaRemision.SetFocus
'                Else
'                    frm_VTA_Documento.Show
'                    frm_VTA_Documento.SetFocus
'                End If
            Case vbKeyF6 And mdiPrincipal.cmdFormaPago.Enabled
                'frm_VTA_FormaPago.Show
                'frm_VTA_FormaPago.SetFocus
                mdiPrincipal.cmdFormaPago_Click
            Case vbKeyF7 And mdiPrincipal.cmdGrabaVenta.Enabled
                mdiPrincipal.cmdGrabaVenta_Click
            Case vbKeyF8 And mdiPrincipal.cmdProforma.Enabled
                mdiPrincipal.cmdProforma_Click
                
            Case tempCtrl And vbKeyQ
                frm_VTA_Modalidad.Show vbModal
            Case tempCtrl And vbKeyE
                frm_VTA_Busqueda.Show
                frm_VTA_Busqueda.SetFocus
            Case tempCtrl And vbKeyH
                frmPedido.abre
            Case tempCtrl And vbKeyR
'             If objVenta.CodigoTipoVenta = Guias_Remision Then
'                frm_VTA_GuiaRemision.Show
'                frm_VTA_GuiaRemision.SetFocus
'            Else
'                frm_VTA_Documento.Show
'                frm_VTA_Documento.SetFocus
'            End If
            Case tempCtrl And vbKeyA
'                frm_VTA_FormaPago.Show
'                frm_VTA_FormaPago.SetFocus
            Case tempCtrl And vbKeyW
            Case tempAlt And vbKeyW
            If mdiPrincipal.txtDLVTelefono.Visible = True Then
                mdiPrincipal.txtDLVTelefono.SetFocus
            End If
            'Case tempCtrl And vbKeyS
       
            Case tempCtrl And vbKeyD And mdiPrincipal.cmdAdministrador.Enabled
                mdiPrincipal.cmdAdministrador_Click
    ''           If (objUsuario.Perfil = "0700") Or (objUsuario.Perfil = "0709") Or (objUsuario.Perfil = "0710") Then
    ''             frm_VTA_Administrador.Show
    ''           Else
    ''             MsgBox "No esta Autorizado por no ser QF", vbCritical, "No autorizado": Exit Sub
    ''           End If
            
            Case tempCtrl And vbKeyF
                psub_Cerrar
            Case tempCtrl And vbKeyM And mdiPrincipal.cmdMantenimientos.Enabled
                'mdiPrincipal.cmdAdministrador_Click
                mdiPrincipal.cmdMantenimientos_Click
                'frm_VTA_Mantenimientos.Show
                'frm_VTA_Mantenimientos.SetFocus
            Case tempCtrl And vbKeyX
                mdiPrincipal.cmdImpresion_Click
            Case tempCtrl And vbKeyC
''                If MsgBox("¿ Desea Borrar todos los Datos del Documento.. ?", vbQuestion + vbYesNo, "Aviso => Se Borrara la Información") = vbYes Then
''                    frmPedido.psub_BeginArry
''                    mdiPrincipal.subNuevo
''                    frm_VTA_Busqueda.grdProductos.Limpiar
''                    frm_VTA_Busqueda.grdAlternativos.Limpiar
''                    frm_VTA_Busqueda.grdComplementarios.Limpiar
''                    frm_VTA_Busqueda.txtBuscar.selection
''                     frm_VTA_Modalidad.Show vbModal
''                 Else
''                    Exit Sub
''                End If

                mdiPrincipal.cmdCancelar_Click
            Case tempCtrl And vbKeyG
'                mdiPrincipal.cmdGrabaVenta_Click
            Case tempCtrl And vbKeyS   'Pedido
                'visible o no el ingreso de cliente (nuevo)
                Dim rs As oraDynaset
                Dim v_Bool As Boolean
                v_Bool = objVenta.MuestraFidelizado("ACTDIDELIZ")
                If v_Bool Then
                    frmPedido_Busca_Cli.Show vbModal
                    'frmPedido.optEfectivo.Value = True
                    frmPedido.flgF6 = 1
                End If
                Set rs = Nothing
        End Select
End Sub

Sub Main()
    gstrConexion = "Provider=MSPersist;"
    xProductoRegaloBK.ReDim 0, -1, 0, 29 'Inicializa los productos regalos seleccionados
    strUsuariosXML = App.Path & IIf(right(App.Path, 1) = "\", "", "\") & "usuarios.xml"
    strPreciosXML = App.Path & IIf(right(App.Path, 1) = "\", "", "\") & "precios.xml"
    strDetalleVentaXML = App.Path & IIf(right(App.Path, 1) = "\", "", "\") & "detalleventa.xml"
    strPagoVentaXML = App.Path & IIf(right(App.Path, 1) = "\", "", "\") & "pagoventa.xml"

    gstrIni = App.Path & IIf(right(App.Path, 1) = "\", "", "\") & "contingencia.ini"

    Dim objArchivoIni As New cls_ArchivoIni
    gvarHabilitadoContingencia = objArchivoIni.LeerIni(gstrIni, "general", "FLG_CONTINGENCIA")
    Set objArchivoIni = Nothing

    'gvarTNSNAME = "TESTING"
    'gvarTNSNAME = "DESA"
    'gvarTNSNAME = "RACDB2"
    gvarTNSNAME = "BTLRAC"
    'gvarTNSNAME = "XEFARMA"
    'gvarTNSNAME = "DESA_RACDB"
    'gvarTNSNAME = "RACDBPTO"

    Dim ConexForzada As String
'    Dim example_command As String
'    Dim devParametros As String
'    devParametros = InputBox("Ingrese Ambiente", App.ProductName)
'    example_command = devParametros
'    example_command = "-DESA -003 -1DLV"
''    m_tipo_maquina = fncGuion(fncGuion(example_command, 1, "-"), 1, "-") 'Command$
    ConexForzada = fncGuion(fncGuion(Command$, 1, "-"), 0, "-")
    m_tipo_maquina = separaCadena(fncGuion(Command$, 1, "-"), "-", 1)
    callcenter = separaCadena(fncGuion(Command$, 1, "-"), "-", 2)
    
    'MsgBox "tsname: " & ConexForzada & Chr(13) & "m_tipo_maquina: " & m_tipo_maquina & Chr(13) & "callcent: " & callcenter
    
    gvarTNSNAME = IIf(ConexForzada = "", gvarTNSNAME, ConexForzada)
    
    gstrAplicacion = App.FileDescription
    gstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    
    garrCallGoogleMaps.ReDim 0, -1, 0, 10
    If gvarHabilitadoContingencia = "1" Then
        OFF_Main
    Else
        frm_VTA_Logueo.Caption = frm_VTA_Logueo.Caption & " " & gstrAplicacion & " * Ver: " & gstrVersion & " - " & gvarTNSNAME
        frm_VTA_Logueo.Show
    End If
End Sub
Public Function separaCadena(Cadena, Optional separador$ = "|", Optional elemento) As String
On Error Resume Next
    Dim obj
    Dim Count
    obj = Split(Cadena, separador)
    Count = UBound(obj)
    If Count < 0 Then Exit Function
    separaCadena = Trim(obj(elemento))
End Function
Function MainX(ByVal Usuario As String, ByVal Password As String) As Boolean
    bk_usuario = Usuario
    bk_password = Password
    Screen.MousePointer = vbHourglass

    Ver.dwOSVersionInfoSize = Len(Ver)

    GetVersionEx Ver
    
    On Error GoTo ErrorConexion
    gvarUSUARIO = "CONECT"
    gvarPASSWORD = "CONECT"

    Dim strSQL$
    If gclsOracle.Conexion(gvarTNSNAME, gvarUSUARIO, gvarPASSWORD) <> "" Then Exit Function
    '''gvarUSUARIO = "BTLPROD"
    ''gvarUSUARIO = "VENTAS"
    gvarUSUARIO = gclsOracle.FN_Valor("BTLPROD.PKG_USUARIO.FN_USUARIO_SIS", Usuario)
    
    If gclsOracle.FN_Valor("BTLPROD.PKG_USUARIO.FN_ES_GENERICO", Usuario) = "1" Then
        gvarPASSWORD = gvarUSUARIO
    Else
        gvarPASSWORD = Password
    End If
    gclsOracle.Cerrar
    If gclsOracle.Conexion(gvarTNSNAME, gvarUSUARIO, gvarPASSWORD) <> "" Then Exit Function

    gclsOracle.Execute "BEGIN DBMS_APPLICATION_INFO.SET_MODULE('" & App.FileDescription & "','" & App.Major & "." & App.Minor & "." & App.Revision & "'); END ;"
    
    Set gosesVentas = gclsOracle.OSession
    Set godbVentas = gclsOracle.ODataBase
    
    '******************************************************************************************************************'
    '*** Validando que cuando sea Delivery levante mas de una session ***'
    '*** Cambio Hecho el 01/10/2007 Por Crueda                        ***'
    '******************************************************************************************************************'
    Dim odynLocal As oraDynaset
    Dim strLocal As String
    Dim strDlv As String
    Dim dblCnt As Integer
    Dim i As Integer
    
    Set odynLocal = gclsOracle.FN_Cursor("BTLPROD.PKG_MAQUINA.FN_LISTA", 0, Replace(Trim(Obtener(12)), Chr(0), ""))
    strLocal = "" & odynLocal("COD_LOCAL").Value
    
    strDlv = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_LOCAL_DELIVERY")
    
    If Trim(strDlv) <> Trim(strLocal) Then
        If App.PrevInstance Then
            Screen.MousePointer = vbDefault
            MsgBox "La Aplicación Ya Esta Siendo Ejecutada", vbInformation, App.ProductName
            End
        End If
    End If
    '******************************************************************************************************************'
    '******************************************************************************************************************'
    gintDec = gclsOracle.Const_Val("CMR.PK_COM_PRECIOS.V_DECIMALES")
    gintDecTot = gclsOracle.Const_Val("CMR.PK_COM_PRECIOS.V_DECIMALES_TOT")
    fMapearUnidadDeRed
    gstrPathLog = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "PATHDLVLOG", "10")
    '----------------------------
    gstrFlagLogFile1 = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGLGDLV01") 'web service
    gstrFlagLogFile2 = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGLGDLV02") 'oracle
    gstrFlagLogFile3 = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGLGDLV03") 'gmaps
    gstrFlagLogBD1 = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGLGDLV08") 'web service
    gstrFlagLogBD2 = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGLGDLV09") 'oracle
    gstrFlagLogBD3 = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGLGDLV00") 'gmaps
    '-----------------------------
    gstrFlagSufDir = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGSFJDR01")
    gstrFlagValRut = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGVALRT01")
    gstrFlagReclamo = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRCLDLV1")
    gstrFlagReservaCap = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVCAP")
    gstrIndRAv3 = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRUAUV3")
    gstrIndRAv4 = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRUAUV4")
    gstrIndDCSAP = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGDCAQPV1")
    gstrIndCreaLogError = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGLGDLV10")
    
    gstrSpecialChrsToWS0 = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "SPCHRTOWS0", "10") 'CARACTERES A QUITAR, REGULAR
    gstrSpecialChrsToWS1 = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "SPCHRTOWS1", "10") 'CARACTERES A QUITAR, DC SAP
    
    gstrVarURL1 = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "URLWS201", "10")
    If gstrVarURL1 = "" Then
          Err.Raise vbObjectError, "mdlPrincipal.MainX", "No se encontro la url de la api google maps (URLWS201) para la CIA 10"
    End If
    gstrVarURL2 = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "URLWS202", "10")
    If gstrVarURL2 = "" Then
          Err.Raise vbObjectError, "mdlPrincipal.MainX", "No se encontro la url de la api google maps (URLWS202) para la CIA 10"
    End If
    gstrVarURL3 = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "URLWS203", "10")
    If gstrVarURL3 = "" Then
          Err.Raise vbObjectError, "mdlPrincipal.MainX", "No se encontro la url de la api google maps (URLWS203) para la CIA 10"
    End If
    gstrVarKEY1 = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "KEYWS200", "10")
    If gstrVarKEY1 = "" Then
          Err.Raise vbObjectError, "mdlPrincipal.MainX", "No se encontro la KEY de la api google maps (KEYWS200) para la CIA 10"
    End If
    gintFilesLog = gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "CNTLOGSAVE")
    On Error GoTo 0
    GoTo Final
ErrorConexion:
    Screen.MousePointer = vbDefault
    MsgBox "Existe un Problema con la Conexión" & Chr(13) & "Error :" & Err.Description, vbCritical, App.ProductName
    Exit Function
Final:
    MainX = True
    Screen.MousePointer = vbDefault
    
End Function

Function MainInkaClub() As Boolean
    On Error GoTo ErrorConexion
    gvarTNSNAME2 = IIf(Len(Trim(gvarTNSNAME2)) = 0, "EDBST008", gvarTNSNAME2)
    gvarUSUARIO2 = "ecventa"
    gvarPASSWORD2 = "venta"
    
    Dim strSQL$
    If gclsOracle.Conexion(gvarTNSNAME2, gvarUSUARIO2, gvarPASSWORD2) <> "" Then Exit Function
    'gvarUSUARIO = gclsOracle.FN_Valor("BTLPROD.PKG_USUARIO.FN_USUARIO_SIS", Usuario)
    
    'If gclsOracle.FN_Valor("BTLPROD.PKG_USUARIO.FN_ES_GENERICO", Usuario) = "1" Then
    '    gvarPASSWORD = gvarUSUARIO
    'Else
    '    gvarPASSWORD = Password
    'End If
    'gclsOracle.Cerrar
    'If gclsOracle.Conexion(gvarTNSNAME, gvarUSUARIO, gvarPASSWORD) <> "" Then Exit Function

    gclsOracle.Execute "BEGIN DBMS_APPLICATION_INFO.SET_MODULE('" & App.FileDescription & "','" & App.Major & "." & App.Minor & "." & App.Revision & "'); END ;"
    
    Set gosesVentas2 = gclsOracle.OSession
    Set godbVentas2 = gclsOracle.ODataBase
    
    '******************************************************************************************************************'
    '*** Validando que cuando sea Delivery levante mas de una session ***'
    '*** Cambio Hecho el 01/10/2007 Por Crueda                        ***'
    '******************************************************************************************************************'
'    Dim odynLocal As oraDynaset
'    Dim strLocal As String
'    Dim strDlv As String
'    Dim dblCnt As Integer
'    Dim i As Integer
'
'    Set odynLocal = gclsOracle.FN_Cursor("BTLPROD.PKG_MAQUINA.FN_LISTA", 0, Replace(Trim(Obtener(12)), Chr(0), ""))
'    strLocal = "" & odynLocal("COD_LOCAL").Value
'
'    strDlv = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_LOCAL_DELIVERY")
'
'    If Trim(strDlv) <> Trim(strLocal) Then
'        If App.PrevInstance Then
'            Screen.MousePointer = vbDefault
'            MsgBox "La Aplicación Ya Esta Siendo Ejecutada", vbInformation, App.ProductName
'            End
'        End If
'    End If
    '******************************************************************************************************************'
    '******************************************************************************************************************'
'    gintDec = gclsOracle.Const_Val("CMR.PK_COM_PRECIOS.V_DECIMALES")
'    gintDecTot = gclsOracle.Const_Val("CMR.PK_COM_PRECIOS.V_DECIMALES_TOT")
'    fMapearUnidadDeRed
    On Error GoTo 0
    GoTo Final
ErrorConexion:
    Screen.MousePointer = vbDefault
    MsgBox "Existe un Problema con la Conexión" & Chr(13) & "Error :" & Err.Description, vbCritical, App.ProductName
    Exit Function
Final:
    MainInkaClub = True
    Screen.MousePointer = vbDefault
End Function
''''''''''''Sub MainX()
''''''''''''
''''''''''''
''''''''''''
''''''''''''    Screen.MousePointer = vbHourglass
''''''''''''
''''''''''''    Ver.dwOSVersionInfoSize = Len(Ver)
''''''''''''
''''''''''''
''''''''''''    GetVersionEx Ver
''''''''''''
''''''''''''    'MsgBox Ver.dwPlatformId & Chr(13) & Ver.dwMajorVersion
''''''''''''
''''''''''''
''''''''''''
''''''''''''
''''''''''''
'''''''''''''    If Not objUsuario.EsDelivery Then
'''''''''''''        'Dim gstrInstancia As frm_VTA_Logueo
'''''''''''''        If App.PrevInstance Then
'''''''''''''            Screen.MousePointer = vbDefault
'''''''''''''            MsgBox "La Aplicación Ya Esta Siendo Ejecutada", vbInformation, App.ProductName
'''''''''''''            End
'''''''''''''        End If
'''''''''''''    End If
''''''''''''
''''''''''''    On Error GoTo ErrorConexion
''''''''''''    gvarUSUARIO = "GENERAL"
''''''''''''    gvarPASSWORD = "GENERAL"
''''''''''''
''''''''''''    Dim StrSql$
''''''''''''    If gclsOracle.Conexion(gvarTNSNAME, gvarUSUARIO, gvarPASSWORD) <> "" Then End
''''''''''''    gvarUSUARIO = "BTLPROD"
''''''''''''
''''''''''''    ''gvarUSUARIO = "VENTAS"
''''''''''''    gvarPASSWORD = gclsOracle.FN_Valor("GENERAL.FN_PASE", UCase(gvarUSUARIO))
''''''''''''
''''''''''''    If gclsOracle.Conexion(gvarTNSNAME, gvarUSUARIO, gvarPASSWORD) <> "" Then End
''''''''''''
''''''''''''    gclsOracle.Execute "BEGIN DBMS_APPLICATION_INFO.SET_MODULE('" & App.FileDescription & "','" & App.Major & "." & App.Minor & "." & App.Revision & "'); END ;"
''''''''''''
''''''''''''
''''''''''''    Set gosesVentas = gclsOracle.OSession
''''''''''''    Set godbVentas = gclsOracle.ODataBase
''''''''''''
''''''''''''    '******************************************************************************************************************'
''''''''''''    '*** Validando que cuando sea Delivery levante mas de una session ***'
''''''''''''    '*** Cambio Hecho el 01/10/2007 Por Crueda                        ***'
''''''''''''    '******************************************************************************************************************'
''''''''''''    Dim odynLocal As oraDynaset
''''''''''''    Dim strLocal As String
''''''''''''    Dim strDlv As String
''''''''''''    Dim dblCnt As Integer
''''''''''''    Dim i As Integer
''''''''''''
''''''''''''    Set odynLocal = gclsOracle.FN_Cursor("BTLPROD.PKG_MAQUINA.FN_LISTA", 0, Replace(Trim(Obtener(12)), Chr(0), ""))
''''''''''''    strLocal = "" & odynLocal("COD_LOCAL").Value
''''''''''''
''''''''''''    strDlv = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_LOCAL_DELIVERY")
''''''''''''
''''''''''''    If Trim(strDlv) <> Trim(strLocal) Then
''''''''''''        If App.PrevInstance Then
''''''''''''            Screen.MousePointer = vbDefault
''''''''''''            MsgBox "La Aplicación Ya Esta Siendo Ejecutada", vbInformation, App.ProductName
''''''''''''            End
''''''''''''        End If
''''''''''''    End If
''''''''''''    '******************************************************************************************************************'
''''''''''''    '******************************************************************************************************************'
''''''''''''    gintDec = gclsOracle.Const_Val("CMR.PK_COM_PRECIOS.V_DECIMALES")
''''''''''''    gintDecTot = gclsOracle.Const_Val("CMR.PK_COM_PRECIOS.V_DECIMALES_TOT")
''''''''''''
''''''''''''    On Error GoTo 0
''''''''''''    GoTo Final
''''''''''''ErrorConexion:
''''''''''''    Screen.MousePointer = vbDefault
''''''''''''    MsgBox "Existe un Problema con la Conexión" & Chr(13) & "Error :" & Err.Description, vbCritical, App.ProductName
''''''''''''    Exit Sub
''''''''''''Final:
''''''''''''
''''''''''''    Screen.MousePointer = vbDefault
''''''''''''
''''''''''''    'MsgBox "Se hizo la conexión con exito '" & gvarUSUARIO & "'"
''''''''''''    '''''''''''''''''''''''''''''''''''''''''''frm_VTA_Logueo.Caption = frm_VTA_Logueo.Caption & " " & gstrAplicacion & " * Ver: " & gstrVersion & " - " & gvarTNSNAME
''''''''''''
''''''''''''    'If Trim(strDlv) = Trim(StrLocal) Then
''''''''''''    '    Dim frm As New frm_VTA_Logueo
''''''''''''    '''''''''''''''''''''''''''''''''''''''''''''''''''frm_VTA_Logueo.Show vbModal
''''''''''''    'frm.Show vbModal
''''''''''''    'End If
''''''''''''    ''''frm_DLV_HistorialCliente.Mostrar objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, "0000001761"
''''''''''''    ''''frm_ADM_GuiasTransito.Show vbModal
''''''''''''
'''''''''''''''''''''''''    If objUsuario.flgDeliveryProv = "1" Then frm_VTA_TipoMaquina.Show vbModal
'''''''''''''''''''''''''
'''''''''''''''''''''''''    If objUsuario.Conectado Then
'''''''''''''''''''''''''        tipoPantalla
'''''''''''''''''''''''''    End If
''''''''''''
''''''''''''
''''''''''''End Sub

Public Sub psub_Cerrar()
    gclsOracle.Cerrar
    End
End Sub

'******** Escritura de solo numero********
Function SoloNumeros(Key%, Optional Simbol$) As Integer
    'If IsMissing(Simbol) Then key = 0: Exit Function
    Select Case Key
        Case 13, 8, 27, 48 To 57
        Case AscW(IIf(Simbol = "", 0, Simbol))
        Case Else: Key = 0
    End Select
    SoloNumeros = Key
End Function

Sub setteaFormulario(Formulario As Form)

'Autor : Arturo Escate Espichan
'Fecha : 02/05/2008
If InStr(1, Formulario.name, "OFF", vbTextCompare) > 0 Then
    With frm_OFF_Principal
        Formulario.Width = 7190
        Formulario.Height = 7300
        Formulario.left = .ScaleLeft '+ .Picture2.left + 200
        Formulario.top = .ScaleTop + .Picture2.top + 400
    End With
    Exit Sub
End If

    Formulario.top = 0
    Formulario.left = 0
    Formulario.Appearance = 0
    Formulario.Width = 12800 '12600 '7240
    Formulario.Height = 8500 '7275
    Formulario.BackColor = RGB(212, 208, 200)
'COMENTADO POR PHERRERA 23/08/07
'ESTO ES ESTANDAR NO ES NECESARIO HACER ESTE CÓDIGO PARA LOS COLORES
'DEMORA DEMASIADO.
'    Dim j As Integer
'    While j < Formulario.Controls.Count
'        Select Case TypeName(Formulario.Controls(j))
'        Case "Label"
'            Formulario.Controls(j).BackStyle = 0
'        Case "Frame"
'            Formulario.Controls(j).BackColor = RGB(212, 208, 200)
'        Case "OptionButton"
'            Formulario.Controls(j).BackColor = RGB(212, 208, 200)
'        End Select
'        j = j + 1
'    Wend
End Sub

Sub tipoPantalla()

    If Not m_tipo_maquina = "" Then
        objUsuario.TipoMaquina = m_tipo_maquina
    End If

    Select Case objUsuario.TipoMaquina
        Case objUsuario.TipoMaquinaAdmin 'cuando es la maquina del administrador
               mdiPrincipal.Show
               mdiPrincipal.pctDelivery.Visible = False
               'mdiPrincipal.ctlCliente1.Modo = False
               mdiPrincipal.ctlCliente1.Modo False
               mdiPrincipal.Caption = gstrAplicacion & " * Ver: " & gstrVersion & " [" & gvarUSUARIO & "@" & gvarTNSNAME & "] " & "Local " & objUsuario.CodigoLocal & " de la Empresa " & objUsuario.Empresa
               'mdiPrincipal.Caption = gstrAplicacion & " * Ver: " & gstrVersion & " [" & gvarUSUARIO & "@" & gvarTNSNAME & "] " & "Local " & objUsuario.CodigoLocal & " >" & objUsuario.Empresa & "<"
               mdiPrincipal.HabilitaPermisos
               frmPedido.FormDragger1.Caption = objUsuario.Nombre
               mdiPrincipal.Show
        Case objUsuario.TipoMaquinaCajero 'cuando es la maquina de un cajero
                mdiPrincipal.Show
               mdiPrincipal.pctDelivery.Visible = False
               'mdiPrincipal.ctlCliente1.Modo = False
               mdiPrincipal.ctlCliente1.Modo False
               mdiPrincipal.Caption = gstrAplicacion & " * Ver: " & gstrVersion & " [" & gvarUSUARIO & "@" & gvarTNSNAME & "] " & "Local " & objUsuario.CodigoLocal & " de la Empresa " & objUsuario.Empresa
               'mdiPrincipal.Caption = gstrAplicacion & " * Ver: " & gstrVersion & " [" & gvarUSUARIO & "@" & gvarTNSNAME & "] " & "Local " & objUsuario.CodigoLocal & " >" & objUsuario.Empresa & "<"
               mdiPrincipal.HabilitaPermisos
               frmPedido.FormDragger1.Caption = objUsuario.Nombre
               mdiPrincipal.Show
        Case objUsuario.TipoMaquinaCabina 'cuando es una cabina
               mdiPrincipal.Show
               mdiPrincipal.Caption = gstrAplicacion & " * Ver: " & gstrVersion & " [" & gvarUSUARIO & "@" & gvarTNSNAME & "] " & "Local " & objUsuario.CodigoLocal & " >" & objUsuario.Empresa & "<"
               mdiPrincipal.pctDelivery.Visible = True
               'mdiPrincipal.ctlCliente1.Modo = True
               mdiPrincipal.ctlCliente1.Modo True
               'mdiPrincipal.cmdGrabaVenta.Visible = False
               mdiPrincipal.Label1(4).Visible = True
              ' mdiPrincipal.txtDLVTelefono.SetFocus
               
        Case objUsuario.TipoMaquinaRuteo ' cuando es el modulo de ruteo
            frm_DLV_Seguimiento.Show
        Case objUsuario.TipoMaquinaVerif
            frm_DLV_Verificacion.Show
        Case Else
            MsgBox "No se ha definido el tipo de maquina |" & objUsuario.TipoMaquina & "|", vbCritical, App.ProductName
            End
    End Select
    
    'frm_DLV_Seguimiento
    
    If Not callcenter = "" Then
        objUsuario.CodLocalCallCenter = callcenter
    End If
    Dim colorFondo As String
    Dim colorR As Integer
    Dim colorG As Integer
    Dim colorB As Integer
    colorFondo = "" & gclsOracle.FN_Valor("BTLPROD.PKG_CALLCENTER_MARCA.FN_GET_FONDO_RGB", objUsuario.CodLocalCallCenter)
    If Not colorFondo = "" Then
        colorR = separaCadena(colorFondo, ",", 0)
        colorG = separaCadena(colorFondo, ",", 1)
        colorB = separaCadena(colorFondo, ",", 2)
        mdiPrincipal.Frame1.BackColor = RGB(colorR, colorG, colorB)
        mdiPrincipal.pctDelivery.BackColor = RGB(colorR, colorG, colorB)
        mdiPrincipal.ctlCliente1.Modo True, colorR, colorG, colorB
    End If
    
    If objUsuario.TipoMaquina = "003" Then
        LogoMarca
    End If
End Sub

Public Function LogoMarca(Optional urlFolderImagen As String)
    Dim imagenFileUrl As String
    'Dim urlFolderImagen As String
    'Dim error As Boolean
    If urlFolderImagen = "" Then
        urlFolderImagen = CStr(gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_FILE_IMAGEN_LOGO"))
    End If
                
    If Dir(urlFolderImagen, vbDirectory) = "" Then Exit Function
    imagenFileUrl = urlFolderImagen & objUsuario.CodLocalCallCenter & ".jpg"
    If Dir(imagenFileUrl, vbDirectory) = "" Then Exit Function
    
    'If Dir(imagenFileUrl, vbDirectory) <> "" Then
    Dim logo As StdPicture
    Set logo = LoadPicture(imagenFileUrl) 'LoadPicture("C:\Desarrollo\imagen\430812\LogoMifarma.jpg")
    mdiPrincipal.pctDelivery.AutoRedraw = True
    mdiPrincipal.pctDelivery.PaintPicture logo, 17535, 225, 2632, 632
    'Else
End Function

Public Function CME(strCampo As String, intEspacios As Integer, Optional intPosicion As String = "I") As String
'''''CadenaMasEspacio'''''
''''''''''''''''''''''''''
'''''''intPosicion''''''''
''''''''''''''''''''''''''
'''''"I" = IZQUIERDA''''''
'''''"D" =  DERECHA  '''''
''''''''''''''''''''''''''
Dim strCampo2 As String
   strCampo2 = Trim(Mid(strCampo, 1, intEspacios))
   If intPosicion = "I" Then
      CME = strCampo2 & Space$(intEspacios - Len(strCampo2))
   Else
      CME = Space$(intEspacios - Len(strCampo2)) & strCampo2
   End If
End Function

'Private Function CME(strCampo As String, intEspacios As Integer, Optional intPosicion As String = "I") As String
'
'Dim strCampo2 As String
'
'   strCampo2 = Trim(Mid(strCampo, 1, intEspacios))
'   If strCampo2 = "0" Then
'      CME = Space$(intEspacios)
'   ElseIf intPosicion = "I" Then
'      CME = strCampo2 & Space$(intEspacios - Len(strCampo2))
'   Else
'      CME = Space$(intEspacios - Len(strCampo2)) & strCampo2
'   End If
'
'End Function

Public Function fstr_Centrar(ByVal vstrCad$, ByVal vintTam%) As String
    If Len(vstrCad) < vintTam Then
        fstr_Centrar = Space((vintTam - Len(vstrCad)) \ 2) & vstrCad & Space(vintTam - Len(vstrCad) - (vintTam - Len(vstrCad)) \ 2)
    Else
        fstr_Centrar = left(vstrCad, vintTam)
    End If
End Function

'**********************************************************
'Resaltar objetos-Martinnets
'**********************************************************
Sub resaltaObjeto(objResaltar As Object, mform As Form, Optional miliSegundos As Long = 300)
    gb_resaltar = 0: gb_Salir = False
    Set objresalta = objResaltar
    SetTimer mform.hwnd, 1, miliSegundos, AddressOf TimerProc
    If TypeOf objresalta Is TextBox Then objresalta.SetFocus
    If TypeOf objresalta Is ctlTextBox Then objresalta.SetFocus
End Sub

Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    On Error Resume Next
    If gb_resaltar = 5 Then gb_Salir = True
    If gb_Salir Then
        Exit Sub
        KillTimer hwnd, idEvent
    Else
        gb_resaltar = gb_resaltar + 1
        If gb_resaltar Mod 2 = 0 Then objresalta.BackColor = &HC0C0FF
        If gb_resaltar Mod 2 = 1 Then objresalta.BackColor = &HFFFFFF
    End If
End Sub
'**********************************************************


'*************************************************************
'PARA ENCRIPTAR CADENA DE TEXTO
'*************************************************************
Public Function strEncrypt(ByVal strMsg As String) As String
    On Local Error Resume Next
    Dim ByteArray() As Byte, byteKey() As Byte, CryptText() As Byte
        Set oTest = New CRijndael
        ByteArray() = strConv(strMsg, vbFromUnicode)
        byteKey() = strConv(pEncryptionKey, vbFromUnicode)
        CryptText = oTest.EncryptData(ByteArray(), byteKey())
        Set oTest = Nothing
        strEncrypt = strConv(CryptText(), vbUnicode)
End Function

Public Function strDecrypt(ByVal strMsg As String) As String
    On Local Error Resume Next
    Dim ByteArray() As Byte, byteKey() As Byte, CryptText() As Byte
    Set oTest = New CRijndael
    ByteArray() = strConv(strMsg, vbFromUnicode)
    byteKey() = strConv(pEncryptionKey, vbFromUnicode)
    CryptText() = oTest.DecryptData(ByteArray(), byteKey())
    Set oTest = Nothing
    strDecrypt = strConv(CryptText(), vbUnicode)
End Function

Public Function Hex2Str(ByVal strdata As String)
    On Local Error Resume Next
    Dim i As Long, CryptString As String, tmpChar As String
    For i = 1 To Len(strdata) Step 2
    CryptString = CryptString & _
     Chr(val("&H" + Mid(strdata, i, 2)))
    Next i
    Hex2Str = CryptString
End Function

Public Function Str2Hex(ByVal strdata As String)
    On Local Error Resume Next
    Dim i As Long, CryptString As String, tmpAppend As String
    For i = 1 To Len(strdata)
    tmpAppend = Hex(Asc(Mid(strdata, i, 1)))
    If Len(tmpAppend) = 1 Then tmpAppend = Trim(str(0)) & tmpAppend
    CryptString = CryptString & tmpAppend: DoEvents
    Next
    Str2Hex = CryptString
End Function

Public Function fn_Encriptar(ByVal strTexto As String) As String
Dim sTemp As String, sPassword As String
    sTemp = strTexto
    sPassword = "thekeymaker" 'InputBox("Key", "Encrypt", "Password Passphrase")
    If StrPtr(sPassword) = 0 Then Exit Function
        sTemp = strEncrypt(sTemp)
    fn_Encriptar = Str2Hex(sTemp)

End Function
'*************************************************************

'funcion que indica si un formulario esta cargado
Public Function IsLoaded(strFormName As String) As Boolean
    Dim i As Integer
    For i = 0 To Forms.Count - 1
       If (Forms(i).name = strFormName) Then
            IsLoaded = True
            Exit For
        End If
    Next
End Function

Public Function fncGuion(strTexto$, Optional intFlag% = 0, Optional strCar$ = "|") As String
Dim intA
    strTexto = CStr(strTexto)
    intA = InStr(strTexto, strCar)
    If intA = 0 Then fncGuion = Trim(strTexto): Exit Function
    If intFlag = 0 Then fncGuion = Trim(left(strTexto, intA - 1))
    If intFlag = 1 Then fncGuion = Trim(right(strTexto, Len(strTexto) - intA))
End Function

Public Function fn_EsIPCorrecta(ByVal strIP As String) As Boolean
Dim arrPartes() As String, i As Long, b As Byte
On Error GoTo ProcesarErrores
arrPartes = Split(strIP, ".")
For i = LBound(arrPartes) To UBound(arrPartes)
    b = 0
    b = b + CByte(arrPartes(i))
Next i
fn_EsIPCorrecta = b > 0
ProcesarErrores:
    
End Function

'---------------------------------------------------------------------------------
Function pfstr_Segmento(ByVal vstrCadena$, Optional ByVal vblnDerecha As Boolean, Optional ByVal vstrSimbolo$) As String
Dim intA%
    intA = InStr(vstrCadena, IIf(vstrSimbolo = "", "-", vstrSimbolo))
    pfstr_Segmento = Trim(IIf(Not vblnDerecha, left(vstrCadena, IIf(intA = 0, Len(vstrCadena) + 1, intA) - 1), right(vstrCadena, Len(vstrCadena) - intA)))
End Function

Function pfstr_Periodo(Optional ByVal vintMes%, Optional ByVal vblnActual As Boolean = False)

    If Not vblnActual Then
        'gvarParametros = Array("A_MES"):
        'gvarValores = Array(vintMes)
        'strSql = " SELECT TO_CHAR(ADD_MONTHS(TRUNC(LAST_DAY(SYSDATE)) + 1, -1*:A_MES), 'YYYYMM') FROM DUAL "
        'Call psub_Ejecuta_Query(gclsOracle.ODataBase, strSql, 0&, odynP, gvarParametros, gvarValores)
        pfstr_Periodo = Format(DateSerial(Format(gclsOracle.Fecha_Servidor, "YYYY"), _
                   Format(gclsOracle.Fecha_Servidor, "MM") - vintMes, _
                   Format(gclsOracle.Fecha_Servidor, "DD")), "YYYYMM")
    Else
        'strSql = " SELECT TO_CHAR(SYSDATE, 'YYYYMM') FROM DUAL "
        'Call psub_Ejecuta_Query(gclsOracle.ODataBase, strSql, 0&, odynP)
        pfstr_Periodo = Format(gclsOracle.Fecha_Servidor, "YYYYMM")
    End If
    
End Function

Sub psub_Cad_Arreglo(ByRef vvarArreglo As Variant, _
                    ByVal vstrCadena$, _
                    Optional ByVal vstrSeparador = "|")
Dim strCadena$

    ReDim vvarArreglo(0 To 0)
    If right(vstrCadena, 1) <> vstrSeparador Then
        strCadena = vstrCadena & vstrSeparador
    Else
        strCadena = vstrCadena
    End If
    
    While strCadena <> ""
        vvarArreglo(UBound(vvarArreglo)) = pfstr_Segmento(strCadena, False, vstrSeparador)
        strCadena = pfstr_Segmento(strCadena, True, vstrSeparador)
        ReDim Preserve vvarArreglo(UBound(vvarArreglo) + 1)
    Wend
    ReDim Preserve vvarArreglo(UBound(vvarArreglo) - 1)
End Sub

Public Function fncPalote(strTexto$, Optional intFlag% = 0, Optional strCar$ = "|") As String
Dim intA
    strTexto = CStr(strTexto)
    intA = InStr(strTexto, strCar)
    If intA = 0 Then fncPalote = Trim(strTexto): Exit Function
    If intFlag = 0 Then fncPalote = Trim(left(strTexto, intA - 1))
    If intFlag = 1 Then fncPalote = Trim(right(strTexto, Len(strTexto) - intA))
End Function




Function EsEscaneado(KeyCode As Integer, Shift As Integer) As Boolean
Dim DIF As Double
EsEscaneado = True
Exit Function
DIF = 0
    Dim ltime As Long
      ltime = timeGetTime
      
If KeyCode = vbKeyReturn Then
    If ACABA = True Then
        ACABA = False
        KeyTime = 0
        EsEscaneado = True
        
    Else
        ACABA = False
        KeyTime = 0
        EsEscaneado = False
        Exit Function
    End If
    
End If
   If ACABA = True Then Exit Function
    
    If KeyTime = 0 Then
        KeyTime = Format(time, "HHmmss") & ltime
        
        DIF = 0
    Else
        DIF = (Format(time, "HHmmss") & ltime) - KeyTime
        KeyTime = Format(time, "HHmmss") & ltime
    End If
    
    Debug.Print KeyCode & "DEMORO:" & DIF & "-->" & IIf(DIF < 40, "Escanea", "digita")

'EsEscaneado = True
If DIF < 40 And DIF <> 0 Then
    ACABA = True
Else
    ACABA = False
End If

End Function

''' David Jara - 2015/06/03
Public Function EncriptarNumeroTarjeta(ByVal vData As String) As String
    'EncriptarNumeroTarjeta = Mid$(vData, 1, 4) & "******" & Mid$(vData, Len(vData) - 3)
    EncriptarNumeroTarjeta = "**********" & Mid$(vData, Len(vData) - 3)
End Function

Public Function SetearImpresoraCupon() As Boolean
    Dim strNombreCuponera As String
    Dim bolValor As Boolean
    Dim Impresora As Printer

    On Error GoTo Control
    strNombreCuponera = objVenta.NombreCuponera(objUsuario.NombrePC)
    If strNombreCuponera = "" Then
        bolValor = False
    Else
        bolValor = False
        For Each Impresora In Printers
            If UCase(strNombreCuponera) = UCase(Impresora.Devicename) Then
                Set Printer = Impresora
                bolValor = True
                Exit For
            End If
        Next
    End If
    SetearImpresoraCupon = bolValor
    Exit Function
Control:
    SetearImpresoraCupon = False
End Function

Public Sub centra_printer(txt)
    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(txt)) \ 2
    Printer.Print txt
End Sub

Public Sub justifica_printer(x0, xf, y0, txt)
    ' x0, xf = posicion de los margenes izquierdo y derecho
    ' y0 = posicion vertical donde se desea empezar a escribir
    ' txt = texto a escribir
    Dim x, Y, k, Ancho
    Dim s As String, ss As String
    Dim x_spc

    s = txt
    x = x0
    Y = y0
    Ancho = (xf - x0)

    While s <> ""

        ss = ""
        While (s <> "") And (Printer.TextWidth(ss) <= Ancho)
            ss = ss & left$(s, 1)
            s = right$(s, Len(s) - 1)
        Wend
        If (Printer.TextWidth(ss) > Ancho) Then
            s = right$(ss, 1) & s
            ss = left$(ss, Len(ss) - 1)
        End If
        ' aqui tenemos en ss lo maximo que cabe en una linea
        If right$(ss, 1) = " " Then
            ss = left$(ss, Len(ss) - 1)
        Else
            If (InStr(ss, " ") > 0) And (left$(s & " ", 1) <> " ") Then
                While right$(ss, 1) <> " "
                    s = right$(ss, 1) & s
                    ss = left$(ss, Len(ss) - 1)
                Wend
                ss = left$(ss, Len(ss) - 1)
            End If
        End If
        x_spc = 0
        x = x0
        If (Len(ss) > 1) And (s & "" <> "") Then
            x_spc = (Ancho - Printer.TextWidth(ss)) / (Len(ss) - 1)
        End If
        Printer.CurrentX = x
        Printer.CurrentY = Y

        If x_spc = 0 Then
            Printer.Print ss;
        Else
            For k = 1 To Len(ss)
                Printer.CurrentX = x
                Printer.Print Mid$(ss, k, 1);
                x = x + Printer.TextWidth("*" & Mid$(ss, k, 1) & "*") - Printer.TextWidth("**")
                x = x + x_spc
            Next
        End If

        Y = Y + Printer.TextHeight(ss)
        While left$(s, 1) = " "
            s = right$(s, Len(s) - 1)
        Wend
    Wend

End Sub
''' ------------------------------------------
Public Function fMapearUnidadDeRed() As Boolean
    Dim unidadDefault
    Dim unidadMifarma As String
    Dim unidadInkafarma As String
    Dim unidad
    Dim Count
    
    unidadMifarma = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "IMGPRODMIF", "10")
    '"K: \\192.168.0.221\Delivery\Imagen /user:MIFARMA\ventas ventas"
    unidadInkafarma = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "IMGPRODINK", "10")
    '"K: \\192.168.0.221\Delivery\Imagen /user:Eckerdholding\delivery20 V3nt4s2019$"
    unidadDefault = Split(unidadMifarma, "/USER")
    unidad = Split(unidadMifarma, ":")
    Count = UBound(unidad)
    If Count < 0 Then Exit Function
    Open "C:\AppBTL\unidadImagen.bat" For Output As #1
        Print #1, "@echo off"
        Print #1, "NET USE "; unidad(0); ": /D /Y"
        Print #1, "NET USE "; IIf(UBound(unidadDefault) < 0, "", unidadDefault(0)); " /PERSISTENT:YES"
        Print #1, "NET USE "; unidadMifarma; " /PERSISTENT:YES"
        Print #1, "NET USE "; unidadInkafarma; " /PERSISTENT:YES"
        Print #1, "TASKKILL /IM cmd.exe /F"
    Close #1
    Shell "C:\AppBTL\unidadImagen.bat", vbHide
End Function

