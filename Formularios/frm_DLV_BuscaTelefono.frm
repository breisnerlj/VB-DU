VERSION 5.00
Begin VB.Form frm_DLV_BuscaTelefono 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda de Clientes"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "frm_DLV_BuscaTelefono.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSoloPrecios 
      Caption         =   "Precios"
      Height          =   435
      Left            =   4605
      TabIndex        =   19
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin vbp_Ventas.ctlTextBox txtNombre 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      Tipo            =   2
      TipoSQL         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbp_Ventas.ctlGrilla grdDireccion 
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3201
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   3285
      TabIndex        =   4
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   5940
      TabIndex        =   5
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Editar"
      Height          =   435
      Left            =   1950
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Nuevo"
      Height          =   435
      Left            =   630
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin vbp_Ventas.ctlGrilla ctlGrilla1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3836
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.CommandButton Command1x 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   5025
      TabIndex        =   20
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   18
      Top             =   195
      Width           =   255
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   0
      Left            =   0
      TabIndex        =   17
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lista de Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   300
      TabIndex        =   16
      Top             =   735
      Width           =   1470
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   2
      Left            =   0
      TabIndex        =   15
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lista de Direcciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   300
      TabIndex        =   14
      Top             =   3735
      Width           =   1830
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   13
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Index           =   1
      Left            =   1050
      TabIndex        =   12
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   2370
      TabIndex        =   11
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   3705
      TabIndex        =   10
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   6270
      TabIndex        =   9
      Top             =   6480
      Width           =   540
   End
End
Attribute VB_Name = "frm_DLV_BuscaTelefono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Telefono As String
Dim flgDescarga As Boolean
Dim objCliente As New clsCliente
Dim objLocal As New clsLocal 'ECASTILLO 27.10.2020

Private Sub cmdAceptar_Click()
Dim strCodigo As String
Dim strFlagJuridico As String
Dim strRazonSocial As String
Dim intVerificado As Integer
Dim strSufijo As String
Dim strLocalAsignado As String
Dim strLocalDespacho As String
Dim strNomComercial As String
Dim strNomCliente As String
Dim strApeCliente As String
Dim strApeMaterno As String
Dim strDireccionSocial As String
Dim strDireccionComercial As String
Dim strCodDocumentoID As String
Dim strNumDocumentoID As String

If ctlGrilla1.ApproxCount = 0 Then Exit Sub
On Error GoTo CtrlErr
    'I.ECASTILLO 17.12.2020 | dc Cappa | 2da etapa reserva
    Dim flg_ruteoA_cnv
    Dim sCia As String
    Dim rsCia As oraDynaset
    Dim flgFunLocal As String
    Dim flg_2e_reserva
    Dim dataUbigeo As oraDynaset
    
    flg_ruteoA_cnv = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRACNV") '1 => ACTIVO, 0 => INACTIVO
    
    flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
    
    If mdiPrincipal.ctlCliente1.seleccionManualLocal = True Then
        strLocalDespacho = "" & objVenta.bk_codLocal
    Else
        strLocalDespacho = "" & IIf(Len(Trim(objVenta.bk_codLocal)) = 0, ctlGrilla1.DataSource("COD_LOCAL_DESPACHO").Value, objVenta.bk_codLocal)
    End If
    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, strLocalDespacho)
    If (rsCia.RecordCount > 0) Then
      sCia = CStr(rsCia(1))
    End If
    Set rsCia = Nothing
    objVenta.flg_2e_reserva_local = objLocal.GetEstConfig(sCia, strLocalDespacho, "RESERVA_STOCK_2DA")
    objVenta.isLocalDcCappa = "" & objLocal.GetIndCDCAP(strLocalDespacho)
    'F.ECASTILLO 17.12.2020

    strCodigo = "" & ctlGrilla1.DataSource("COD_CLIENTE").Value
    strLocalAsignado = "" & IIf(Len(Trim(objVenta.bk_codLocal)) = 0, ctlGrilla1.DataSource("COD_LOCAL_PRECIO").Value, objVenta.bk_codLocal)
    'I.ECASTILLO 17.12.2020
'    If flg_2e_reserva = "0" Then
'        strLocalDespacho = "" & ctlGrilla1.DataSource("COD_LOCAL_DESPACHO").Value
'    Else
''        objVenta.isLocalDcCappa = "" & objLocal.GetIndCDCAP(mdiPrincipal.ctlCliente1.LocalDespacho)
'        If objVenta.isLocalDcCappa = "0" And objVenta.flg_2e_reserva_local = "1" Then
'            If mdiPrincipal.ctlCliente1.seleccionManualLocal = True Then
'                strLocalDespacho = "" & mdiPrincipal.ctlCliente1.LocalDespacho
'            Else
'                strLocalDespacho = "" & mdiPrincipal.ctlCliente1.LocalDespacho '"" & ctlGrilla1.DataSource("COD_LOCAL_DESPACHO").Value
'            End If
'        Else
'            strLocalDespacho = "" & ctlGrilla1.DataSource("COD_LOCAL_DESPACHO").Value
'        End If
'    End If
    'F.ECASTILLO 17.12.2020
    strFlagJuridico = "" & ctlGrilla1.DataSource("FLG_TIPO_JURIDICA").Value
    strRazonSocial = "" & ctlGrilla1.DataSource("DES_RAZON_SOCIAL").Value
    intVerificado = "" & ctlGrilla1.DataSource("FLG_CLIENTE_VERIFICADO").Value
    strSufijo = "" & ctlGrilla1.DataSource("SUFIJO").Value
    strNomComercial = "" & ctlGrilla1.DataSource("DES_NOM_COMERCIAL").Value
    strNomCliente = "" & ctlGrilla1.DataSource("DES_NOM_CLIENTE").Value
    strApeCliente = "" & ctlGrilla1.DataSource("DES_APE_CLIENTE").Value
    strApeMaterno = "" & ctlGrilla1.DataSource("DES_APE2_CLIENTE").Value
    strDireccionSocial = "" & ctlGrilla1.DataSource("DES_DIRECCION_SOCIAL").Value
    strDireccionComercial = "" & ctlGrilla1.DataSource("DES_DIRECCION_COMERCIAL").Value
    strCodDocumentoID = "" & ctlGrilla1.DataSource("COD_DOCUMENTO_IDENTIDAD").Value
    strNumDocumentoID = "" & ctlGrilla1.DataSource("NUM_DOCUMENTO_ID").Value
    objVenta.Latitud = "" & grdDireccion.DataSource("LATITUD").Value
    objVenta.Longitud = "" & grdDireccion.DataSource("LONGITUD").Value
    objVenta.bk_DescSufijoDir = "" & grdDireccion.Columns("DES_SUFIJO_DIRECCION").Value
    objVenta.bk_AbrSufijoDir = "" & grdDireccion.Columns("DES_ABREVIATURA_DIRECCION").Value
    If gstrFlagSufDir = "1" Then
        objVenta.bk_SufijoDir = objVenta.bk_AbrSufijoDir
    ElseIf gstrFlagSufDir = "2" Then
        objVenta.bk_SufijoDir = objVenta.bk_DescSufijoDir
    End If


    
    With mdiPrincipal.ctlCliente1
        ''.Cargar
        ''.ConsultaCliente "" & ctlGrilla1.DataSource("COD_CLIENTE").Value
        
        
        
        .CargaDatosCliente strCodigo, intVerificado, strLocalAsignado, strFlagJuridico, _
            strRazonSocial, strNomComercial, strNomCliente, strApeCliente, strApeMaterno, _
            strLocalDespacho, strSufijo, strDireccionSocial, strDireccionComercial, _
            strNumDocumentoID, strCodDocumentoID
        
        
        If (.LocalAsignado = "" Or strLocalAsignado = "") Then
            MsgBox "El cliente no tiene local asignado", vbCritical, App.ProductName
            Exit Sub
        End If
        
        
        
        If .FlagJuridico = 1 Then
            objVenta.NombreClienteDLV = IIf(IsNull(.RazonSocial), "", .RazonSocial)
            objVenta.DesAuxCliDirecc = IIf(IsNull(.DireccionComercial), "", .DireccionSocial)
            objVenta.NumeroDocumentoID = .NumeroDocumentoID
        Else
            objVenta.NombreClienteDLV = ctlGrilla1.DataSource("XNOMBRE").Value
            objVenta.DesAuxCliDirecc = IIf(IsNull(.DireccionSocial), "", .DireccionSocial)
            objVenta.NumeroDocumentoID = ""
        End If
        

        
        
        '*************************************************************************************'
        'I.ECASTILLO 22.06.2021
    If objVenta.respetaLocal <> True Then
        GoTo cnvNoRuteaAuto
    End If
    If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
        GoTo cnvNoRuteaAuto
    End If
    Set rsCia = Nothing
    sCia = ""
    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, strLocalDespacho)
    If (rsCia.RecordCount > 0) Then
      sCia = CStr(rsCia(1))
    End If
    Set rsCia = Nothing
    flgFunLocal = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVET3") '1 => ACTIVO, 0 => INACTIVO
    If flg_2e_reserva = "1" And flgFunLocal = "1" Then
        GoTo CallServiceReserva2
    End If
    
    If flg_2e_reserva = "0" Or objVenta.flg_2e_reserva_local = "0" Then
        GoTo cnvNoRuteaAuto
    Else
CallServiceReserva2:
        If (objVenta.bk_codCliente) <> Trim(ctlGrilla1.DataSource("COD_CLIENTE").Value) Then
            mdiPrincipal.ctlCliente1.seleccionManualLocal = False
            GoTo cnvNoRuteaAuto
        ElseIf Len(Trim(objVenta.dc_street)) <> 0 And Trim(objVenta.bk_codCliente) = Trim(ctlGrilla1.DataSource("COD_CLIENTE").Value) And _
            Trim(grdDireccion.Columns("DES_DIRECCION").Value) <> Trim(objVenta.dc_street) Then
            If MsgBox("Ud. geolocalizó la dirección " & objVenta.dc_street & ", desea modificar la dirección original?", _
                        vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
                objVenta.DireccionClienteDLV = "" & Trim(objVenta.dc_street)
                objVenta.dc_street = IIf(Len(Trim(objVenta.dc_street)) = 0, "", objVenta.DireccionClienteDLV)
                objVenta.DireccionClienteDLV = objVenta.DireccionClienteDLV
                
                objVenta.DireccionCliente = objVenta.DireccionClienteDLV
                objVenta.Out_Direccion = objVenta.DireccionClienteDLV
            Else
                GoTo cnvNoRuteaAuto
            End If
        Else
cnvNoRuteaAuto:
            objVenta.DireccionClienteDLV = "" & grdDireccion.Columns("DES_DIRECCION").Value
        End If
    End If
    'F.ECASTILLO 22.06.2021
        'objVenta.DireccionClienteDLV = "" & grdDireccion.Columns("DES_DIRECCION").Value
        objVenta.UbigeoEntrega = "" & grdDireccion.Columns("COD_UBIGEO")
        objVenta.bk_Ubigeo = objVenta.UbigeoEntrega
        objVenta.DesReferenciaCli = "" & grdDireccion.Columns("DES_REFERENCIA_DIRECCION")
        objVenta.DesAuxCliTlf = IIf(IsNull(Telefono), "", Telefono)
        objVenta.DesDistritoDLV = "" & grdDireccion.Columns("DES_DISTRITO").Value
        objVenta.DesUrbanizacionDLV = "" & grdDireccion.Columns("DES_URBANIZACION").Value
        objVenta.ObservacionClienteDLV = "" & ctlGrilla1.DataSource("DES_OBSERVACION").Value
        objVenta.CodDireccionCli = "" & grdDireccion.DataSource("COD_DIRECCION_CLI").Value
        
        
        objVenta.Out_NumeroId = objVenta.NumeroDocumentoID
        objVenta.Out_NombreCliente = objVenta.NombreClienteDLV
        objVenta.Out_Tipo = .FlagJuridico
        objVenta.Out_CodDocumentoId = .CodigoDocumentoID
        objVenta.Out_Direccion = objVenta.DireccionClienteDLV
        If gstrIndRAv4 = "1" Then
            objVenta.dc_urbanizacion = objVenta.DesUrbanizacionDLV
            objVenta.dc_referencia = objVenta.DesReferenciaCli
            Set dataUbigeo = objCliente.getUbigeoDesc(objVenta.bk_Ubigeo)
            objVenta.dc_departamentBK = ""
            If Not dataUbigeo.EOF Then
                objVenta.dc_departamentBK = "" & dataUbigeo("DES_DEPARTAMENTO").Value
            End If
        End If
    End With
    'operado 15/10/09 por pherrera para el problema de DELIVERY, que era que
    'habia ocasiones que el codigo de cliente se grababa erroneamente y en el punto de venta
    'se mostraba un error de que no tenia linea de credito. (de hacer pruebas 1 hora seguida,
    'se encontro que se chancaban los datos del beneficiario cuando se ingresaba nuevamente al telefono)
    If objVenta.CodigoCliente = "" Or objVenta.CodModalidadVenta <> "002" Then
        objVenta.CodigoCliente = ctlGrilla1.DataSource("COD_CLIENTE").Value
    End If
    objVenta.CodigoClienteDLV = ctlGrilla1.DataSource("COD_CLIENTE").Value
    objVenta.EntregaTercero = "0"
    
    
    ''objVenta.Out_CodigoCliente = ctlGrilla1.DataSource("COD_CLIENTE").Value
    ''objVenta.UbigeoEntrega = mdiPrincipal.ctlCliente1.
    Dim rsTelefono As oraDynaset
    Set rsTelefono = objCliente.TelefonoCliente(ctlGrilla1.DataSource("COD_CLIENTE").Value, mdiPrincipal.txtDLVTelefono.Text)
    If rsTelefono.EOF() Then
        Dim strMensaje As String
        Dim rsSecTelefono As oraDynaset
        Set rsSecTelefono = objCliente.Ultima_Sec_Telefono(ctlGrilla1.DataSource("COD_CLIENTE").Value)
        strMensaje = objCliente.GrabarSoloLineas(ctlGrilla1.DataSource("COD_CLIENTE").Value, _
                                                 "001", IIf(IsNull(rsSecTelefono("SEC_LINEA_CEN").Value), "0", rsSecTelefono("SEC_LINEA_CEN").Value) + 1, Trim(mdiPrincipal.txtDLVTelefono.Text))
    End If
    'I.ECASTILLO 17.12.2020 | dc Cappa | 2da etapa reserva
'    Dim flg_2e_reserva
'    flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV2") '1 => ACTIVO, 0 => INACTIVO
'    Dim sCia As String
'    Dim rsCia As oraDynaset
'    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, mdiPrincipal.ctlCliente1.LocalDespacho)
'    If (rsCia.RecordCount > 0) Then
'      sCia = CStr(rsCia(1))
'    End If
'    Set rsCia = Nothing
'    objVenta.flg_2e_reserva_local = objLocal.GetEstConfig(sCia, mdiPrincipal.ctlCliente1.LocalDespacho, "RESERVA_STOCK_2DA")
    If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
        GoTo geo:
    End If
    If flg_2e_reserva = "0" Then
        GoTo geo:
    Else
'        objVenta.isLocalDcCappa = "" & objLocal.GetIndCDCAP(mdiPrincipal.ctlCliente1.LocalDespacho)
        'If objVenta.isLocalDcCappa = "0" And objVenta.flg_2e_reserva_local = "1" Then
        If objVenta.flg_2e_reserva_local = "1" Then
CallServiceReserva:
            If mdiPrincipal.ctlCliente1.seleccionManualLocal = True Then
                frm_VTA_MetodosSegmentos.Parametro = objLocal.GetCodPosu(mdiPrincipal.ctlCliente1.LocalDespacho)
                frm_VTA_MetodosSegmentos.Tipo = 3
                frm_VTA_MetodosSegmentos.permiteCerrar = "0"
                frm_VTA_MetodosSegmentos.Show vbModal
            Else
                GoTo geo:
            End If
        Else
        
            'flgFunLocal = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVET3") '1 => ACTIVO, 0 => INACTIVO
            If flg_2e_reserva = "1" And flgFunLocal = "1" Then
                GoTo CallServiceReserva
            End If
geo:
            If objVenta.CodigoCliente <> "" Then
                frm_VTA_PreviaTomaPedido.Show vbModal
            End If
        End If
    End If
    'F.ECASTILLO 12.2020
    '*******'
'    If objVenta.CodigoCliente <> "" Then
'        frm_VTA_PreviaTomaPedido.Show vbModal
'    End If
    '*******'
    Unload Me
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdBuscar_Click()
On Error GoTo Control
    Set ctlGrilla1.DataSource = Nothing
    Call SetGrd
    If frm_DLV_BuscaTelefono.Telefono = "" Then
        Set ctlGrilla1.DataSource = objCliente.ListaCliente(txtNombre.Text, "", "", "", "", "")
      Else
        Set ctlGrilla1.DataSource = objCliente.ListaCliente(txtNombre.Text, "", "", "", "", frm_DLV_BuscaTelefono.Telefono)
    End If
    ctlGrilla1.Rebind 'ECASTILLO 17.12.2020
    ctlGrilla1.Refresh
    ctlGrilla1.SetFocus
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdCancelar_Click()
    'mdiPrincipal.ctlCliente1.Limpiar
    Unload Me
End Sub

Private Sub cmdSoloPrecios_Click()
On Error GoTo Control

    frm_DLV_BuscaDistrito.Show vbModal
    If Not frm_DLV_BuscaDistrito.bolCancelar Then Unload Me
   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
    
End Sub

Private Sub Command1_Click()
    Dim xForm As Form
    
    On Error GoTo Control

    If ctlGrilla1.ApproxCount = 0 Then Exit Sub
    '''  frm_VTA_Cliente.ctlCliente1.Cargar
    
    If objUsuario.flgDeliveryProv = "1" Then
        Set xForm = New frm_VTA_Cliente_Prov
    Else
        Set xForm = New frm_VTA_Cliente
    End If
    With xForm
        .Telefono = Telefono
        .ctlCliente1.XTipoFuncion = "Editar"
        .ctlCliente1.CodDireccionCli = "" & grdDireccion.Columns("COD_DIRECCION_CLI").Value
        .strCodigo = ctlGrilla1.Columns(0).Value
        .CargarValores
        ''''.ctlCliente1.ConsultaCliente ctlGrilla1.Columns(0).Value
        .Show vbModal
        'cmdBuscar_Click 'ECASTILLO 17.12.2020
    End With
   
    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Command2_Click()
    Dim xForm As Form

    On Error GoTo Control

    '''    frm_VTA_Cliente.ctlCliente1.Cargar Telefono
    If objUsuario.flgDeliveryProv = "1" Then
        Set xForm = New frm_VTA_Cliente_Prov
    Else
        Set xForm = New frm_VTA_Cliente
    End If

    With xForm
        .ctlCliente1.Codigo = ""
        .ctlCliente1.XTipoFuncion = "Nuevo"
        .Telefono = Telefono
        .CargarValores
        .Show vbModal
   End With
   Set xForm = Nothing
   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
    Set xForm = Nothing
End Sub

Private Sub ctlGrilla1_DblClick()
    cmdAceptar_Click
End Sub

Private Sub ctlGrilla1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then ctlGrilla1_DblClick
End Sub

Private Sub ctlGrilla1_RegistroSeleccionado(ByVal DatoColumna0 As String)
    BuscaDireccion
End Sub
Sub BuscaDireccion()
Dim objCliente As New clsCliente
Dim strNombre As String
    strNombre = "" & Trim(ctlGrilla1.Columns("COD_CLIENTE").Value)
    If Len(strNombre) > 3 Then
        Set grdDireccion.DataSource = objCliente.ListaDireccion(strNombre) 'strflgActivo)
        grdDireccion.Columns("COD_DIRECCION_CLI").Visible = False
        grdDireccion.Columns("COD_UBIGEO").Visible = False
        'grdDireccion.Columns("DES_DISTRITO").Visible = False
        'grdDireccion.Columns("DES_URBANIZACION").Visible = False
    End If
    Set objCliente = Nothing
End Sub
Private Sub Form_Activate()
    Dim Cia As String
    If objUsuario.CodLocalCallCenter = "1DLV" Then
        Cia = "94"
    Else
        Cia = objUsuario.CodigoEmpresa
    End If
    Set ctlGrilla1.DataSource = objCliente.ListaTelefono(Telefono, Cia)
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            Command2_Click
        Case vbKeyF2
             Command1_Click
        Case vbKeyF3
            cmdAceptar_Click
        Case vbKeyF4
            txtNombre.SetFocus
        Case vbKeyF5
            ctlGrilla1.SetFocus
        Case vbKeyF6
            grdDireccion.SetFocus
        Case vbKeyF7
            cmdSoloPrecios_Click
    End Select
End Sub
''''Private Sub Form_Paint()
''''    If flgDescarga = True Then
''''        frm_VTA_Cliente.Show vbModal
''''        If Not objVenta.CodCliente = "" Then cargaCodigo objVenta.CodCliente
''''        flgDescarga = False
''''    End If
''''End Sub

Private Sub Form_Load()
    Dim Cia As String
    If objUsuario.CodLocalCallCenter = "1DLV" Then
        Cia = "94"
    Else
        Cia = objUsuario.CodigoEmpresa
    End If
    Call SetGrd
    
    Set ctlGrilla1.DataSource = objCliente.ListaTelefono(Telefono, Cia)
    frm_VTA_PreviaTomaPedido.flgContinua = False
    
'''    If ctlGrilla1.DataSource.RecordCount = 0 Then
'''        frm_VTA_Cliente.Telefono = Telefono
'''        frm_VTA_Cliente.Show vbModal
'''        If Not objVenta.CodigoCliente = "" Then cargaCodigo objVenta.CodigoCliente
'''        Exit Sub
'''    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not frm_VTA_PreviaTomaPedido.flgContinua Then
        mdiPrincipal.ctlCliente1.Limpiar
        LimpiaClienteObjVenta
        mdiPrincipal.printCantCallsGMaps "", "Limpia Datos Cliente"
    End If
'    Set objCliente = Nothing
'    Debug.Print objVenta.CodModalidadVenta
End Sub

Public Sub cargaCodigo(Codigo)
    Set ctlGrilla1.DataSource = objCliente.Lista(Codigo)
    flgDescarga = False
End Sub

Private Sub grdDireccion_DblClick()
    ctlGrilla1_DblClick
End Sub

Private Sub grdDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control

    If KeyCode = vbKeyReturn Then ctlGrilla1_DblClick
    If KeyCode = vbKeyDelete Then
        EliminarDireccion
        BuscaDireccion
    End If

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub
Sub EliminarDireccion()
    Dim strConfimacion As String
    If MsgBox("Seguro que desea eliminar esta dirección ?", vbYesNo + vbQuestion, "Confirme") = vbYes Then
        Dim objCliente As New clsCliente
        Dim strMensaje As String
        strMensaje = objCliente.AnularDireccion(grdDireccion.DataSource("COD_CLIENTE"), grdDireccion.DataSource("COD_DIRECCION_CLI"), grdDireccion.DataSource("COD_TIPO_DIRECCION"))
        If Not strMensaje = "" Then
            MsgBox strMensaje, vbCritical, App.ProductName
        End If
        Set objCliente = Nothing
    End If
End Sub
Private Sub SetGrd()
    Dim arrField As Variant
    Dim arrCaption As Variant
    Dim arrWidth As Variant
    Dim arrAlignment As Variant
    
    arrField = Array("COD_CLIENTE", "SUFIJO", "XNOMBRE", "DES_OBSERVACION", "COD_LOCAL_DESPACHO", "FLG_TIPO_JURIDICA", "DES_RAZON_SOCIAL", "DES_DIRECCION_COMERCIAL", "DES_DIRECCION_SOCIAL", "NUM_DOCUMENTO_ID", "DES_LOCAL_PRECIO", "DES_LOCAL_DESPACHO", "DES_RAZON_SOCIAL", "DES_NOM_COMERCIAL", "DES_NOM_CLIENTE", "DES_APE_CLIENTE", "DES_APE2_CLIENTE", "FLG_CLIENTE_VERIFICADO", "SUFIJO", "COD_LOCAL_PRECIO")
    arrCaption = Array("Cliente", "Sufijo", "Nombre", "Observaciones", "Despacho", "Juridica", "Razón Social", "Dirección Comercial", "Dirección Social", "DNI", "Local Despacho", "Local Precio", "Razón Social", "Nombre Comercial", "Nombre", "Apellido Paterno", "Apellido Materno", "Verificado", "Sufijo", "Local Precio")
    arrWidth = Array(1000, 500, 4000, 15000, 700, 500, 1000, 1500, 1500, 900, 800, 800, 2000, 2000, 1500, 1000, 1000, 800, 500, 700)
    arrAlignment = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter)
    ctlGrilla1.FormatoGrilla arrField, arrCaption, arrWidth, arrAlignment
    ctlGrilla1.Columns(0).Visible = True
    
    arrField = Array("DES_TIPO_DIRECCION", "DES_DIRECCION", "DES_REFERENCIA_DIRECCION", "COD_DIRECCION_CLI", "COD_UBIGEO", "DES_DISTRITO", "DES_URBANIZACION", "LATITUD", "LONGITUD", "DES_ABREVIATURA_DIRECCION", "DES_SUFIJO_DIRECCION")
    arrCaption = Array("Tipo", "Direccción", "Referencia", "COD_DIRECCION_CLI", "COD_UBIGEO", "Distrito", "Urbanización", "LATITUD", "LONGITUD", "DES_ABREVIATURA_DIRECCION", "DES_SUFIJO_DIRECCION")
    arrWidth = Array(800, 2500, 4000, 0, 0, 1500, 1800, 0, 0, 0, 0)
    arrAlignment = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgCenter)
    grdDireccion.FormatoGrilla arrField, arrCaption, arrWidth, arrAlignment
    grdDireccion.Columns("LATITUD").Visible = False
    grdDireccion.Columns("LONGITUD").Visible = False
    grdDireccion.Columns("DES_ABREVIATURA_DIRECCION").Visible = False
    grdDireccion.Columns("DES_SUFIJO_DIRECCION").Visible = False
End Sub

Public Function LimpiaClienteObjVenta()
    objVenta.Latitud = ""
    objVenta.Longitud = ""
        
    objVenta.NombreClienteDLV = ""
    objVenta.DesAuxCliDirecc = ""
    objVenta.NumeroDocumentoID = ""
        
        '*************************************************************************************'
    objVenta.DireccionClienteDLV = ""
    objVenta.UbigeoEntrega = ""
    objVenta.DesReferenciaCli = ""
    objVenta.DesAuxCliTlf = ""
    objVenta.DesDistritoDLV = ""
    objVenta.DesUrbanizacionDLV = ""
    objVenta.ObservacionClienteDLV = ""
    objVenta.CodDireccionCli = ""
        
    objVenta.Out_NumeroId = ""
    objVenta.Out_NombreCliente = ""
    objVenta.Out_Tipo = ""
    objVenta.Out_CodDocumentoId = ""
    objVenta.Out_Direccion = ""
    
    objVenta.CodigoCliente = ""
    
    objVenta.CodigoClienteDLV = ""
    objVenta.EntregaTercero = "0"
    objVenta.bk_Ubigeo = ""
End Function


