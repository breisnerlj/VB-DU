VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_VTA_CantidadProducto 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   15945
   ClientLeft      =   7065
   ClientTop       =   1575
   ClientWidth     =   6915
   ControlBox      =   0   'False
   Icon            =   "frm_VTA_CantidadProducto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15945
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCantidadFrac 
      Height          =   315
      Left            =   1920
      TabIndex        =   32
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtStockRealFrac_V1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdEliminar 
      Height          =   495
      Left            =   5160
      Picture         =   "frm_VTA_CantidadProducto.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton cmdAgregar 
      Height          =   495
      Left            =   5160
      Picture         =   "frm_VTA_CantidadProducto.frx":0894
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3960
      Width           =   615
   End
   Begin vbp_Ventas.ctlGrillaArray ctlGrillaArray2 
      Height          =   2565
      Left            =   360
      TabIndex        =   18
      Top             =   8160
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4524
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlDataCombo CboNroLote 
      Height          =   315
      Left            =   1680
      TabIndex        =   12
      Top             =   3960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
   End
   Begin vbp_Ventas.ctlGrillaArray ctlGrillaArray1 
      Height          =   735
      Left            =   5040
      TabIndex        =   13
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin MSMask.MaskEdBox txtFechaVencimiento 
      Height          =   315
      Left            =   1680
      TabIndex        =   14
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin vbp_Ventas.ctlTextBox txtNroLote 
      Height          =   315
      Left            =   1800
      TabIndex        =   11
      Top             =   3960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      MaxLength       =   20
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
   Begin vbp_Ventas.ctlTextBox txtCantidad 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   2720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Tipo            =   3
      MaxLength       =   4
      TABAuto         =   0   'False
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
   Begin vbp_Ventas.ctlGrillaArray ctlGrillaArray3 
      Height          =   2565
      Left            =   360
      TabIndex        =   17
      Top             =   5400
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4524
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.CheckBox chkFraccionamiento 
      Caption         =   "&Fraccionamiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label txtStockRealFrac 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   315
      Left            =   1920
      TabIndex        =   33
      Top             =   2330
      Width           =   1095
   End
   Begin VB.Label lblCantidadFrac 
      Caption         =   "Cantidad Frac."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   3150
      Width           =   1455
   End
   Begin VB.Label lblStockRealFrac 
      Caption         =   "Stock Real Frac."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   2350
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Sync"
      Height          =   240
      Left            =   5040
      TabIndex        =   28
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblSemaforo 
      Height          =   15
      Left            =   6000
      TabIndex        =   27
      Top             =   1950
      Width           =   255
   End
   Begin VB.Label lblFechaServ 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4080
      TabIndex        =   26
      Top             =   2085
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Ctrl+R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   3120
      TabIndex        =   25
      Top             =   1950
      Width           =   615
   End
   Begin VB.Label lblStockRealCant 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   315
      Left            =   1920
      TabIndex        =   24
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblStockReal 
      AutoSize        =   -1  'True
      Caption         =   "Stock Real :"
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
      Left            =   240
      TabIndex        =   23
      Top             =   1950
      Width           =   1080
   End
   Begin VB.Label lblCantSugerida 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   6480
      TabIndex        =   22
      Top             =   3600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblSugerido 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   3585
      Width           =   5955
   End
   Begin VB.Label lblFraccionamiento 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   5310
      TabIndex        =   20
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label lblCodigoSap 
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3120
      TabIndex        =   19
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Vencimiento"
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
      Left            =   240
      TabIndex        =   10
      Top             =   4590
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nro de Lote"
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
      Left            =   240
      TabIndex        =   9
      Top             =   3990
      Width           =   1050
   End
   Begin VB.Label lblIndicador 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3960
      TabIndex        =   8
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad :"
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
      Left            =   240
      TabIndex        =   6
      Top             =   2750
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripción :"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1170
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Código :"
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
      Left            =   240
      TabIndex        =   4
      Top             =   750
      Width           =   750
   End
   Begin VB.Label lblCodigo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000040&
      Height          =   555
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frm_VTA_CantidadProducto.frx":0E1E
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad del producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frm_VTA_CantidadProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objProducto As New clsProducto
Dim objConvenio As New clsConvenio
Dim objLocal As New clsLocal
Dim objWS As New clsWebService
Dim oraDato As oraDynaset
Public flgStockTotal As Boolean
Public flgCerrar As Boolean
''---------------------------------------'
''-- Variables del Recetario Magistral --'
Dim strCadCodIns As String
Dim strCadDesIns As String
Dim strCadCodigo As String
Dim strCadProd As String
Dim strCadCodUnd As String
Dim strCadUnd As String
Dim strCadPctMargen As String
Dim strCtdBase As String
Dim strCadCant As String
Dim strCadPrecio As String
Dim strCadSubTotal As String
Public strCantAnt As String
Dim strFlgReceta As String
Public strStock As String
'-----------------------------
Public strLaboratorio As String
Public strStockRealCant As String
Public strCodLocalPosu As String
''---------------------------------'
Dim strCod$, strDes$
Dim strTipoPrecio As String
Dim cRegalo As eProdRegalo
Public flgEspecieValorada As String
Public pBlnModal As Boolean ''Permite el control de las teclas de funcion'
Dim strFracciona As String
Private pbytIntentos As Byte
Dim vstrStock As String
Dim indice As Integer

Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant

Dim xTempNroLote As New XArrayDB
Dim xTempNroLotePedido As New XArrayDB

Dim xValidacionesProducto As New XArrayDB
Dim rs As OracleInProcServer.oraDynaset

Dim FracEnabled As String
Dim arrInfo As New XArrayDB

'I.ECASTILLO 22.10.2020
'Dim flgDCCAP As String
Dim valFracDcCap, unidDcCap As String
Dim CTDFRAC2 As Double
Dim res As String
'F.ECASTILLO 22.10.2020
Dim valFracMaxDc As String

Public Sub subDatos(strCodProd As String, strDesProd As String, ByVal pstrTipoPrecio As String, pstrCaption As String, Regalo As eProdRegalo, strIdFrac As String, strIndicador As String, Optional strFlgRec As String = "", Optional strStockq As String, Optional strLote As String = "", Optional strFchVcto As String = "__/__/____", Optional strCodigoSAP As String = "", Optional Laboratorio As String = "")
    valFracDcCap = "": unidDcCap = "": valFracMaxDc = ""
    lblCodigo.Caption = strCodProd
    lblFraccionamiento.Caption = ""
   
   If strCodigoSAP = "" Then
    lblCodigoSap.Caption = "" & objProducto.DevCodSap(strCodProd)
Else
    lblCodigoSap.Caption = strCodigoSAP
End If
    lblDescripcion.Caption = strDesProd
    strTipoPrecio = pstrTipoPrecio
    lblIndicador.Caption = strIndicador
    strFlgReceta = strFlgRec
    chkFraccionamiento.Enabled = True
    strFracciona = strIdFrac
    strStock = strStockq
    
    If strIdFrac = "0" Then chkFraccionamiento.Enabled = False
    cRegalo = Regalo
    If cRegalo = Producto_Regalo Then
        txtCantidad.Bloqueado = True
        If strIdFrac = "1" Then chkFraccionamiento.Enabled = False
    End If
    
    If cRegalo = Producto_Regalo_precio Then
        txtCantidad.Bloqueado = False
        If strIdFrac = "1" Then chkFraccionamiento.Enabled = True
    End If
    frm_VTA_CantidadProducto.Height = 3680 '2980 '2580 '10000 '2580
        
        If frm_VTA_CantidadProducto.Height = 2580 Then
            frm_VTA_CantidadProducto.top = 4000
        ElseIf frm_VTA_CantidadProducto.Height = 2980 Then
            frm_VTA_CantidadProducto.top = 3700
        ElseIf frm_VTA_CantidadProducto.Height = 3680 Then
            frm_VTA_CantidadProducto.top = 3400
        End If
        
    txtCantidad.TABAuto = False
    If objConvenio.Lote(objVenta.CodModalidadVenta, objVenta.CodigoConvenio) = "1" Then
        Label4.Visible = False
        Label6.Visible = False
        txtNroLote.Visible = False
        
        CboNroLote.Visible = False
        
        txtFechaVencimiento.Visible = False
        Dim rsLot As oraDynaset
        Set rsLot = objVenta.ListaVentaxLotes(objVenta.CodModalidadVenta, objVenta.CodigoConvenio, "")
                        
        If "" & rsLot("FLG_CON_LOTE") = "1" Then
            CboNroLote.Visible = False
            Label4.Visible = True
            txtNroLote.Visible = True
            txtNroLote.Text = strLote
            
    'Seleccionar Nro de Lote MLevano 27/08/2012
    '
    '===========================================================================
    If objVenta.CodigoTipoVenta = Guias_Remision Then
        Set CboNroLote.RowSource = objProducto.ListaLote(strCodProd, objUsuario.CodigoLocal) 'objUsuario.ListaUsuarioDLV
            CboNroLote.ListField = "NLOTE"
            CboNroLote.BoundColumn = "FVENC"
            
            'CboNroLote.BoundText = "*"
            CboNroLote.Visible = True
            txtNroLote.Visible = False
    End If
    '===========================================================================
            
        End If
        
        If "" & rsLot("FLG_CON_FECVEN") = "1" Then
            Label6.Visible = True
            txtFechaVencimiento.Visible = True
'            txtFechaVencimiento.Text = strFchVcto
        End If
       
        If objVenta.CodigoTipoVenta = Guias_Remision Then
            frm_VTA_CantidadProducto.Height = 7665 '9500 '7665 '3960
            txtCantidad.TABAuto = True
            cmdAgregar.Visible = True
            cmdEliminar.Visible = True
GoTo x
        End If
            frm_VTA_CantidadProducto.Height = 3960
            txtCantidad.TABAuto = True
x:
        
    End If
            If objUsuario.MayoristaFracciona = False And objVenta.CodigoTipoVenta = Venta_Mayorista Then
    
        chkFraccionamiento.Enabled = False
    End If
    If objProducto.ParametroSugerido = 3 And objUsuario.EsDelivery = True Then
       Dim rs As oraDynaset
        Set rs = objProducto.ListaFraccionamientoSugerido(strCodProd)
        lblSugerido.Caption = "RECUERDE OFRECER EN " & rs("UNID_VTA").Value
    End If
    Me.Caption = pstrCaption
    'I. ECASTILLO 05.07.2020 | 13.01.2021 agregar flg para activar
'    flgCerrar = False
    Dim flgFun
    flgFun = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV") '1 => ACTIVO, 0 => INACTIVO
    If flgFun = "1" Then
        StockReal (strCodProd)
    ElseIf flgFun = "0" Then
        StockBD (strCodProd)
    End If
    'If flgCerrar Then Exit Sub
    'F. ECASTILLO 05.07.2020
    
    
    'I.ECASTILLO 22.10.2020 | 13.01.2021 agregar flg para activar
    'si local es DC CAP valida fraccionamiento multiplo o igual
     'flgDCCAP = objLocal.GetIndCDCAP(mdiPrincipal.ctlCliente1.LocalDespacho)
     If objVenta.isLocalDcCappa = "1" Then
        Dim cntSug_old As String
        Dim rs2 As oraDynaset
        Set rs2 = objProducto.ListaFraccionamientoSugeridoV2(lblCodigo.Caption, mdiPrincipal.ctlCliente1.LocalDespacho)
        'cntSug_old = lblCantSugerida.Caption
        unidDcCap = "" & rs2("UNID_VTA").Value
        valFracDcCap = "" & rs2("VAL_FRAC_LOCAL").Value
        'I.ECASTILLO 12.08.2021
'        If gstrIndDCSAP = "1" Then
'            Dim confStockXTextDC As String
'            Dim LocalDespacho As String
'            Dim strCiaDespacho As String
'            LocalDespacho = mdiPrincipal.ctlCliente1.LocalDespacho
'            strCiaDespacho = mdiPrincipal.ctlCliente1.sCia
'
'            If strCiaDespacho = "" Then strCiaDespacho = objUsuario.CodigoEmpresa
'            confStockXTextDC = objLocal.GetEstConfig(strCiaDespacho, LocalDespacho, "STOCK_DC_SAP")
'            If confStockXTextDC = "1" Then
'                valFracMaxDc = "" & rs2("VAL_MAX_FRAC").Value
'            End If
'        End If
        If Len(Trim(valFracDcCap)) = 0 Then chkFraccionamiento.Enabled = False: strFracciona = 0
        'CTDFRAC = objProducto.DevuelveCTDFRAC(lblCodigo.Caption)
        'res = Val(CTDFRAC) Mod Val(SugDcCap)
        'If res <> 0 Then
        '    lblCantSugerida.Caption = cntSug_old
        'End If
        If gstrIndDCSAP = "1" Then
            If objVenta.isDCSAP = "1" Then
                lblFraccionamiento.Caption = unidDcCap
                lblFraccionamiento.left = chkFraccionamiento.left
                If Len(Trim(valFracDcCap)) = 0 Then valFracDcCap = 1
            End If
        End If
     End If
    'F.ECASTILLO 22.10.2020
    Me.Show vbModal

End Sub
'I.ECASTILLO 13.01.2021
Private Function StockBD(ByVal codProducto As String)
On Error GoTo Err
    txtCantidad.BackColor = &HFFFFFF
    Dim IsFractional As Integer
    Dim codProductoPosu, codLocalPosu As String
    codProductoPosu = objProducto.GetCodPosu(codProducto)
    codLocalPosu = objLocal.GetCodPosu(mdiPrincipal.ctlCliente1.LocalDespacho)
    
    strStockRealCant = 0
    Dim arrResp As New XArrayDB
    Dim rsData As oraDynaset
    arrInfo.ReDim 0, -1, 0, 4
    
    Set rsData = objProducto.ListaStockLocal(codLocalPosu, codProductoPosu)
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        Dim ultimo As Integer
        While Not rsData.EOF
            IsFractional = "" & rsData("FLG_FRACCIONAMIENTO")
            If IsFractional <> "0" Then 'ENTERO - FRACCION
                ultimo = arrInfo.Count(1)
                arrInfo.AppendRows
                arrInfo(ultimo, 0) = "PACK_MODE"
                arrInfo(ultimo, 1) = "" & rsData("STOCK_ENTERO")
                arrInfo(ultimo, 2) = "" & rsData("VAL_FRAC_LOCAL")
                '
                ultimo = arrInfo.Count(1)
                arrInfo.AppendRows
                arrInfo(ultimo, 0) = "PART_MODE"
                arrInfo(ultimo, 1) = "" & rsData("STOCK_FRACCION")
                arrInfo(ultimo, 2) = "" & rsData("VAL_FRAC_LOCAL")
            End If
            
            If IsFractional = "0" Then 'ENTERO
                ultimo = arrInfo.Count(1)
                arrInfo.AppendRows
                arrInfo(ultimo, 0) = "PACK_MODE"
                arrInfo(ultimo, 1) = "" & rsData("STOCK_ENTERO")
                arrInfo(ultimo, 2) = "" & rsData("VAL_FRAC_LOCAL")
            End If
            rsData.MoveNext
        Wend
        
        Dim x As Integer
        If IsFractional = 1 Then chkFraccionamiento.Enabled = True Else chkFraccionamiento.Enabled = False
        If chkFraccionamiento.Value = 1 Then
            For x = 0 To arrInfo.Count(1) - 1
                If arrInfo(x, 0) <> "PACK_MODE" Then lblStockRealCant = arrInfo(x, 1)
            Next x
        Else
            For x = 0 To arrInfo.Count(1) - 1
                If arrInfo(x, 0) = "PACK_MODE" Then lblStockRealCant = arrInfo(x, 1)
            Next x
        End If
    End If
    'I.ECASTILLO 12.08.2021
    If gstrIndDCSAP = "1" Then
        If objVenta.isDCSAP = "1" Then
            lblStockReal.Caption = "Stock Real U."
            lblStockRealFrac.Caption = "Stock Real Frac."
            Label3.Caption = "Cantidad U."
            lblCantidadFrac.Caption = "Cantidad Frac."
            chkFraccionamiento.Visible = False
            lblStockRealFrac.Visible = True
            txtStockRealFrac.Visible = True
            lblCantidadFrac.Visible = True
            txtCantidadFrac.Visible = True
            'Label3.top = lblStockRealFrac.top + Label3.Height + 120
            'txtCantidad.top = txtStockRealFrac.top + txtCantidad.Height + 120
            'frm_VTA_CantidadProducto.Height = frm_VTA_CantidadProducto.Height + (txtStockRealFrac.Height + txtCantidadFrac.Height + 240)
        
            For x = 0 To arrInfo.Count(1) - 1
                If arrInfo(x, 0) = "PACK_MODE" Then
                    lblStockRealCant.Caption = arrInfo(x, 1)
                ElseIf arrInfo(x, 0) = "PART_MODE" Then
                    txtStockRealFrac.Caption = arrInfo(x, 1)
                End If
            Next x
            If Len(Trim(lblStockRealCant.Caption)) = 0 Then txtCantidad.Enabled = False: txtCantidad.BackColor = &HE0E0E0
            If Len(Trim(txtStockRealFrac.Caption)) = 0 Then txtCantidadFrac.Enabled = False: txtCantidadFrac.BackColor = &HE0E0E0
        Else
            GoTo prevDCSAP
        End If
    Else
prevDCSAP:
        lblStockReal.Caption = "Stock Real"
        lblStockRealFrac.Caption = "Stock Real Frac."
        Label3.Caption = "Cantidad :"
        lblCantidadFrac.Caption = "Cantidad Frac."
        chkFraccionamiento.Visible = True
        lblStockRealFrac.Visible = False
        txtStockRealFrac.Visible = False
        lblCantidadFrac.Visible = False
        txtCantidadFrac.Visible = False
        Label3.top = lblStockRealFrac.top
        txtCantidad.top = txtStockRealFrac.top
        frm_VTA_CantidadProducto.Height = frm_VTA_CantidadProducto.Height - (txtStockRealFrac.Height + txtCantidadFrac.Height + 240)
        If IsFractional = 1 Then chkFraccionamiento.Enabled = True Else chkFraccionamiento.Enabled = False:
        If chkFraccionamiento.Value = 1 Then
            For x = 0 To arrInfo.Count(1) - 1
                If arrInfo(x, 0) <> "PACK_MODE" Then lblStockRealCant = arrInfo(x, 1)
            Next x
        Else
            For x = 0 To arrInfo.Count(1) - 1
                If arrInfo(x, 0) = "PACK_MODE" Then lblStockRealCant = arrInfo(x, 1)
            Next x
        End If
    End If
    'F.ECASTILLO 12.08.2021
    strStockRealCant = lblStockRealCant
    strCodLocalPosu = codLocalPosu
    Exit Function
Err:
    Err.Raise Err.Number, "frm_VTA_CantidadProducto", Err.Description
End Function

'F.ECASTILLO 13.01.2021
'I. ECASTILLO 05.07.2020
Private Function StockReal(ByVal codProducto As String)
On Error GoTo Err
    Dim obj As New Dictionary
    Dim ArrCode As Variant
    Dim x, i As Integer
    Dim IsFractional As Integer
    Dim ZeroStock As Boolean
    Dim objProducto As New clsProducto
    Dim CTDFRAC_WS As Double
    Dim rsCia As oraDynaset
    Dim sCia As String
    Dim sCodLoc As String
    txtCantidad.BackColor = &HFFFFFF
    Dim codProductoPosu, codLocalPosu As String
    codProductoPosu = objProducto.GetCodPosu(codProducto)
    codLocalPosu = objLocal.GetCodPosu(mdiPrincipal.ctlCliente1.LocalDespacho)
    ArrCode = Array(codProductoPosu)
    'I.ECASTILLO 28.01.2021 | se reversa el cambio realizado por gnibin
    'Set obj = objWS.GetStockRealWS(objUsuario.CodLocalCallCenter, codLocalPosu, ArrCode) 'GNIBIN 20210127 Proyecto Multimarca
    sCodLoc = mdiPrincipal.ctlCliente1.LocalDespacho
    
    'I.ECASTILLO 17.09.2021 PARAMETRIZAR MARCA
    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, sCodLoc)
    If (rsCia.RecordCount > 0) Then
      sCia = CStr(rsCia(1))
    End If
    Set rsCia = Nothing
    Dim strCia As String
    strCia = "" & objLocal.GetMarcaLocal(sCodLoc, 1)
    strCia = Trim(strCia)
    'F.ECASTILLO 17.09.2021
    
    'Set obj = objWS.GetStockRealWS(mdiPrincipal.ctlCliente1.sCia, codLocalPosu, ArrCode) 'GNIBIN 20210127 Proyecto Multimarca
    
    'I.ECASTILLO 12.08.2021
    If gstrIndDCSAP = "1" Then
        If objVenta.isDCSAP = "1" Then
            Set obj = objWS.GetStockRealWSDCSAP(sCia, codLocalPosu, ArrCode, strCia)
        Else
            Set obj = objWS.GetStockRealWS(sCia, codLocalPosu, ArrCode, strCia)
        End If
    Else
        Set obj = objWS.GetStockRealWS(sCia, codLocalPosu, ArrCode, strCia) 'GNIBIN 20210127 Proyecto Multimarca
    End If
    'F.ECASTILLO 12.08.2021
    'F.ECASTILLO 28.01.2021
    lblStockRealCant = 0
    arrInfo.ReDim 0, -1, 0, 4
    If obj.Count > 0 Then
'        arrInfo.ReDim 0, -1, 0, 2
        Dim rngV, rngA, rngR
        Dim arrRngV, arrRngA, arrRngR
        rngV = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "RNG0001")
        rngA = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "RNG0002")
        rngR = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "RNG0003")
        arrRngV = Split(rngV, "|")
        arrRngA = Split(rngA, "|")
        arrRngR = Split(rngR, "|")
        Debug.Print obj("data").Count()
        For x = 1 To obj("data").Count()
            IsFractional = obj("data")(x)("isFractional"): CTDFRAC_WS = 1
            For i = 1 To obj("data")(x)("fractionType").Count()
                arrInfo.AppendRows
                arrInfo(i - 1, 0) = obj("data")(x)("fractionType")(i)("fractionatedText")
                arrInfo(i - 1, 1) = obj("data")(x)("fractionType")(i)("stock")
                'SI ES FRACCIONABLE O NO, SIEMPRE SE DEBE MOSTRAR unitQuantity DE PACK_MODE
                If arrInfo(i - 1, 2) = "" Then arrInfo(i - 1, 2) = 1
                If arrInfo(i - 1, 0) = "PACK_MODE" Then
                     arrInfo(i - 1, 2) = obj("data")(x)("fractionType")(i)("unitQuantity")
                End If
                If CTDFRAC_WS = 1 Then CTDFRAC_WS = arrInfo(i - 1, 2)
                'If arrInfo(i - 1, 0) = "PART_MODE" Then
                '    arrInfo(i - 1, 3) = obj("data")(X)("fractionType")(i)("stock")
                'End If
            Next i
            'I.ECASTILLO 12.01.2021
            Dim fchLastSync, fechSysDate, fchResp
            fchLastSync = "" & obj("data")(x)("fecSymLocal")
            fechSysDate = "" & obj("data")(x)("fechSysDate")
            If fchLastSync = "" Or Len(Trim(fchLastSync)) = 0 Then fchLastSync = "2020-01-01" 'fecha default
            If fechSysDate = "" Or Len(Trim(fechSysDate)) = 0 Then fechSysDate = "2020-01-30" 'fecha default
            fchLastSync = Format(fchLastSync, "yyyy/mm/dd hh:mm:ss AMPM")
            fechSysDate = Format(fechSysDate, "yyyy/mm/dd hh:mm:ss AMPM")
            lblFechaServ = fchLastSync
            fchResp = dateDiff("n", fchLastSync, fechSysDate)
            '&H8000000F&
            lblFechaServ.BackColor = &H8000000F
            'lblFechaServ.ForeColor = &H8000000F
            fchResp = val(fchResp)
            fchResp = IIf(fchResp <= 0, 0, fchResp)
'            MsgBox "Diff: " & fchResp & vbNewLine _
'                    & rngV & " - " & rngA & " - " & rngR & vbNewLine _
'                    & arrRngV(0) & " hasta " & arrRngV(1) & vbNewLine _
'                    & arrRngA(0) & " hasta " & arrRngA(1) & vbNewLine _
'                    & arrRngR(0)
            If fchResp >= val(arrRngV(0)) And fchResp <= val(arrRngV(1)) Then '5min | verde
                lblFechaServ.BackColor = &HFF00&
                lblFechaServ.ForeColor = &H0&
            ElseIf fchResp >= val(arrRngA(0)) And fchResp <= val(arrRngA(1)) Then '10min | amarillo
                lblFechaServ.BackColor = &HFFFF&
                lblFechaServ.ForeColor = &H0&
            ElseIf fchResp >= val(arrRngR(0)) Then '>10min | rojo
                lblFechaServ.BackColor = &HFF&
                lblFechaServ.ForeColor = &HFFFFFF
            End If
            'F.ECASTILLO 12.01.2021
        Next x
        
        'I.ECASTILLO 12.08.2021
    If gstrIndDCSAP = "1" Then
        If objVenta.isDCSAP = "1" Then
            lblStockReal.Caption = "Stock Real U."
            lblStockRealFrac.Caption = "Stock Real Frac."
            Label3.Caption = "Cantidad U."
            lblCantidadFrac.Caption = "Cantidad Frac."
            chkFraccionamiento.Visible = False
            lblStockRealFrac.Visible = True
            txtStockRealFrac.Visible = True
            lblCantidadFrac.Visible = True
            txtCantidadFrac.Visible = True
            'Label3.top = lblStockRealFrac.top + Label3.Height + 120
            'txtCantidad.top = txtStockRealFrac.top + txtCantidad.Height + 120
            'frm_VTA_CantidadProducto.Height = frm_VTA_CantidadProducto.Height + (txtStockRealFrac.Height + txtCantidadFrac.Height + 240)
        
            For x = 0 To arrInfo.Count(1) - 1
                If arrInfo(x, 0) = "PACK_MODE" Then
                    lblStockRealCant.Caption = arrInfo(x, 1)
                ElseIf arrInfo(x, 0) = "PART_MODE" Then
                    txtStockRealFrac.Caption = arrInfo(x, 1)
                End If
            Next x
            If Len(Trim(lblStockRealCant.Caption)) = 0 Then txtCantidad.Enabled = False: txtCantidad.BackColor = &HE0E0E0
            If Len(Trim(txtStockRealFrac.Caption)) = 0 Then txtCantidadFrac.Enabled = False: txtCantidadFrac.BackColor = &HE0E0E0
        Else
            GoTo prevDCSAP
        End If
    Else
prevDCSAP:
        lblStockReal.Caption = "Stock Real"
        lblStockRealFrac.Caption = "Stock Real Frac."
        Label3.Caption = "Cantidad :"
        lblCantidadFrac.Caption = "Cantidad Frac."
        chkFraccionamiento.Visible = True
        lblStockRealFrac.Visible = False
        txtStockRealFrac.Visible = False
        lblCantidadFrac.Visible = False
        txtCantidadFrac.Visible = False
        Label3.top = lblStockRealFrac.top
        txtCantidad.top = txtStockRealFrac.top
        frm_VTA_CantidadProducto.Height = frm_VTA_CantidadProducto.Height - (txtStockRealFrac.Height + txtCantidadFrac.Height + 240)
        If IsFractional = 1 Then chkFraccionamiento.Enabled = True Else chkFraccionamiento.Enabled = False:
        If chkFraccionamiento.Value = 1 Then
            For x = 0 To arrInfo.Count(1) - 1
                If arrInfo(x, 0) <> "PACK_MODE" Then lblStockRealCant = arrInfo(x, 1)
            Next x
        Else
            For x = 0 To arrInfo.Count(1) - 1
                If arrInfo(x, 0) = "PACK_MODE" Then lblStockRealCant = arrInfo(x, 1)
            Next x
        End If
    End If
    'F.ECASTILLO 12.08.2021
        
        Dim ctdFrac As Double
        
        ctdFrac = objProducto.DevuelveCTDFRAC(lblCodigo.Caption)
        
'        If Not CTDFRAC = CTDFRAC_WS Then 'WMORI 23/09/2020 SE COMENTA POR SOLICITUD DE BREISNER.
'            chkFraccionamiento.Enabled = False
'            txtCantidad.Enabled = False
'            txtCantidad.BackColor = &H8000000F
'            MsgBox "Las cantidades de fraccionamiento no son iguales" & vbCrLf & "RAC:" & CTDFRAC & vbCrLf & "WS:" & CTDFRAC_WS, vbExclamation + vbOKOnly, "Cuidado"
'        End If
        
    Else
        chkFraccionamiento.Enabled = False
        txtCantidad.Enabled = False
        txtCantidad.BackColor = &H8000000F
        lblStockReal.Caption = "Stock Real"
        
        lblStockRealFrac.Caption = "Stock Real Frac."
        Label3.Caption = "Cantidad :"
        lblCantidadFrac.Caption = "Cantidad Frac."
        chkFraccionamiento.Visible = True
        lblStockRealFrac.Visible = False
        txtStockRealFrac.Visible = False
        lblCantidadFrac.Visible = False
        txtCantidadFrac.Visible = False
        Label3.top = lblStockRealFrac.top
        txtCantidad.top = txtStockRealFrac.top
        frm_VTA_CantidadProducto.Height = frm_VTA_CantidadProducto.Height - (txtStockRealFrac.Height + txtCantidadFrac.Height + 240)
        
        MsgBox "El servicio no ha devuelto productos", vbExclamation + vbOKOnly, "Verificar"
        'web service solo retorna productos con stock > 0;
        'si no retorna puede que no exista o tenga stock 0
'        ZeroStock = True
'        GoTo errorZero
    End If
    strStockRealCant = lblStockRealCant
    strCodLocalPosu = codLocalPosu
    Exit Function
Err:
    Err.Raise Err.Number, "frm_VTA_CantidadProducto", Err.Description
End Function
'F. ECASTILLO 05.07.2020

'Seleccionar FechaVencimiento segun el Nro de Lote MLevano
'Se asigno el valor al txtNroLote para no modificar la secuencia al grabar
Private Sub CboNroLote_Change()
    
    If CboNroLote.Text = "NO TIENE N LOTE" Then
        txtNroLote.Visible = True
        
        txtNroLote.SetFocus
        txtNroLote.Text = ""
        txtFechaVencimiento.Text = "__/__/____"
        CboNroLote.Visible = False
    Else
'    MsgBox CboNroLote.BoundText
        If CboNroLote.Text <> "" Then
            txtFechaVencimiento.Text = IIf(IsDate(CboNroLote.BoundText), CboNroLote.BoundText, "__/__/____")
            txtNroLote.Text = CboNroLote.Text
            'txtNroLote.Enabled = False
            CboNroLote.Visible = True
        Else
            MsgBox "Seleccione un Nro de Lote"
        End If
    End If
End Sub

Private Sub chkFraccionamiento_Click()
    Dim x As Integer
    
    If objUsuario.EsDelivery = True And objProducto.ParametroSugerido = 1 Then
        If chkFraccionamiento.Value = 1 Then
            Dim rs As oraDynaset
            Set rs = objProducto.ListaFraccionamientoSugerido(lblCodigo.Caption)
            lblFraccionamiento.Caption = "" & rs("UNID_VTA").Value
            lblCantSugerida.Caption = "" & rs("VAL_FRAC_LOCAL").Value
            'I.ECASTILLO 22.10.2020
            lblFraccionamiento.Caption = "" & IIf(Len(Trim(unidDcCap)) > 0, unidDcCap, lblFraccionamiento.Caption)
            lblCantSugerida.Caption = "" & IIf(Len(Trim(valFracDcCap)) > 0, valFracDcCap, lblCantSugerida)
            'F.ECASTILLO 22.10.2020
        Else
            lblFraccionamiento.Caption = ""
            'I.ECASTILLO 22.10.2020
            lblCantSugerida.Caption = "1"
            'F.ECASTILLO 22.10.2020
        End If
    End If
    
    'I.ECASTILLO 12.08.2021
    If gstrIndDCSAP = "1" Then
        If objVenta.isDCSAP = "1" Then
            For x = 0 To arrInfo.Count(1) - 1
                If arrInfo(x, 0) = "PACK_MODE" Then
                    lblStockRealCant.Caption = arrInfo(x, 1)
                ElseIf arrInfo(x, 0) = "PART_MODE" Then
                    txtStockRealFrac.Caption = arrInfo(x, 1)
                End If
            Next x
            If Len(Trim(lblStockRealCant.Caption)) = 0 Then txtCantidad.Enabled = False
            If Len(Trim(txtStockRealFrac.Caption)) = 0 Then txtCantidadFrac.Enabled = False
        Else
            GoTo prevDCSAP
        End If
    Else
prevDCSAP:
        If chkFraccionamiento.Value = 1 Then
            For x = 0 To arrInfo.Count(1) - 1
                If arrInfo(x, 0) <> "PACK_MODE" Then lblStockRealCant.Caption = arrInfo(x, 1)
            Next x
        Else
            For x = 0 To arrInfo.Count(1) - 1
                If arrInfo(x, 0) = "PACK_MODE" Then lblStockRealCant.Caption = arrInfo(x, 1)
            Next x
            lblFraccionamiento.Caption = "": lblCantSugerida.Caption = 1
        End If
    End If
    'F.ECASTILLO 12.08.2021
End Sub





Private Sub txtCantidadFrac_KeyPress(KeyAscii As Integer)
On Error GoTo handle
    If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = 13 Or KeyAscii = 8) Then
       KeyAscii = 0
    End If
    If KeyAscii = 13 Then
    
        If objConvenio.Lote(objVenta.CodModalidadVenta, objVenta.CodigoConvenio) = "0" Then
            Call aceptar
            frmPedido.Recalcular
        Else
            If objVenta.CodigoTipoVenta = Guias_Remision Then
                If objProducto.DevIndicadorLote(lblCodigo.Caption) = "N" Then
                    txtFechaVencimiento.SetFocus
                Else
                    If CboNroLote.Visible = True Then
                        CboNroLote.SetFocus
                    Else
                        txtNroLote.SetFocus
                    End If
                End If
            Else
                txtNroLote.SetFocus
            End If
        End If
    End If
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub txtNroLote_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And txtFechaVencimiento.Visible = False Then Call aceptar
End Sub

Private Sub chkFraccionamiento_GotFocus()
    ctlGrillaArray3.Array1 = muestraArray(objVenta.ProductoLote, lblCodigo.Caption)
End Sub

Private Sub chkFraccionamiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

'Agregado para lotes por producto MLevano 30/08/2012
Private Sub cmdAgregar_Click()
    If objVenta.CodigoTipoVenta = Guias_Remision Then
        If txtNroLote.Visible = True Then
            If txtCantidad.Text = "" Then
                MsgBox "Debe ingresar una cantidad", vbCritical
                txtCantidad.SetFocus
            Else
                If UCase(txtNroLote.Text) = UCase("Ingresar Nro de Lote") Or Trim(txtNroLote.Text) = "" Then
                    MsgBox "Debe ingresar numero de lote"
                    txtNroLote.SetFocus
                Else
                    Call txtFechaVencimiento_KeyPress(13)
                    txtCantidad.SetFocus
                    txtCantidad.Clear
                End If
            End If
        Else
            If txtCantidad.Text = "" Then
                MsgBox "Debe ingresar una cantidad"
                txtCantidad.SetFocus
            Else
                If UCase(CboNroLote.Text) = UCase("Ingresar Nro de Lote") Or Trim(txtNroLote.Text) = "" Then
                    MsgBox "Debe seleccionar numero de lote"
                    CboNroLote.SetFocus
                Else
                    Call txtFechaVencimiento_KeyPress(13)
                    txtCantidad.SetFocus
                    txtCantidad.Clear
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdEliminar_Click()
    Call ctlGrillaArray3_KeyDown(vbKeyDelete, 0)
End Sub

Function muestraArray(arr As XArrayDB, Optional cod As String = "") As XArrayDB
xTempNroLote.ReDim 0, -1, 0, 27

Dim u, g As Integer
xTempNroLote.Clear
'MsgBox arr.Count(1)
    While u <= arr.Count(1)
        If u <> arr.Count(1) Then
            If cod = arr(u, 0) Then
                xTempNroLote.AppendRows
                g = xTempNroLote.Count(1) - 1
                xTempNroLote(g, 0) = arr(u, 0)
                xTempNroLote(g, 1) = arr(u, 1)
                xTempNroLote(g, 2) = arr(u, 2)
                xTempNroLote(g, 3) = arr(u, 3)
                xTempNroLote(g, 4) = arr(u, 4)
                xTempNroLote(g, 5) = arr(u, 5)
                xTempNroLote(g, 6) = arr(u, 6)
                xTempNroLote(g, 7) = arr(u, 7)
                xTempNroLote(g, 8) = arr(u, 8)
                xTempNroLote(g, 9) = arr(u, 9)
                xTempNroLote(g, 10) = arr(u, 10)
                xTempNroLote(g, 11) = arr(u, 11)
                xTempNroLote(g, 12) = arr(u, 12)
                xTempNroLote(g, 13) = arr(u, 13)
                xTempNroLote(g, 14) = arr(u, 14)
                xTempNroLote(g, 15) = arr(u, 15)
                xTempNroLote(g, 16) = arr(u, 16)
                xTempNroLote(g, 17) = arr(u, 17)
                xTempNroLote(g, 18) = arr(u, 18)
                xTempNroLote(g, 19) = arr(u, 19)
                xTempNroLote(g, 20) = arr(u, 20)
                xTempNroLote(g, 21) = arr(u, 21)
                xTempNroLote(g, 22) = arr(u, 22)
                xTempNroLote(g, 23) = arr(u, 23)
                xTempNroLote(g, 24) = arr(u, 24)
            End If
        Else
                xTempNroLote.AppendRows
                g = xTempNroLote.Count(1) - 1
                xTempNroLote(g, 0) = ""
                xTempNroLote(g, 1) = ""
                xTempNroLote(g, 2) = ""
                xTempNroLote(g, 3) = ""
                xTempNroLote(g, 4) = ""
                xTempNroLote(g, 5) = ""
                xTempNroLote(g, 6) = ""
                xTempNroLote(g, 7) = ""
                xTempNroLote(g, 8) = ""
                xTempNroLote(g, 9) = ""
                xTempNroLote(g, 10) = ""
                xTempNroLote(g, 11) = ""
                xTempNroLote(g, 12) = ""
                xTempNroLote(g, 13) = ""
                xTempNroLote(g, 14) = ""
                xTempNroLote(g, 15) = ""
                xTempNroLote(g, 16) = ""
                xTempNroLote(g, 17) = ""
                xTempNroLote(g, 18) = ""
                xTempNroLote(g, 19) = ""
                xTempNroLote(g, 20) = ""
                xTempNroLote(g, 21) = ""
                xTempNroLote(g, 22) = ""
                xTempNroLote(g, 23) = ""
                xTempNroLote(g, 24) = ""
        End If
        u = u + 1
    Wend
    
Set muestraArray = xTempNroLote
End Function



Private Sub ctlGrillaArray3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo handle
        Select Case KeyCode

        Case vbKeyDelete
        If ctlGrillaArray3.row = ctlGrillaArray3.ApproxCount - 1 Then
            Exit Sub
        End If
                    
            If xTempNroLote.Count(1) > 0 Then
                If ctlGrillaArray3.ApproxCount = 0 Then
                    MsgBox "No hay productos seleccionados"
                    Exit Sub
                Else
                        Dim PromstrCodProducto As String
                        Dim PromstrCodTipoVenta As TipoVenta
                        Dim PromstrNroLote As String
                        
                        '----INICIO JOSE MELGAR
                        Dim bolActualizaPromo As Boolean
                        Dim indice As Integer
                        Dim i%
                        '----FIN JOSE MELGAR
                        
                        PromstrCodProducto = ctlGrillaArray3.Columns(0).Value
                        PromstrCodTipoVenta = ctlGrillaArray3.Columns(5).Value
                        PromstrNroLote = ctlGrillaArray3.Columns(22).Value
                        '----INICIO JOSE MELGAR
                        i = 0
                        For i = 0 To xTempNroLote.Count(1) - 1
                            If objVenta.ProductoLote(i, 0) = PromstrCodProducto And objVenta.ProductoLote(i, 22) = PromstrNroLote Then
                                indice = i
GoTo salta
                            End If
                        Next
salta:
                        
                        If objVenta.ProductoLote(indice, 7) <> 0 Then
                            bolActualizaPromo = False
                        Else
                            If objVenta.ProductoLote(indice, 26) = "" Then
                                bolActualizaPromo = False
                            Else
                                bolActualizaPromo = True
                            End If
                        End If
                        
                    

                   
                        
                   objVenta.EliminaProductoNroLote PromstrCodProducto, PromstrCodTipoVenta, PromstrNroLote, FracEnabled
                   
                    'frmPedido.grdPedido.Rebind
                    ' frmPedido.grdPedido.Limpiar
                    frmPedido.grdPedido.Array1 = objVenta.Producto
                    'ctlGrillaArray3.Array1 = muestraArray(objVenta.ProductoLote, lblCodigo.Caption)
                    ctlGrillaArray2.Rebind
                    ctlGrillaArray3.Array1 = muestraArray(objVenta.ProductoLote, lblCodigo.Caption)
                    'MsgBox ctlGrillaArray3.ApproxCount
                    'ctlGrillaArray3.Refresh
                    
                    objVenta.LimpiaRecetario
                    Unload frm_VTA_RecetarioM
                    '----INICIO JOSE MELGAR
                    If bolActualizaPromo Then
                        frmPedido.Cal_Promo
                    End If
                    '----FIN JOSE MELGAR
                    frmPedido.Cal_Montos
                    'frmPedido.grdPedido.Rebind
                    txtCantidad.SetFocus
                    'frmPedido.RefrescarGrilla
                    'frmPedido.grdPedido.Array1 = objVenta.Producto
                    frmPedido.grdPedido.Rebind
                End If
            End If
        End Select


'    txtCantidad.Focus
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub
'Fin de Actualizacion

Private Sub Form_Activate()
    pBlnModal = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control
    Dim tmpCtrl As Boolean, tmpAlt As Boolean
    tmpCtrl = (Shift And vbCtrlMask) > 0
    tmpAlt = (Shift And vbAltMask) > 0
    psub_KeyDownAplicacion KeyCode, Shift
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
        Case vbKeyF1
        Case vbKeyF2
        Case vbKeyF3
        Case vbKeyF4
        ''Case vbKeyReturn
        
            ''Dim objTimer As New cGetTimer
            
            ''objTimer.StartTimer
        Case tmpCtrl And vbKeyR
            frm_DLV_Reporte0.bool = False
            frm_DLV_Reporte0.stockProducto = lblStockRealCant
            frm_DLV_Reporte0.codProducto = lblCodigo.Caption
            frm_DLV_Reporte0.desProducto = lblDescripcion
            frm_DLV_Reporte0.strLaboratorio = strLaboratorio
            frm_DLV_Reporte0.CodLocal = strCodLocalPosu
            frm_DLV_Reporte0.Show vbModal
'            If frm_DLV_Reporte0.bool = True Then flgCerrar = True
'            Exit Function
    End Select

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If objUsuario.MayoristaFracciona = False And objVenta.CodigoTipoVenta = Venta_Mayorista Then Exit Sub
    If KeyAscii = 39 Then KeyAscii = 0
    If val(strFracciona) = 0 Then
         chkFraccionamiento.Enabled = False
      Else
        If UCase(Chr(KeyAscii)) = "F" And chkFraccionamiento.Visible Then _
           chkFraccionamiento.Value = IIf(chkFraccionamiento.Value = 1, 0, 1)
    End If
End Sub

Private Sub Form_Load()
    ctlGrillaArray2.Visible = False
    ctlGrillaArray3.Visible = False
    cmdAgregar.Visible = False
    cmdEliminar.Visible = False
        
        'Agregado para lotes por producto MLevano 30/08/2012
    If objVenta.CodigoTipoVenta = Guias_Remision Then
        arrCampos = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        arrCaption = Array("Código", "Descripción", "F", "Cantidad", "Precio", "T", "FlagFraccion", "Regalo", "TipDcto", "ImpDcto", "PrcAnt", "CodAutoriza", "CodUsuario", "Pct_Comi", "CtdProductoOrig", "FlgFraccionOrig", "Dato1", "Dato2", "Dato3", "Dato4", "Dato5", "FlgReceta", "NroLote", "FchVmto", "FlgFarmaco", "Pre Publico")
        arrAncho = Array(700, 1800, 400, 400, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 800, 1000, 0, 0)
        arrAlineacion = Array(dbgLeft, dbgLeft, dbgCenter, dbgRight, dbgRight, dbgCenter, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgGeneral, dbgCenter, dbgRight)

        'ctlGrillaArray2.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        ctlGrillaArray3.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        
        ctlGrillaArray3.Columns(4).Visible = False
        ctlGrillaArray3.Columns(6).Visible = False
        ctlGrillaArray3.Columns(7).Visible = False
        ctlGrillaArray3.Columns(8).Visible = False
        ctlGrillaArray3.Columns(9).Visible = False
        ctlGrillaArray3.Columns(10).Visible = False
        ctlGrillaArray3.Columns(11).Visible = False
        ctlGrillaArray3.Columns(12).Visible = False
        ctlGrillaArray3.Columns(13).Visible = False
        ctlGrillaArray3.Columns(14).Visible = False
        ctlGrillaArray3.Columns(15).Visible = False
        ctlGrillaArray3.Columns(16).Visible = False
        ctlGrillaArray3.Columns(17).Visible = False
        ctlGrillaArray3.Columns(18).Visible = False
        ctlGrillaArray3.Columns(19).Visible = False
        ctlGrillaArray3.Columns(20).Visible = False
        ctlGrillaArray3.Columns(21).Visible = False
        ctlGrillaArray3.Columns(24).Visible = False
        ctlGrillaArray3.Columns(25).Visible = False
        
        
        CboNroLote.Visible = True
        txtNroLote.Visible = False
        ctlGrillaArray3.Visible = True
        cmdAgregar.Visible = True
        cmdEliminar.Visible = True

    ElseIf objVenta.CodigoTipoVenta = Venta_Regular Then
        ctlGrillaArray2.Visible = False
        ctlGrillaArray3.Visible = False
        txtNroLote.Visible = False
        CboNroLote.Visible = False
        txtFechaVencimiento.Visible = False
        cmdAgregar.Visible = False
        cmdEliminar.Visible = False
    End If
    lblSugerido.Caption = ""
    
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pBlnModal = False
    pbytIntentos = 0
End Sub

Private Sub txtCantidad_GotFocus()
    If objVenta.CodigoTipoVenta = Guias_Remision Then
        ctlGrillaArray3.Array1 = muestraArray(objVenta.ProductoLote, lblCodigo.Caption)
        If objProducto.DevIndicadorLote(lblCodigo.Caption) = "N" Then
            CboNroLote.Visible = False
            txtNroLote.Visible = True
            txtNroLote.Text = "NOTHING"
            txtNroLote.Enabled = False
        End If
    End If
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
On Error GoTo handle
    If Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = 13 Or KeyAscii = 8) Then
       KeyAscii = 0
    End If
    If KeyAscii = 13 Then
    
    '''< 21-ABR-15  TCT Aqui Ingresa Cantidad a Comprar>
    '''< /21-ABR-15  TCT Aqui Ingresa Cantidad a Comprar>
    
        If objConvenio.Lote(objVenta.CodModalidadVenta, objVenta.CodigoConvenio) = "0" Then
            Call aceptar
            frmPedido.Recalcular
        'Modificado para el uso de nro de lote
        Else
            If objVenta.CodigoTipoVenta = Guias_Remision Then
                If objProducto.DevIndicadorLote(lblCodigo.Caption) = "N" Then
                    txtFechaVencimiento.SetFocus
                Else
                    If CboNroLote.Visible = True Then
                        CboNroLote.SetFocus
                    Else
                        txtNroLote.SetFocus
                    End If
                End If
            Else
                txtNroLote.SetFocus
            End If
        'Terminó la actualización
        End If
    End If
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

''-- Recetario Magistral --'
Private Sub sub_arma_cadena()
    ''strCantAnt = ""

    strCadCodIns = frm_VTA_Busqueda.grdProductos.Columns("COD_TIPO_INSUMO").Value
    strCadDesIns = frm_VTA_Busqueda.grdProductos.Columns("DES_TIPO_INSUMO").Value
    strCadCodigo = frm_VTA_Busqueda.grdProductos.Columns("COD_PRODUCTO").Value
    strCadProd = frm_VTA_Busqueda.grdProductos.Columns("DES_PRODUCTO").Value
    strCadUnd = frm_VTA_Busqueda.grdProductos.Columns("DES_UND_CAPACIDAD_ABREV").Value
    strCadPrecio = frm_VTA_Busqueda.grdProductos.Columns("IMP_PRECIO_VTA").Value
    strCadCodUnd = frm_VTA_Busqueda.grdProductos.Columns("COD_UND_CAPACIDAD").Value
    
    Select Case frm_VTA_Busqueda.grdProductos.Columns("COD_TIPO_INSUMO").Value
        Case "001"
            strCadPctMargen = 0
            strCtdBase = ""
            strCadCant = Trim(txtCantidad.Text)
            strCadSubTotal = strCadPrecio * strCadCant
            strCantAnt = strCadCant
        Case "002", "003", "004"
            If strCantAnt <> "" Then
                strCtdBase = strCantAnt
                strCadCant = Trim(txtCantidad.Text)
                strCadPctMargen = (strCadCant / strCtdBase) * 100
                strCadSubTotal = strCadPrecio * strCadCant
            Else
                ''Esta linea de validacion se agrego el dia 21/03/2007'
                If frm_VTA_RecetarioM.GrdInsumos.ApproxCount <= 0 Then strCtdBase = ""
                
                strCadPctMargen = "" ''strCadSubTotal = ""
                strCadCant = Trim(txtCantidad.Text)
                strCadSubTotal = strCadPrecio * strCadCant
            End If
        Case Else
            If frm_VTA_RecetarioM.GrdInsumos.ApproxCount <= 0 Then strCtdBase = ""
            strCadPctMargen = "" '' strCadSubTotal = ""
            strCadCant = Trim(txtCantidad.Text)
            strCadSubTotal = strCadPrecio * strCadCant
    End Select
    
''    If frm_VTA_Busqueda.grdProductos.Columns("COD_TIPO_INSUMO").Value = "001" Then
''        strCadPctMargen = 0
''        strCtdBase = ""
''        'strCantAnt = ""
''        strCadCant = Trim(txtCantidad.Text)
''        strCadSubTotal = strCadPrecio * strCadCant
''        strCantAnt = strCadCant
''    Else
''        If strCantAnt <> "" Then
''            strCtdBase = strCantAnt
''            strCadCant = Trim(txtCantidad.Text)
''            strCadPctMargen = (strCadCant / strCtdBase) * 100
''            strCadSubTotal = strCadPrecio * strCadCant
''          Else
''            'Esta linea de validacion se agrego el dia 21/03/2007'
''            If frm_VTA_RecetarioM.GrdInsumos.ApproxCount <= 0 Then strCtdBase = ""
''            strCadPctMargen = "": strCadSubTotal = ""
''            strCadCant = Trim(txtCantidad.Text)
''        End If
''    End If
    
End Sub


''Private Sub cmdProbar_Click()
''Dim i As Integer
''Dim j As Integer
''
''
''                'If grdPedido.ApproxCount = 0 Then Exit Sub
''
''    j = 0
''    For i = 0 To objVenta.Producto.UpperBound(1)
''        If objVenta.Producto(i, 7) = "0" Then
''                j = j + 1
''        End If
''    Next i
''
''    If j = 0 Then Exit Sub
''
''    If objUsuario.PrecioOnLine = 1 Then Cal_Promo
''
''End Sub

Function buscaExcluido(codProducto As String) As String
Dim ODyn As oraDynaset
Set ODyn = objProducto.ListaExcluyeMayorista(codProducto)
ODyn.MoveFirst
While Not ODyn.EOF
    If ODyn("COD_PRODUCTO") = codProducto Then
        buscaExcluido = ODyn("DES_MOTIVO")
        Exit Function
    Else
        buscaExcluido = ""
    End If
Wend
buscaExcluido = ""
End Function


Private Sub aceptar()



If objVenta.CodigoTipoVenta = Servicio Then
    MsgBox "No se puede realizar cobranza de Servicios con productos", vbCritical, App.ProductName
    Exit Sub
End If


'condiciones de venta mayorista
'If objVenta.CodigoTipoVenta = Venta_Mayorista Then
'    Dim excluido As String
'    excluido = buscaExcluido(lblCodigo.Caption)
'    If excluido <> "" Then
'        MsgBox "El producto fue excluido para venta Mayorista", vbCritical, App.ProductName
'        Unload Me
'        Exit Sub
'    End If
'End If
'fin de condiciones de venta mayorista

If objVenta.NumMaximoUnidades > 0 Then
    If chkFraccionamiento.Value = 0 And val(txtCantidad.Text) > objVenta.NumMaximoUnidades Then
        MsgBox "El convenio no permite seleccionar más de " & objVenta.NumMaximoUnidades & " unidades por producto", vbCritical, App.ProductName
        Exit Sub
    ElseIf chkFraccionamiento.Value = 1 Then
        Dim xCtdFrac As Integer
        Dim xCtdUnidades As Double
        xCtdFrac = objProducto.intCtdFrac(lblCodigo.Caption)
        xCtdUnidades = val(txtCantidad.Text) / xCtdFrac
                If val(xCtdUnidades) > objVenta.NumMaximoUnidades Then
                      MsgBox "El convenio no permite seleccionar más de " & objVenta.NumMaximoUnidades & " unidades por producto", vbCritical, App.ProductName
                      Exit Sub
                End If
    End If
          
End If
On Error GoTo handle
           If frm_VTA_RecetarioM.pstrFlgRM = "1" Then
                ''Recetario Magistral'
                sub_arma_cadena
                Unload Me
                
                ''-- Pasando el Codigo Unico RM --'
                If frm_VTA_RecetarioM.pstrRucProv = "" Then MsgBox "Seleccione Proveedor", vbCritical, App.ProductName: frm_VTA_RecetarioM.Show: Exit Sub
                
                strCod = objProducto.ListaDevRM(objUsuario.CodigoLocal, frm_VTA_RecetarioM.pstrRucProv)
                If val(strCod) = 0 Then MsgBox "El local no esta permitido de hacer" & Chr(13) & " recetario magistral con el proveedor", vbCritical, Caption: Exit Sub
                strDes = objProducto.ListaDescripcion(strCod)
                
                
                
                Call frm_VTA_RecetarioM.psub_Agrega_Insumo(strCadCodIns, strCadDesIns, _
                                                           strCadCodigo, strCadProd, _
                                                           strCadUnd, strCadPctMargen, _
                                                           strCtdBase, strCadCant, _
                                                           strCadPrecio, strCadSubTotal, _
                                                           strCod, strCadCodUnd)
                frm_VTA_RecetarioM.GrdInsumos.Rebind
                frm_VTA_RecetarioM.Hide
                frm_VTA_RecetarioM.Show
                
             Else
                
                If objVenta.ptmModalidad <> Recetario Then
                    If objProducto.FnEsModalidad_Recetario(Trim(lblCodigo.Caption)) > 0 Then
                        MsgBox "Cambie a Modadidad Recetario Magistral para hacer la Venta", vbCritical, Caption: Exit Sub
                    End If
                End If
                
                If flgEspecieValorada = "1" And Not objVenta.ptmModalidad = Guias_Remision Then
                    Dim f As Integer
                    
                    objVenta.EspeciesValoradas.ReDim 0, -1, 0, 5
                    
                    Do While f < val(txtCantidad.Text)
                        frm_VTA_EspeciesValoradas.strCodigoProducto = lblCodigo.Caption
                        frm_VTA_EspeciesValoradas.bolCancel = False
                        ''objVenta.ptmModalidad = Guias_Remision
                        frm_VTA_EspeciesValoradas.Show vbModal
                        If frm_VTA_EspeciesValoradas.bolCancel Then
                            Exit Do
                        End If
                        f = f + 1
                    Loop
                    If frm_VTA_EspeciesValoradas.bolCancel Then
                        txtCantidad.SetFocus
                        Exit Sub
                    End If
                End If
                
                'Val(txtCantidad.Text) [ORIGINAL] -> devuelve siempre 0 por la funcion Val, no deberia dejar pasar ""
                If txtCantidad.Text = "0" And cRegalo = Producto_Regalo_precio Then
                ElseIf txtCantidad.Text = "0" And cRegalo = Producto_Normal Then
                Else
                    If objVenta.isDCSAP = "1" Then
                        If (txtCantidad.Text = "" Or val(txtCantidad.Text) = 0) And _
                            (txtCantidadFrac.Text = "" Or val(txtCantidadFrac.Text) = 0) Then Exit Sub
                    Else
                        If txtCantidad.Text = "" Or val(txtCantidad.Text) = 0 Then Exit Sub
                    End If
                End If
                Dim strLocal As String
                ''strLocal = objUsuario.CodigoLocal
                
                If flgStockTotal = True Then
                    strLocal = IIf(IsNull(frm_DLV_Stock_Total.ctlGrillaArray2.Columns(0).Value), mdiPrincipal.ctlCliente1.LocalAsignado, frm_DLV_Stock_Total.ctlGrillaArray2.Columns(0).Value)
                    ''frm_DLV_Stock_Total.Cerrar
                    ''frm_DLV_Stock_Total.Visible = True
                                    frm_DLV_Stock_Total.ctlGrillaArray2.Limpiar
                            frm_DLV_Stock_Total.ctlGrillaArray2.Rebind

                End If
                If strLocal = "" And objUsuario.EsDelivery = True Then
                    strLocal = mdiPrincipal.ctlCliente1.LocalAsignado
                End If
                
                If strLocal = "" And objVenta.ptmModalidad = Guias_Remision Then
                    strLocal = objUsuario.CodigoLocal
                End If
                
                'AGREGADO POR MLEVANO
                Dim MayorExistente As String
                MayorExistente = "0"
                objVenta.evaluaStockGuiaLote = 0
                'FIN
                        
''                If objVenta.ptmModalidad <> Guias_Remision And objUsuario.TipoMaquina <> objUsuario.TipoMaquinaCabina Then
                If objVenta.ptmModalidad <> Guias_Remision Then
                    If objUsuario.TipoMaquina <> objUsuario.TipoMaquinaCabina Then
                    
                        strLocal = objUsuario.CodigoLocal
                        If objProducto.EvaluaCtdVta(objUsuario.CodigoEmpresa, strLocal, lblCodigo.Caption, txtCantidad.Text, chkFraccionamiento.Value) = "0" Then
                            ''agregado por PHERRERA, solo debe grabar vta fallida cuando digite una vez
                            If pbytIntentos < 1 Then
                                objVenta.GrabarVentaFallida objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, txtCantidad.Text, strStock, chkFraccionamiento.Value, objUsuario.Codigo, Format(objVenta.ptmModalidad, "000"), lblCodigo.Caption
                            End If
                            pbytIntentos = pbytIntentos + 1
                            vstrStock = objProducto.Stock(lblCodigo.Caption, strLocal, "1")
                            MsgBox "La cantidad solicitada es mayor a la existente" & Chr(13) & "Stock Actual: " & vstrStock, vbExclamation + vbOKOnly, App.ProductName
                            
                            'AGREGADO POR MLEVANO
                            MayorExistente = "1"
                            'FIN
                            
                            
                            If objVenta.ptmModalidad <> Guias_Remision Or objVenta.ptmModalidad = Guias_Remision Then Exit Sub
                        End If
                    End If
                End If
                
                '''''  cambiar este valor por uno que devuelva el CMR.PKG_PRODUCTO.FN_LISTA
                Dim PctComi As Double
                PctComi = objProducto.pctComision(lblCodigo.Caption, objUsuario.CodigoLocal, Format(ptmTipoPrecio, "000"))
                ''''' *********************
                If objUsuario.CodLocalCallCenter = "1DLV" Then 'ECASTILLO 22.06.2020
                    Set oraDato = objProducto.ListaDato("94", strLocal, strTipoPrecio, lblCodigo.Caption, txtCantidad.Text, chkFraccionamiento, "", objUsuario.CodLocalCallCenter)
                Else
                    Set oraDato = objProducto.ListaDato(objUsuario.CodigoEmpresa, strLocal, strTipoPrecio, lblCodigo.Caption, txtCantidad.Text, chkFraccionamiento, "", objUsuario.CodLocalCallCenter)
                End If
                
                
                ''' b jct, se cambio empresa, 30Mar12
                '''Set oraDato = objProducto.ListaDato("99", strLocal, strTipoPrecio, lblCodigo.Caption, txtCantidad.Text, chkFraccionamiento, "")
                ''' e jct
                
                frmPedido.grdPedido.Limpiar
                
                '** Guarda el valor del precio publico en una propiedad publica para despue pasarlo a la grilla de pedido /-/ 30/12/2008 **'
                'objVenta.PrecPublic = Format(oraDato(5).Value, "###,##0.00")
                '****************************************'
                
                'Para lotes por producto MLEVANO 27/08/2012
                If (objVenta.CodigoTipoVenta = Guias_Remision Or objVenta.CodigoTipoVenta = Venta_Mayorista) And MayorExistente = "0" Then

                    If chkFraccionamiento.Enabled = True Then
                        FracEnabled = "1"
                    Else
                        FracEnabled = "0"
                    End If
                    
                    ctlGrillaArray2.Limpiar
                    
                    ctlGrillaArray2.Array1 = objVenta.AgregaProductoLote(lblCodigo.Caption, lblDescripcion.Caption, val(txtCantidad.Text), chkFraccionamiento, oraDato(4).Value, objVenta.CodigoTipoVenta, cRegalo, , , , , , strFlgReceta, PctComi, txtNroLote.Text, txtFechaVencimiento.Text, , oraDato(5), , , FracEnabled)
                    
                    'Call objVenta.AgregaProductoLoteDetalle(objVenta.ProductoLote, lblCodigo.Caption, FracEnabled)
                    
                    'Array que suma y muestra por fracciones
                    'ctlGrillaArray3.Array1 = objVenta.AgregaProductoLoteDetalle(lblCodigo.Caption, lblDescripcion.Caption, Val(txtCantidad.Text), chkFraccionamiento, oraDato(4).Value, objVenta.CodigoTipoVenta, cRegalo, , , , , , , , txtNroLote.Text, txtFechaVencimiento.Text)
                    'Array que muestra segun el codigo
                    
                    'a comentar
                    'ctlGrillaArray2.Array1 = muestraArray(objVenta.ProductoLote, lblCodigo.Caption)
                    
                    ctlGrillaArray3.Array1 = muestraArray(objVenta.ProductoLote, lblCodigo.Caption)
                    
                    '(lblCodigo.Caption, lblDescripcion.Caption, Val(txtCantidad.Text), chkFraccionamiento, oraDato(4).Value, objVenta.CodigoTipoVenta, cRegalo, , , , , , , , txtNroLote.Text, txtFechaVencimiento.Text)
                End If
                
                'if agregado para validar la evaluacion del productolote 15/10/2012 MLEVANO
                If objVenta.evaluaStockGuiaLote = 0 Then
                    
                    'EMPIEZA Modificacion para la Implementacion codigo de Barras y claves del QF para vender estupefacientes
                    'MLevano 09/11/2012
                If objVenta.CodigoTipoVenta <> Guias_Remision Then
                    Set rs = objProducto.ValidacionesPorProducto(lblCodigo.Caption)
'VALIDAR rs("FLG_VAL_LOCAL")
                    If Not (rs(1) = "0" And rs(2) = "0" And rs(3) = "0" And rs(4) = "0" And rs(5) = "0") Then       'Or rs("1") = "0" Then
                        Call frm_VTA_ValidaCantidadProducto.subDatos(lblCodigo.Caption, lblDescripcion.Caption, rs(1), rs(2), rs(3), rs(4), rs(5))
                    End If
                End If
                    'MsgBox "Continua?"
                        
                        'EMPIEZA ORIGINAL
                        Dim STRtxtCantidad As Integer
                        Dim STRtxtCantidadFrac As Integer
                        If objUsuario.EsDelivery = True And objProducto.ParametroSugerido = 1 Then
                            STRtxtCantidad = val(txtCantidad.Text) * val(lblCantSugerida.Caption)
                            Else
                            STRtxtCantidad = val(txtCantidad.Text)
                        End If
                        'I.ECASTILLO 12.08.2021
                        'I.CVIERA 31.12.2020 | 05.01.2021.REVISAR | 06.01.2021
                        Dim flg_ruteoA_cnv
                        flg_ruteoA_cnv = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRACNV") '1 => ACTIVO, 0 => INACTIVO
                        If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
                            GoTo cnvNoRuteaAuto
                        End If
                        Dim flg_2e_reserva
                        flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
                        If flg_2e_reserva = "0" Then
cnvNoRuteaAuto:
                            frmPedido.grdPedido.Array1 = objVenta.AgregaProducto(lblCodigo.Caption, lblDescripcion.Caption, STRtxtCantidad, chkFraccionamiento, oraDato(4).Value, objVenta.CodigoTipoVenta, cRegalo, , , , , , strFlgReceta, PctComi, txtNroLote.Text, txtFechaVencimiento.Text, , oraDato(5), , , FracEnabled)
                        Else
                            Dim Buscar As String
                            Dim a As Integer
                            'I.ECASTILLO 17.12.2020
                            Dim valPrecioNuevo As Double
                            Dim flg_Transf As String
                            Dim STRtxtCantidadRealCant As Long
                            If objVenta.isDCSAP = "1" Or objVenta.isLocalDcCappa = "1" Then
                                STRtxtCantidadRealCant = val(lblStockRealCant) * val(lblCantSugerida.Caption)
                            Else
                                STRtxtCantidadRealCant = val(lblStockRealCant)
                            End If
                            'If val(STRtxtCantidad) > val(lblStockRealCant) Then flg_Transf = "1"
                            If val(STRtxtCantidad) > val(STRtxtCantidadRealCant) Then flg_Transf = "1"
                            'F.ECASTILLO 17.12.2020
                            a = 0
                            Dim valChkFraccionamiento As String
                            valChkFraccionamiento = chkFraccionamiento
                            If gstrIndDCSAP = "1" Then
                                If objVenta.isDCSAP = "1" Then
                                    STRtxtCantidadFrac = val(txtCantidadFrac.Text) * val(valFracDcCap)
                                    STRtxtCantidad = STRtxtCantidad '* val(valFracMaxDc)
                                    'STRtxtCantidad = STRtxtCantidad + STRtxtCantidadFrac
                                    If STRtxtCantidad > 0 Then frmPedido.grdPedido.Array1 = objVenta.AgregaProducto(lblCodigo.Caption, lblDescripcion.Caption, STRtxtCantidad, valChkFraccionamiento, oraDato(4).Value, objVenta.CodigoTipoVenta, cRegalo, , , , , , strFlgReceta, PctComi, txtNroLote.Text, txtFechaVencimiento.Text, , oraDato(5), , , FracEnabled, , , , flg_Transf, 1)
                                    If STRtxtCantidadFrac > 0 Then valChkFraccionamiento = 1: frmPedido.grdPedido.Array1 = objVenta.AgregaProducto(lblCodigo.Caption, lblDescripcion.Caption, STRtxtCantidadFrac, valChkFraccionamiento, oraDato(4).Value, objVenta.CodigoTipoVenta, cRegalo, , , , , , strFlgReceta, PctComi, txtNroLote.Text, txtFechaVencimiento.Text, , oraDato(5), , , FracEnabled, , , , flg_Transf, 1)
                                Else
                                    GoTo agregaProductoRegular
                                End If
                            Else
agregaProductoRegular:
                                frmPedido.grdPedido.Array1 = objVenta.AgregaProducto(lblCodigo.Caption, lblDescripcion.Caption, STRtxtCantidad, valChkFraccionamiento, oraDato(4).Value, objVenta.CodigoTipoVenta, cRegalo, , , , , , strFlgReceta, PctComi, txtNroLote.Text, txtFechaVencimiento.Text, , oraDato(5), , , FracEnabled, , , , flg_Transf)
                            End If
                            

'                            For a = 0 To frm_VTA_Busqueda.grdProductos.ApproxCount - 1
'                                Buscar = frm_VTA_Busqueda.grdProductos.Columns(0).Value
'                                If Buscar = lblCodigo.Caption Then
'                                    If frm_VTA_Busqueda.grdProductos.Columns("FLG_SEG").Value = "1" Then
'                                        valPrecioNuevo = IIf(chkFraccionamiento.Value = "0", STRtxtCantidad * frm_VTA_Busqueda.grdProductos.Columns("PRECIO").Value, (frm_VTA_Busqueda.grdProductos.Columns("PRECIO").Value / oraDato(3)) * STRtxtCantidad) 'ECASTILLO 05.01.2021
'                                        frmPedido.grdPedido.Array1 = objVenta.AgregaProducto(lblCodigo.Caption, lblDescripcion.Caption, STRtxtCantidad, chkFraccionamiento, valPrecioNuevo, objVenta.CodigoTipoVenta, cRegalo, , , , , , strFlgReceta, PctComi, txtNroLote.Text, txtFechaVencimiento.Text, , valPrecioNuevo, , , FracEnabled, , , frm_VTA_Busqueda.grdProductos.Columns("FLG_SEG").Value, flg_Transf) 'ECASTILLO 17.12.2020
'                                    Else
'                                        frmPedido.grdPedido.Array1 = objVenta.AgregaProducto(lblCodigo.Caption, lblDescripcion.Caption, STRtxtCantidad, chkFraccionamiento, oraDato(4).Value, objVenta.CodigoTipoVenta, cRegalo, , , , , , strFlgReceta, PctComi, txtNroLote.Text, txtFechaVencimiento.Text, , oraDato(5), , , FracEnabled, , , frm_VTA_Busqueda.grdProductos.Columns("FLG_SEG").Value, flg_Transf)
'                                        'frmPedido.grdPedido.Array1 = objVenta.AgregaProducto(lblCodigo.Caption, lblDescripcion.Caption, STRtxtCantidad, chkFraccionamiento, CStr(frm_VTA_MetodosSegmentos.strPrecioTipo), objVenta.CodigoTipoVenta, cRegalo, , , , , , strFlgReceta, PctComi, txtNroLote.Text, txtFechaVencimiento.Text, , CStr(frm_VTA_MetodosSegmentos.strPrecioTipo), , , FracEnabled, , , 1)
'                                    End If
'                                    Exit For
'                                Else
'                                    frmPedido.grdPedido.Array1 = objVenta.AgregaProducto(lblCodigo.Caption, lblDescripcion.Caption, STRtxtCantidad, chkFraccionamiento, oraDato(4).Value, objVenta.CodigoTipoVenta, cRegalo, , , , , , strFlgReceta, PctComi, txtNroLote.Text, txtFechaVencimiento.Text, , oraDato(5), , , FracEnabled, , , , flg_Transf)
'                                    Exit For
'                                End If
'                                'frm_VTA_Busqueda.grdProductos.MoveNext
'                            Next a
                        End If
                        'F.CVIERA 31.12.2020
                        
                        '''< 21-ABR-15  TCT  Devuelve Prods  para Grilla >
                        '''< /21-ABR-15  TCT  Devuelve Prods  para Grilla >
                        
                        'MsgBox "Datos Cargados..." + CStr(frmPedido.grdPedido.Array1(0, 26))
                        
                        ctlGrillaArray1.Limpiar
                        'TERMINA ORIGINAL
                    'TERMINA
                
                End If
                
                ''' ***** 24/03/2007
                ''*** Solo que calcule para cualquier promocion pero menos para devolucion de SOAT
                ''*** Cambio hecho 03/07/2007 por crueda
                If Not flgEspecieValorada = "1" And Not objVenta.ptmModalidad = Guias_Remision Then
                    frmPedido.Cal_Promo
                End If
                ''' *****
                
                frmPedido.Cal_Montos
                frmPedido.grdPedido.Rebind
           End If
           
           '=======================================================================Para lotes x producto MLEVANO
           If objVenta.CodigoTipoVenta <> Guias_Remision Then
                Unload Me
           End If
           '======================================================================================
           Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub


Private Sub txtFechaVencimiento_KeyPress(KeyAscii As Integer)
Dim dfecha As Date
On Error GoTo CtrlErr

    dfecha = ("1/" & month(objUsuario.sysdate) & "/" & year(objUsuario.sysdate))
    If KeyAscii = 13 Then
        If objProducto.DevIndicadorLote(lblCodigo.Caption) = "N" Then
            Call aceptar
        Else
            
                If txtFechaVencimiento.Text = "__/__/____" Then
                    If CboNroLote.Visible = True Then
                        If Trim(CboNroLote.Text) <> "" Then
                            Call aceptar
                        Else
                            MsgBox "Debe seleccionar un numero de lote", vbCritical, App.ProductName
                        End If
                    Else
                        If Trim(txtNroLote.Text) <> "" Then
                            Call aceptar
                        Else
                            MsgBox "Debe ingresar un numero de lote", vbCritical, App.ProductName
                        End If
                    End If
                Else
                    If Not IsDate(txtFechaVencimiento.Text) Or CDate(txtFechaVencimiento.Text) < dfecha Then
                        MsgBox "La fecha de vencimiento es incorrecta", vbCritical + vbOKOnly, App.ProductName
                        txtFechaVencimiento.SetFocus
                        KeyAscii = 0
                    Else
                        If txtCantidad.Text = "" Then
                            MsgBox "Debe ingresar una cantidad", vbCritical, App.ProductName
                            txtCantidad.SetFocus
                        Else
                            Call aceptar
                        End If
                    End If
                End If
            
            
                
        End If
    End If

    
Exit Sub

CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName


End Sub

Private Sub txtNroLote_GotFocus()
    Me.KeyPreview = False

End Sub

Private Sub txtNroLote_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    
End Sub

Private Sub txtNroLote_LostFocus()
    Me.KeyPreview = True
End Sub
