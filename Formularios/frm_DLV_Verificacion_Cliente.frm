VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_DLV_Verificacion_Cliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificación de Cliente Nuevo"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   Icon            =   "frm_DLV_Verificacion_Cliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAnular 
      Caption         =   "&Anular"
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Detalle"
      Height          =   375
      Left            =   2985
      TabIndex        =   28
      Top             =   8760
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   14208
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos del Cliente [F1]"
      TabPicture(0)   =   "frm_DLV_Verificacion_Cliente.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ctlCliente1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos del Pedido [F2]"
      TabPicture(1)   =   "frm_DLV_Verificacion_Cliente.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Datos de Entrega"
         Height          =   3975
         Left            =   840
         TabIndex        =   6
         Top             =   720
         Width           =   7215
         Begin VB.CheckBox ChkFlgEntLocal 
            Caption         =   "Entrega en Local"
            Height          =   255
            Left            =   4800
            TabIndex        =   8
            Top             =   3480
            Width           =   1575
         End
         Begin VB.CheckBox ChkFlgEntTercero 
            Caption         =   "Entrega a Tercero"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker dtpFechaPed 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd/mm/yyyy hh:mm:ss AMPM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   375
            Left            =   1320
            TabIndex        =   9
            Top             =   3480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   65339393
            CurrentDate     =   39044
         End
         Begin vbp_Ventas.ctlTextBox txtNombreEntrega 
            Height          =   375
            Left            =   960
            TabIndex        =   10
            Top             =   600
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   661
            ColorDefault    =   -2147483639
            ColorDefault    =   -2147483639
            Enabled         =   0   'False
            Bloqueado       =   -1  'True
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
         Begin vbp_Ventas.ctlTextBox txtDirecccionEntrega 
            Height          =   375
            Left            =   960
            TabIndex        =   11
            Top             =   960
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   661
            ColorDefault    =   -2147483639
            ColorDefault    =   -2147483639
            Enabled         =   0   'False
            Bloqueado       =   -1  'True
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
         Begin vbp_Ventas.ctlTextBox TxtReferencia 
            Height          =   255
            Left            =   1320
            TabIndex        =   12
            Top             =   2760
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   450
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
         Begin vbp_Ventas.ctlTextBox TxtFono 
            Height          =   255
            Left            =   1320
            TabIndex        =   13
            Top             =   3120
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   450
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
         Begin MSComCtl2.DTPicker dtpHora 
            Height          =   375
            Left            =   3000
            TabIndex        =   14
            Top             =   3480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   65339394
            CurrentDate     =   39044
         End
         Begin vbp_Ventas.ctlDataCombo cboPais 
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin vbp_Ventas.ctlDataCombo cboProvincia 
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   2280
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin vbp_Ventas.ctlDataCombo cboDepartamento 
            Height          =   315
            Left            =   2880
            TabIndex        =   17
            Top             =   1680
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin vbp_Ventas.ctlDataCombo cboDistrito 
            Height          =   315
            Left            =   2880
            TabIndex        =   18
            Top             =   2280
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Pais"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1440
            Width           =   300
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Provincía"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   2040
            Width           =   690
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Left            =   2880
            TabIndex        =   25
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Distrito"
            Height          =   195
            Left            =   2880
            TabIndex        =   24
            Top             =   2040
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha "
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   3600
            Width           =   495
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefóno"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   3120
            Width           =   630
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   2760
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   555
         End
      End
      Begin vbp_Ventas.ctlCliente ctlCliente1 
         Height          =   7455
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   13150
      End
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5595
      TabIndex        =   1
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4275
      TabIndex        =   0
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Label LblTelefono 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   450
      Left            =   5520
      TabIndex        =   30
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Telefono"
      Height          =   195
      Left            =   4560
      TabIndex        =   29
      Top             =   240
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nº Pedido"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   720
   End
   Begin VB.Label LblPedido 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   450
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frm_DLV_Verificacion_Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objPedido As New clsProforma
Public strCodigoCliente As String
Public strNumeroPedido As String
Public strTelefono As String
Public pBlnCliente As Boolean
Dim rsPedido As oraDynaset
Dim strCodDireccionCli As String

Private Sub cmdAceptar_Click()
   If ctlCliente1.Grabar = "" Then
   Dim objProforma As New clsCliente
   Dim strMensaje As String
   
   '-- Al objeto proforma se agrego parametro que es Numero de pedido
   '-- Hecho 04/07/2007 por CRUEDA
   strMensaje = objProforma.GrabaVerificacion(ctlCliente1.Codigo, _
                                              ctlCliente1.Verificacion, _
                                              objUsuario.CodigoEmpresa, _
                                              objUsuario.CodigoLocal, _
                                              objUsuario.Codigo, _
                                              LblPedido.Caption)
   If strMensaje <> "" Then
    'MsgBox "Se grabo satisfactoriamente ", vbExclamation, App.ProductName
   'Else
    MsgBox strMensaje, vbCritical, App.ProductName
   End If
   
   Dim objProforma1 As New clsProforma
   Dim strMensaje2 As String
   strMensaje2 = objProforma1.ActualizaCliente(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, LblPedido.Caption)
   If Not strMensaje2 = "" Then
    MsgBox strMensaje2, vbCritical, App.ProductName
   End If
   Set objProforma1 = Nothing
   objVenta.CodigoCliente = ctlCliente1.Codigo
   objVenta.FchHoraPactEntr = "" & dtpFechaPed.Value
   objVenta.FlgEntregaLocal = "" & ChkFlgEntLocal.Value
   
   Unload Me
   
''''''    DTPicker1.Value = Format(.FchHoraPactEntr, "dd/mm/yyyy")
''''''    DTPicker2.Value = Format(.FchHoraPactEntr, "hh:mm:ss AMPM")
   
   
   End If
End Sub

Private Sub CmdAnular_Click()
    If MsgBox("Se procederá a Anular la Proforma de la pantalla de verificación", vbInformation + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
        objPedido.Anula objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strNumeroPedido, objUsuario.Codigo
        MsgBox "Se anuló satisfactoriamente", vbExclamation, App.ProductName
        Unload Me
    End If
End Sub

Private Sub Command1_Click()
    frm_VTA_DetallePedido.NumeroPedido = LblPedido
    frm_VTA_DetallePedido.ReCargaDetPedido
    frm_VTA_DetallePedido.Show vbModal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF1
        SSTab1.Tab = 0
    Case vbKeyF2
       SSTab1.Tab = 1

End Select
End Sub

Private Sub Form_Load()
      ctlCliente1.Cargar
      Carga_Datos_Pedido
      ctlCliente1.Verificar
      ctlCliente1.CodDireccionCli = strCodDireccionCli
      ctlCliente1.ConsultaCliente strCodigoCliente
      If pBlnCliente = True Then
            LblPedido.Caption = strNumeroPedido
            LblTelefono.Caption = strTelefono
      End If
      dtpFechaPed.Value = Null
      dtpHora.Value = Null
      SSTab1.Tab = 0
End Sub

Sub Carga_Datos_Pedido()
    Dim strUbigeo As String
    
    Set cboPais.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_PAIS", 0)
    cboPais.ListField = "Descripcion"
    cboPais.BoundColumn = "Codigo"
    cboPais.BoundText = "00"
    
    'Set cboDepartamento.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DEPARTAMENTO", 0)
    'cboDepartamento.ListField = "Descripcion"
    'cboDepartamento.BoundColumn = "Codigo"
    'cboDepartamento.BoundText = Mid(objUsuario.UbigeoLocal, 1, 2)
    
    'strMaxTelefono = Val(gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_NUM_DIG_TELEFONO"))
    'strMaxAnexo = Val(gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_NUM_DIG_ANEXO"))
    
    'Set cboProvincia.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_PROVINCIA", 0, cboDepartamento.BoundText)
    'cboProvincia.ListField = "Descripcion"
    'cboProvincia.BoundColumn = "Codigo"
    'cboProvincia.BoundText = Mid(objUsuario.UbigeoLocal, 3, 2)

    'Set cboDistrito.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DISTRITO", 0, cboDepartamento.BoundText, cboProvincia.BoundText)
    'cboDistrito.ListField = "Descripcion"
    'cboDistrito.BoundColumn = "Codigo"
    'cboDistrito.BoundText = Mid(objUsuario.UbigeoLocal, 5, 2)
    
    Set rsPedido = objPedido.ListaCabecera(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, strNumeroPedido)
    strCodDireccionCli = "" & rsPedido("COD_DIRECCION_CLI").Value
    txtNombreEntrega.Text = "" & rsPedido("DES_CLIENTE").Value    '"" & rsPedido("DES_AUX_CLI_NOMBRE").Value
    txtDirecccionEntrega.Text = "" & rsPedido("DES_AUX_CLI_DIRECC").Value
    ChkFlgEntLocal.Value = IIf(IsNull(rsPedido("FLG_ENTREGA_LOCAL").Value), "0", rsPedido("FLG_ENTREGA_LOCAL").Value) 'IIf(rsPedido("FLG_ENTREGA_LOCAL").Value = "", "", rsPedido("FLG_ENTREGA_LOCAL").Value)
    TxtReferencia.Text = "" & rsPedido("DES_REFERENCIA").Value
    TxtFono.Text = "" & rsPedido("DES_AUX_CLI_TLF").Value
    dtpFechaPed.Value = "" & rsPedido("FCH_REGISTRA").Value
    dtpHora.Value = "" & rsPedido("HORA").Value
    
    strUbigeo = IIf(IsNull(rsPedido("COD_UBIGEO").Value), objUsuario.UbigeoLocal, rsPedido("COD_UBIGEO").Value)
    
    Set cboDepartamento.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DEPARTAMENTO", 0)
    cboDepartamento.ListField = "Descripcion"
    cboDepartamento.BoundColumn = "Codigo"
    cboDepartamento.BoundText = Mid(strUbigeo, 1, 2)
    
    Set cboProvincia.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_PROVINCIA", 0, cboDepartamento.BoundText)
    cboProvincia.ListField = "Descripcion"
    cboProvincia.BoundColumn = "Codigo"
    cboProvincia.BoundText = Mid(strUbigeo, 3, 2)
    
    Set cboDistrito.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DISTRITO", 0, cboDepartamento.BoundText, cboProvincia.BoundText)
    cboDistrito.ListField = "Descripcion"
    cboDistrito.BoundColumn = "Codigo"
    cboDistrito.BoundText = Mid(strUbigeo, 5, 2)
    
    Set objPedido = Nothing
    
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

