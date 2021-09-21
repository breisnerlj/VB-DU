VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_cat_Busquedacodigo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de Producto"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   180
   ClientWidth     =   6930
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_cat_Busquedacodigo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   6930
   Begin VB.OptionButton Option1 
      Caption         =   "Todos"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   32
      Top             =   3360
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "No Fracciona"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   31
      Top             =   3120
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Fracciona"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   30
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Frame fraStock 
      Caption         =   "Stock"
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   120
      TabIndex        =   28
      Top             =   2880
      Width           =   2295
      Begin VB.CheckBox chkConStock 
         Caption         =   "Tienen Stock"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      Picture         =   "frm_cat_Busquedacodigo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2880
      Width           =   1095
   End
   Begin vbp_Ventas.ctlGrilla GrdProductos 
      Height          =   2775
      Left            =   0
      TabIndex        =   24
      Top             =   3720
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4895
   End
   Begin MSComctlLib.ImageList IlsImagen 
      Left            =   2520
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_cat_Busquedacodigo.frx":0894
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_cat_Busquedacodigo.frx":0E2E
            Key             =   "buscar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_cat_Busquedacodigo.frx":13C8
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_cat_Busquedacodigo.frx":1962
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_cat_Busquedacodigo.frx":1EFC
            Key             =   "Chek"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_cat_Busquedacodigo.frx":2496
            Key             =   "Bien"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_cat_Busquedacodigo.frx":2A30
            Key             =   "Agregar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_cat_Busquedacodigo.frx":2FCA
            Key             =   "Atras"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_cat_Busquedacodigo.frx":3564
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_cat_Busquedacodigo.frx":3AFE
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_cat_Busquedacodigo.frx":4098
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_cat_Busquedacodigo.frx":4632
            Key             =   "Hora"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraFechaCreacion 
      Caption         =   "Fecha Creación"
      ForeColor       =   &H00800000&
      Height          =   1065
      Left            =   0
      TabIndex        =   17
      Top             =   720
      Width           =   3375
      Begin MSComCtl2.DTPicker dtpFchIni 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "MMMM"
         DateIsNull      =   -1  'True
         Format          =   61079553
         CurrentDate     =   38219
      End
      Begin MSComCtl2.DTPicker dtpFchFin 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   61079553
         CurrentDate     =   38219
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1680
         TabIndex        =   18
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame FraClase 
      Caption         =   "Clase Comercial"
      ForeColor       =   &H00800000&
      Height          =   1035
      Left            =   0
      TabIndex        =   15
      Top             =   1800
      Width           =   6855
      Begin vbp_Ventas.ctlDataCombo ctlCboClase 
         Height          =   315
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo ctlCboSubClase 
         Height          =   315
         Left            =   840
         TabIndex        =   7
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo ctlCboFamilia 
         Height          =   315
         Left            =   4080
         TabIndex        =   8
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo ctlCboCategoria 
         Height          =   315
         Left            =   4080
         TabIndex        =   9
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label lblSubCategoria 
         Caption         =   "SCateg:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblCategoria 
         Caption         =   "Categ:"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3360
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblClase 
         Caption         =   "Clase"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblSubClase 
         Caption         =   "S.Clase"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame FraLbLin 
      Caption         =   "Laboratorio y Linea"
      ForeColor       =   &H00800000&
      Height          =   1070
      Left            =   3360
      TabIndex        =   16
      Top             =   720
      Width           =   3500
      Begin vbp_Ventas.ctlDataCombo ctlCboLaboratorio 
         Height          =   315
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo ctlCboLinea 
         Height          =   315
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Lab"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblLinea 
         Caption         =   "Li&n"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
   End
   Begin MSComctlLib.Toolbar TblProd 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   1111
      ButtonWidth     =   1032
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "IlsImagen"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin VB.Frame Frame2 
         Caption         =   "&Stock"
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   6240
         TabIndex        =   25
         Top             =   0
         Width           =   735
         Begin VB.CheckBox chkStock 
            Height          =   255
            Left            =   240
            TabIndex        =   26
            ToolTipText     =   "Mostrar la columna Stock"
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "&Estado"
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   4080
         TabIndex        =   22
         Top             =   0
         Width           =   2175
         Begin vbp_Ventas.ctlDataCombo ctlCboEstado 
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            MatchEntry      =   1
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "&Producto"
         ForeColor       =   &H00800000&
         Height          =   705
         Left            =   1920
         TabIndex        =   21
         Top             =   0
         Width           =   2175
         Begin vbp_Ventas.ctlTextBox TxtProducto 
            Height          =   345
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
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
      End
   End
End
Attribute VB_Name = "frm_cat_Busquedacodigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objProducto As New clsProducto
Dim oraBusqueda As oraDynaset
Dim objLaboratorio As New clsLaboratorio
Dim objLinea As New clsLinea


Private Sub chkConStock_Click()
    If TxtProducto.Text <> "" Then sub_Buscar
End Sub

Private Sub chkStock_Click()
'    If chkStock.Value = 1 Then
'        fraStock.Visible = True
'    Else
'        fraStock.Visible = False
'    End If
    If TxtProducto.Text <> "" Then sub_Buscar
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo CtrlErr
    
    Call objProducto.ImprimeBusquedaProducto(oraBusqueda, IIf(chkStock.Value = "1", "S", "N"))
    
    Exit Sub
    
CtrlErr:
    MsgBox Err.Description, vbCritical + vbInformation, App.ProductName
End Sub


Private Sub Form_Load()
On Error GoTo Control

    setteaFormulario Me
    Carga_Laboratorio
    Carga_Estados
    Carga_Clase
    SetteaGrd
    'fraStock.Visible = False

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub ctlCboClase_Change()
    Carga_SubClase ctlCboClase.BoundText
End Sub

Private Sub ctlCboSubClase_Change()
    Carga_Familia ctlCboClase.BoundText, ctlCboSubClase.BoundText
End Sub

Private Sub ctlCboFamilia_Change()
    Carga_Categoria ctlCboClase.BoundText, _
                    ctlCboSubClase.BoundText, _
                    ctlCboFamilia.BoundText
End Sub


Private Sub ctlCboLaboratorio_Change()
    carga_Linea ctlCboLaboratorio.BoundText
End Sub

Sub Carga_Estados()

    On Error GoTo handle

    Set ctlCboEstado.RowSource = objProducto.ListaEstado
    ctlCboEstado.ListField = "DES"
    ctlCboEstado.BoundColumn = "COD"
    ctlCboEstado.BoundText = "*"
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
    
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Sub Carga_Laboratorio()

    On Error GoTo handle
     Set ctlCboLaboratorio.RowSource = objLaboratorio.Lista     '''objProducto.ListaLaboratorio
     ctlCboLaboratorio.ListField = "DES"
     ctlCboLaboratorio.BoundColumn = "COD"
     ctlCboLaboratorio.BoundText = "*"
     
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
     
End Sub

Sub carga_Linea(ByVal vstrCodLab As String)
    On Error GoTo handle
    Set ctlCboLinea.RowSource = objLinea.Lista(vstrCodLab)       '''objProducto.ListaLinea(vstrCodLab)
    ctlCboLinea.ListField = "DES"
    ctlCboLinea.BoundColumn = "COD"
    ctlCboLinea.BoundText = "*"
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Sub Carga_Clase()
    On Error GoTo handle
     Set ctlCboClase.RowSource = objProducto.ListaClase
     ctlCboClase.ListField = "DES"
     ctlCboClase.BoundColumn = "COD"
     ctlCboClase.BoundText = "*"
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Sub Carga_SubClase(ByVal vstrCodClase As String)
    On Error GoTo handle
     Set ctlCboSubClase.RowSource = objProducto.ListaSubclase(vstrCodClase)
     ctlCboSubClase.ListField = "DES"
     ctlCboSubClase.BoundColumn = "COD"
     ctlCboSubClase.BoundText = "*"
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Sub Carga_Familia(ByVal vstrCodClase As String, ByVal vstrCodSClase As String)
    On Error GoTo handle
     Set ctlCboFamilia.RowSource = objProducto.ListaFamilia(vstrCodClase, vstrCodSClase)
     ctlCboFamilia.ListField = "DES"
     ctlCboFamilia.BoundColumn = "COD"
     ctlCboFamilia.BoundText = "*"
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Sub Carga_Categoria(ByVal vstrCodClase As String, _
                    ByVal vstrCodSClase As String, _
                    ByVal vstrCodFam As String)
                    
    On Error GoTo handle
     Set ctlCboCategoria.RowSource = objProducto.ListaCategoria(vstrCodClase, vstrCodSClase, vstrCodFam)
     ctlCboCategoria.ListField = "DES"
     ctlCboCategoria.BoundColumn = "COD"
     ctlCboCategoria.BoundText = "*"
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Private Sub TblProd_ButtonClick(ByVal Button As MSComctlLib.Button)
     Select Case Button.Key
        Case "Nuevo"
            sub_Nuevo
        Case "Buscar"
            sub_Buscar
        Case "Salir"
            Unload Me
    End Select
End Sub

Sub sub_Nuevo()
    ctlCboLaboratorio.BoundText = "*": ctlCboLinea.BoundText = "*"
    ctlCboClase.BoundText = "*": ctlCboSubClase.BoundText = "*"
    ctlCboFamilia.BoundText = "*": ctlCboCategoria.BoundText = "*"
    ctlCboEstado.BoundText = "*"
    TxtProducto.Text = "": TxtProducto.SetFocus
    GrdProductos.Limpiar
End Sub

Sub sub_Buscar()
    
    'If chkStock = vbChecked Then
    
        'Set oraBusqueda = objProducto.ConsultaLocal(ctlCboLaboratorio.BoundText, _
                                                              ctlCboLinea.BoundText, _
                                                              ctlCboClase.BoundText, _
                                                              ctlCboSubClase.BoundText, _
                                                              ctlCboFamilia.BoundText, _
                                                              ctlCboCategoria.BoundText, _
                                                              TxtProducto.Text, _
                                                              ctlCboEstado.BoundText, _
                                                              objUsuario.CodigoLocal, _
                                                              chkStock.Value)
                                                              
                                                              
                                                              
        Dim strFaccionamiento As String
        
        On Error GoTo handle
        
        If Option1(0).Value = True Then strFaccionamiento = "1"
        If Option1(1).Value = True Then strFaccionamiento = "0"
        If Option1(2).Value = True Then strFaccionamiento = ""
        
        
        Set oraBusqueda = objProducto.ConsultaProductoLocal(ctlCboLaboratorio.BoundText, _
                                                              ctlCboLinea.BoundText, _
                                                              ctlCboClase.BoundText, _
                                                              ctlCboSubClase.BoundText, _
                                                              ctlCboFamilia.BoundText, _
                                                              ctlCboCategoria.BoundText, _
                                                              TxtProducto.Text, _
                                                              ctlCboEstado.BoundText, _
                                                              objUsuario.CodigoLocal, _
                                                              chkStock.Value, _
                                                              chkConStock.Value, strFaccionamiento)
    
    
        Set GrdProductos.DataSource = oraBusqueda
        
        SetteaGrd
    'Else
    '    Set GrdProductos.DataSource = objProducto.Consulta(ctlCboLaboratorio.BoundText, _
                                                              ctlCboLinea.BoundText, _
                                                              ctlCboClase.BoundText, _
                                                              ctlCboSubClase.BoundText, _
                                                              ctlCboFamilia.BoundText, _
                                                              ctlCboCategoria.BoundText, _
                                                              TxtProducto.Text, _
                                                              ctlCboEstado.BoundText)
    'End If
        Exit Sub
handle:

    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub TxtProducto_KeyPress(KeyAscii As Integer)
    TxtProducto.Tipo = AlfaNumerico
End Sub



Private Sub SetteaGrd()
Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant


arrCampos = Array("CODIGO", "DESCRIPCION", "ESTADO", "STOCK", "DES_LABORATORIO", "DES_LINEA", "CTD_FRACCIONAMIENTO", "VAL_MIN", "VAL_MAX")
arrCaption = Array("Codigo", "Descripción", "Estado", "Stock", "Laboratorio", "Línea", "Fracc.", "Min", "Max")
arrAncho = Array(800, 2500, 700, 900, 2500, 2500, 900, 800, 800)
arrAlineacion = Array(dbgCenter, dbgLeft, dbgCenter, dbgRight, dbgLeft, dbgLeft, dbgRight, dbgRight, dbgRight)
GrdProductos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
GrdProductos.Columns(1).WrapText = True
'GrdProductos.RowHeight = GrdProductos.RowHeight * 2.2
 
GrdProductos.Columns(3).Visible = True
If chkStock.Value = 0 Then GrdProductos.Columns(3).Visible = False

End Sub

