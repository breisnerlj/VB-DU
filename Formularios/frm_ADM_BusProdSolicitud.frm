VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_ADM_BusProdSolicitud 
   BorderStyle     =   0  'None
   Caption         =   "Búsqueda de Productos - Solicitud"
   ClientHeight    =   5070
   ClientLeft      =   7005
   ClientTop       =   2505
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlGrilla grdProductos 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7011
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin ORADCLibCtl.ORADC oradcProductos 
      Height          =   255
      Left            =   4560
      Top             =   7800
      Visible         =   0   'False
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   450
      _StockProps     =   207
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   ""
      Connect         =   ""
      RecordSource    =   ""
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   1111
      ButtonWidth     =   1217
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "IlsImagen"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Buscar"
            Key             =   "Buscar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Key             =   "Salir"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   1440
         TabIndex        =   1
         Top             =   0
         Width           =   5415
         Begin vbp_Ventas.ctlTextBox txtDesProducto 
            Height          =   375
            Left            =   960
            TabIndex        =   3
            Top             =   120
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   661
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Producto:"
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
            Left            =   0
            TabIndex        =   2
            Top             =   240
            Width           =   840
         End
      End
   End
   Begin MSComctlLib.ImageList IlsImagen 
      Left            =   2160
      Top             =   1320
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
            Picture         =   "frm_ADM_BusProdSolicitud.frx":0000
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BusProdSolicitud.frx":059A
            Key             =   "buscar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BusProdSolicitud.frx":0B34
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BusProdSolicitud.frx":10CE
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BusProdSolicitud.frx":1668
            Key             =   "Chek"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BusProdSolicitud.frx":1C02
            Key             =   "Bien"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BusProdSolicitud.frx":219C
            Key             =   "Agregar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BusProdSolicitud.frx":2736
            Key             =   "Atras"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BusProdSolicitud.frx":2CD0
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BusProdSolicitud.frx":326A
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BusProdSolicitud.frx":3804
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_BusProdSolicitud.frx":3D9E
            Key             =   "Hora"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblReg 
      AutoSize        =   -1  'True
      Caption         =   "Registros: 0"
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
      Left            =   120
      TabIndex        =   5
      Top             =   4845
      Width           =   1035
   End
End
Attribute VB_Name = "frm_ADM_BusProdSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim lodynConsulta As oraDynaset
Dim lxdbProd As New XArrayDB
Dim lvarNewRegistro As Variant
Dim objSSGG As New clsSSGG
Public blnNuevoForm As Boolean

Sub sub_sge_muestra_productos()
    On Error GoTo ERROR
    Dim StrSql As String, strFiltro As String
    Screen.MousePointer = vbHourglass
    'StrSql = " SELECT COD_PRODUCTO_GEN,DES_PRODUCTO " & _
             " FROM SSGG.MAE_PRODUCTO_GENERICO" & _
             " WHERE COD_CATEG_PRODUCTO IN " & _
             " (SELECT COD_CATEG_PRODUCTO FROM SSGG.MAE_CATEGORIA WHERE FLG_ACTIVO='1') " & _
             " AND (UPPER(DES_PRODUCTO) LIKE '" & UCase(Trim(txtDesProducto.Text)) & "%' OR " & _
             " UPPER(DES_PRODUCTO) LIKE '% " & UCase(Trim(txtDesProducto.Text)) & "%') " & _
             " ORDER BY DES_PRODUCTO"
             '"AND FLG_MUESTRA = '1' "
   ' Set lodynConsulta = godbVentas.CreateDynaset(StrSql, 0&)
    'trFiltro = "'" & UCase(Trim(txtDesProducto.Text)) & "'"
    Set lodynConsulta = objSSGG.listaProductoGenerico(UCase(Trim(txtDesProducto.Text)))
    Set grdProductos.DataSource = lodynConsulta
    lblReg.Caption = "Registros:" & lodynConsulta.RecordCount
    grdProductos.Rebind
    grdProductos.Columns(0).Width = 1500
    grdProductos.Columns(1).Width = 3500
    
    If lodynConsulta.RecordCount = 0 Then
        grdProductos.Enabled = False
    Else
        grdProductos.Enabled = True
    End If
    Screen.MousePointer = vbNormal
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdBuscar_Click()
    Call sub_sge_muestra_productos
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        'mdiPrincipal.picComandos.Enabled = True
        Unload Me
End Select
End Sub

Private Sub Form_Load()
    'On Error GoTo ERROR
    Me.left = 0
    Me.top = 0
    Call spSetGrdDetalle(grdProductos)

    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub sub_sge_Agregar_Producto(ByRef varNewRegistro As Variant, ByVal xdbProd As XArrayDB)
    On Error GoTo ERROR
    Set lxdbProd = xdbProd
    lvarNewRegistro = Array()
    Me.Show vbModal
    Set lxdbProd = Nothing
    varNewRegistro = lvarNewRegistro
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub sub_sge_SeleccionProducto()
    On Error GoTo ERROR
    Dim ldblCantidad As Double
    Screen.MousePointer = vbHourglass
    If lxdbProd.UpperBound(1) <> (lxdbProd.LowerBound(1) - 1) Then
        If lxdbProd.Find(0, 1, lodynConsulta("COD_PRODUCTO_GEN").Value, XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING) > lxdbProd.LowerBound(1) - 1 Then
            MsgBox "El Producto ya ha sido seleccionado", vbExclamation, "Aviso"
            Screen.MousePointer = vbNormal
            txtDesProducto.selection
            Exit Sub
        End If
    End If
    ldblCantidad = 0
    Call frm_ADM_CantProdSolicitud.sub_sge_Cantidad(ldblCantidad, grdProductos.Columns(1).Text)
    If ldblCantidad > 0 Then
        lvarNewRegistro = Array(lodynConsulta("COD_PRODUCTO_GEN").Value, _
                                lodynConsulta("DES_PRODUCTO").Value, _
                                ldblCantidad)
        Screen.MousePointer = vbNormal
        mdiPrincipal.picComandos.Enabled = Not blnNuevoForm
        Unload Me
    Else
        grdProductos.SetFocus
    End If
    Screen.MousePointer = vbNormal
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub


Private Sub grdProductos_DblClick()
    On Error GoTo ERROR
    If lodynConsulta.RecordCount <> 0 Then
        Call sub_sge_SeleccionProducto
    End If
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub


Private Sub spSetGrdDetalle(ByRef rgrd As ctlGrilla)
Dim pvarAncho As Variant
Dim pvarTitulo As Variant
Dim pvarAlinea As Variant
Dim pvarCampoDato As Variant

    pvarAncho = Array(800, 6450)
    pvarTitulo = Array("Código", "DESCRIPCION")
    pvarAlinea = Array(2, 0)
    pvarCampoDato = Array("COD_PRODUCTO_GEN", "DES_PRODUCTO")

   'objSSGG.spGrilla_Carga rgrd, pvarTitulo, pvarAncho, pvarAlinea, pvarCampoDato

   rgrd.MarqueeStyle = dbgHighlightRow
   'rgrd.HeadBackColor = &H80000007
   'rgrd.HeadForeColor = &HFFFF&
   rgrd.RowHeight = 0
   rgrd.RowHeight = 1 * 320
   rgrd.HeadLines = 1
   'rgrd.Font.Size = 8
   rgrd.Columns(0).Width = 1500
   rgrd.Columns(1).Width = 5500
   
   'rgrd.Styles(5).Font.Size = 8
   'grdProductos.Styles(5).BackColor = &H8000000C
End Sub


Private Sub grdProductos_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
    If KeyCode = vbKeyUp And grdProductos.Row = 0 Then txtDesProducto.SetFocus
    If KeyCode = 13 And lodynConsulta.RecordCount <> 0 Then
        Call sub_sge_SeleccionProducto
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        Call sub_sge_muestra_productos
    Case 2
        mdiPrincipal.picComandos.Enabled = Not blnNuevoForm
        Unload Me
End Select
End Sub
Private Sub txtDesProducto_Change()
    If grdProductos.DataSource Is Nothing Then Exit Sub
    grdProductos.DataSource.FindFirst " DES_PRODUCTO " & " LIKE '" & CStr(txtDesProducto.Text) & "%'"
End Sub
Private Sub txtDesProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
    If KeyCode = vbKeyDown And grdProductos.ApproxCount > 0 Then grdProductos.SetFocus
    If KeyCode = 13 Then Call sub_sge_muestra_productos
    If KeyCode = vbKeyEscape Then Unload Me
    If KeyCode = 40 And grdProductos.Row >= 0 Then grdProductos.SetFocus
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub


Private Sub grdProductos_GotFocus()
        'grdProductos.Styles(5).BackColor = &H8000000D
        'grdProductos.Styles(5).Font.Bold = True
End Sub

Private Sub grdProductos_LostFocus()
        'grdProductos.Styles(5).BackColor = &H8000000C
        'grdProductos.Styles(5).Font.Bold = False
End Sub

