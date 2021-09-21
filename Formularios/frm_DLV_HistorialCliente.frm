VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_DLV_HistorialCliente 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Productos de pedido"
      TabPicture(0)   =   "frm_DLV_HistorialCliente.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CmdSalir"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Cmd"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraFormasPagos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdBuscar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboNum_Meses"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraProductos"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Ultimos pedidos"
      TabPicture(1)   =   "frm_DLV_HistorialCliente.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdExit"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdDetPedido"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdFind"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cboNroPedido"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.ComboBox cboNroPedido 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frm_DLV_HistorialCliente.frx":0038
         Left            =   840
         List            =   "frm_DLV_HistorialCliente.frx":003A
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   650
         Width           =   855
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4440
         Picture         =   "frm_DLV_HistorialCliente.frx":003C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos del Pedido"
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
         Height          =   4935
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   9375
         Begin vbp_Ventas.ctlGrilla grdPedido 
            Height          =   4215
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   7435
            MenuPopUp       =   0   'False
            Resalte         =   0   'False
         End
      End
      Begin VB.CommandButton cmdDetPedido 
         Caption         =   "&Detalle Pedido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4230
         Picture         =   "frm_DLV_HistorialCliente.frx":05C6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6960
         Width           =   1335
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8160
         Picture         =   "frm_DLV_HistorialCliente.frx":0B50
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6960
         Width           =   1335
      End
      Begin VB.Frame fraProductos 
         Caption         =   "F1 - Productos"
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
         Height          =   3015
         Left            =   -74760
         TabIndex        =   7
         Top             =   1440
         Width           =   9375
         Begin vbp_Ventas.ctlGrilla grdHistorico 
            Height          =   2655
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   4683
            MenuPopUp       =   0   'False
            Resalte         =   0   'False
         End
      End
      Begin VB.ComboBox cboNum_Meses 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frm_DLV_HistorialCliente.frx":10DA
         Left            =   -73440
         List            =   "frm_DLV_HistorialCliente.frx":10DC
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -70560
         Picture         =   "frm_DLV_HistorialCliente.frx":10DE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Frame fraFormasPagos 
         Caption         =   "F2 - Datos del Pedido"
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
         Height          =   2175
         Left            =   -74760
         TabIndex        =   3
         Top             =   4560
         Width           =   9375
         Begin vbp_Ventas.ctlGrilla grdFormasPago 
            Height          =   1815
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   3201
            MenuPopUp       =   0   'False
            Resalte         =   0   'False
         End
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Detalle Pedido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -70770
         Picture         =   "frm_DLV_HistorialCliente.frx":1668
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6960
         Width           =   1335
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -66840
         Picture         =   "frm_DLV_HistorialCliente.frx":1BF2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6960
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Los                           últimos pedidos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Width           =   2880
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Datos de los                           últimos pedidos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   9
         Top             =   720
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frm_DLV_HistorialCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objProforma As New clsProforma
Dim lstrCia As String
Dim lstrCod_Local As String
Dim lstrCod_Cliente As String

Public Sub Mostrar(ByVal vstrCia As String, _
                   ByVal vstrCod_Local As String, _
                   ByVal vstrCod_Cliente As String)
Dim intNum_Meses As Integer
Dim i As Integer
    lstrCia = vstrCia
    lstrCod_Local = vstrCod_Local
    lstrCod_Cliente = vstrCod_Cliente
    
    intNum_Meses = objProforma.NumeroMesesHistorico(lstrCia)
    
    If intNum_Meses < 1 Then intNum_Meses = 1
    
    cboNum_Meses.Clear
    
    For i = 1 To intNum_Meses
        cboNum_Meses.AddItem i
    Next i
    
    cboNum_Meses.Text = intNum_Meses
    cmdBuscar_Click
    
    Me.Show vbModal
End Sub


Private Sub cboNum_Meses_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
    
    
End Sub

Private Sub Cmd_Click()
    Dim frm As frm_VTA_DetallePedido
    
On Error GoTo Control
    
    If grdFormasPago.ApproxCount <= 0 Then Exit Sub

    Set frm = New frm_VTA_DetallePedido
    With frm
        .NumeroPedido = grdFormasPago.Columns("NUM_PROFORMA").Value
        .CodigoLocal = grdFormasPago.Columns("COD_LOCAL").Value
        .ReCargaDetPedido
        .Show vbModal
    End With
    Set frm = Nothing
    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdBuscar_Click()
On Error GoTo Control

    Set grdHistorico.DataSource = objProforma.ListaHistoricoCliente(lstrCia, _
                                                                    lstrCod_Local, _
                                                                    lstrCod_Cliente, _
                                                                    Val(cboNum_Meses.Text))
    
                                                                    

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
                                                                        
End Sub

Private Sub Consulta()
    On Error GoTo Control
    
      Set grdPedido.DataSource = objProforma.ListaUlt_Diez_Pedidos(lstrCod_Cliente, _
                                                                   Trim(cboNroPedido.Text))
    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error: " & Err.Number
    
End Sub

Private Sub cmdDetPedido_Click()
    Dim frm As frm_VTA_DetallePedido
        
    On Error GoTo Control
        
        If grdPedido.ApproxCount <= 0 Then Exit Sub
    
        Set frm = New frm_VTA_DetallePedido
        With frm
            .NumeroPedido = grdPedido.Columns("NUM_PROFORMA").Value
            .CodigoLocal = grdPedido.Columns("COD_LOCAL").Value
            .ReCargaDetPedido
            .Show vbModal
        End With
        Set frm = Nothing
        Exit Sub
Control:
        MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Consulta
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
'    If grdHistorico.ApproxCount > 0 Then
'        grdHistorico.SetFocus
'    End If
    
    cboNum_Meses.SetFocus
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            If grdHistorico.ApproxCount > 0 Then
                grdHistorico.SetFocus
            End If
        Case vbKeyF2
            If grdFormasPago.ApproxCount > 0 Then
                grdFormasPago.SetFocus
            End If
            
        Case vbKeyEscape
            Call cmdSalir_Click
    End Select
End Sub

Private Sub Form_Load()
    SeteaGrilla
    SeteaGrilla2
    SSTab1.Tab = 0
    
    Dim i As Integer
    
    For i = 1 To 10
        cboNroPedido.AddItem i
    Next i
    cboNroPedido.ListIndex = 9
    
    Consulta
End Sub

Sub SeteaGrilla()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    
    arrCampos = Array("NUM_PROFORMA", "COD_PRODUCTO", "DES_PRODUCTO", "FCH_MAXIMA", "CANTIDAD", "FRECUENCIA")
    arrCaption = Array("Pedido", "Código", "Descripción", "Ultimo Pedido", "Cantidad", "Frecuencia")
    arrAncho = Array(0, 900, 4800, 1250, 700, 700)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgLeft, dbgCenter, dbgCenter, dbgCenter)
    
    grdHistorico.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdHistorico.Columns(0).Visible = False
    
    '---------------------
    '---------------------
    arrCampos = Array("NUM_PROFORMA", "COD_CLIENTE", "DES_CLIENTE", "COD_TIPO_DOCUMENTO", "DES_ESTADO", "FCH_REGISTRA", "MTO_TOTAL", "COD_LOCAL")
    arrCaption = Array("Pedido", "Codigo", "Cliente", "Tipo Doc", "Estado", "Fecha", "Total", "CodigoLocal")
    arrAncho = Array(1200, 0, 2400, 900, 1200, 1200, 900, 900)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgRight, dbgCenter)
    
    grdFormasPago.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdFormasPago.Columns("COD_CLIENTE").Visible = False
    grdFormasPago.Columns("COD_LOCAL").Visible = False
End Sub

Private Sub grdFormasPago_DblClick()
    Cmd_Click
End Sub

Private Sub grdFormasPago_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            grdFormasPago_DblClick
    End Select
End Sub

Private Sub grdHistorico_RegistroSeleccionado(ByVal DatoColumna0 As String)
On Error GoTo Control

    If grdHistorico.ApproxCount <= 0 Then Exit Sub
    
    '''''Set grdFormasPago.DataSource = objProforma.ListaCabecera(objUsuario.CodigoEmpresa, _
                                                             objUsuario.CodigoLocal, _
                                                             grdHistorico.Columns("NUM_PROFORMA").Value)
                                                             
                                                             
    Set grdFormasPago.DataSource = objProforma.ListaUltimoPedidos(lstrCia, _
                                                                    lstrCod_Local, _
                                                                    lstrCod_Cliente, _
                                                                    grdHistorico.Columns("COD_PRODUCTO").Value, _
                                                                    Val(cboNum_Meses.Text))
                                                             
                                                             

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
                                                             
End Sub

Private Sub grdPedido_DblClick()
    cmdDetPedido_Click
End Sub

Sub SeteaGrilla2()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    
    arrCampos = Array("NUM_PROFORMA", "COD_PRODUCTO", "DES_PRODUCTO", "FCH_MAXIMA", "CANTIDAD", "FRECUENCIA")
    arrCaption = Array("Pedido", "Código", "Descripción", "Ultimo Pedido", "Cantidad", "Frecuencia")
    arrAncho = Array(0, 900, 4800, 1250, 700, 700)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgLeft, dbgCenter, dbgCenter, dbgCenter)
    
    grdHistorico.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdHistorico.Columns(0).Visible = False
    
    '---------------------
    '---------------------
    arrCampos = Array("NUM_PROFORMA", "COD_CLIENTE", "DES_CLIENTE", "COD_TIPO_DOCUMENTO", "DES_ESTADO", "FCH_REGISTRA", "MTO_TOTAL", "COD_LOCAL")
    arrCaption = Array("Pedido", "Codigo", "Cliente", "Tipo Doc", "Estado", "Fecha", "Total", "CodigoLocal")
    arrAncho = Array(1200, 0, 2400, 900, 1200, 1200, 900, 900)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgRight, dbgCenter)
    
    grdPedido.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdPedido.Columns("COD_CLIENTE").Visible = False
    grdPedido.Columns("COD_LOCAL").Visible = False
End Sub


