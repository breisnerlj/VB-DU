VERSION 5.00
Begin VB.Form frm_VTA_Cliente_Bus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda de Cliente"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Editar"
      Height          =   435
      Left            =   7080
      TabIndex        =   15
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Nuevo"
      Height          =   435
      Left            =   5760
      TabIndex        =   14
      Top             =   5880
      Width           =   1215
   End
   Begin vbp_Ventas.ctlGrilla GrdBusCliente 
      Height          =   3975
      Left            =   0
      TabIndex        =   10
      Top             =   1800
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7011
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8415
      Begin VB.Frame Frame2 
         Caption         =   "Avanzado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   8175
         Begin VB.CheckBox chkActivo 
            Caption         =   "Inactivos"
            Height          =   375
            Left            =   6600
            TabIndex        =   11
            Top             =   120
            Width           =   1215
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Por Documento"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   690
            Width           =   1455
         End
         Begin vbp_Ventas.ctlDataCombo ctlCboTipCliente 
            Height          =   315
            Left            =   1920
            TabIndex        =   3
            Top             =   360
            Width           =   2895
            _ExtentX        =   4260
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin vbp_Ventas.ctlTextBox txtNumDocumento 
            Height          =   315
            Left            =   4800
            TabIndex        =   6
            Top             =   720
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            Tipo            =   8
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
         Begin vbp_Ventas.ctlDataCombo cboTipoDocumentoCli 
            Height          =   315
            Left            =   1920
            TabIndex        =   5
            Top             =   720
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            MatchEntry      =   1
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Por Tipo Cliente"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   330
            Width           =   1575
         End
         Begin VB.CommandButton Command2 
            Cancel          =   -1  'True
            Caption         =   "Command2"
            Height          =   195
            Left            =   5880
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   840
            Width           =   255
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Busqueda"
         Height          =   375
         Left            =   4800
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin vbp_Ventas.ctlTextBox txtNombres 
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         Tipo            =   8
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
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "F1"
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
         Left            =   2280
         TabIndex        =   18
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "F1"
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
         Index           =   8
         Left            =   240
         TabIndex        =   16
         Top             =   292
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "[F3] Para iniciar la Busqueda"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   6000
         TabIndex        =   12
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CRITERIO"
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
         Left            =   600
         TabIndex        =   8
         Top             =   330
         Width           =   900
      End
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
      Index           =   2
      Left            =   7560
      TabIndex        =   19
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label lblTitle 
      Caption         =   "F2"
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
      Height          =   255
      Index           =   0
      Left            =   6180
      TabIndex        =   17
      Top             =   6360
      Width           =   375
   End
End
Attribute VB_Name = "frm_VTA_Cliente_Bus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''Variables publicas
Public Out_CodigoCliente As String
Public Out_NombreCliente As String
Public Out_NumeroId As String
Public Out_Tipo As String
Public Out_Direccion As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim objCliente As New clsCliente


Private Sub Check1_Click()
    If Check1.Value = 1 Then
        ctlCboTipCliente.Enabled = True
    Else
        ctlCboTipCliente.Enabled = False
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        cboTipoDocumentoCli.Enabled = True
    Else
        cboTipoDocumentoCli.Enabled = False: txtNumDocumento.Text = ""
    End If
End Sub

Private Sub Command1_Click()
On Error GoTo Control
If Check2.Value = 0 And Check1.Value = 0 And Len(Trim(Me.txtNombres.Text)) < 3 Then MsgBox "Debe ingresar un dato para realizar la búsqueda.", vbOKOnly + vbExclamation, "Error": txtNombres.SetFocus: Exit Sub
        BuscaCliente

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
On Error GoTo Control
    ''09/06/2008 - Comentado por JLopez xq la consulta de cliente se hace en Load del formulario
    ''frm_VTA_Cliente.ctlCliente1.Cargar
    frm_VTA_Cliente.strCodigo = ""
    frm_VTA_Cliente.ctlCliente1.XTipoFuncion = "Nuevo"
    frm_VTA_Cliente.CargarValores
    frm_VTA_Cliente.Show vbModal

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
    
End Sub

Private Sub Command4_Click()
On Error GoTo Control

If GrdBusCliente.ApproxCount = 0 Then Exit Sub

    ''09/06/2008 - Comentado por JLopez xq la consulta de cliente se hace en Load del formulario
    ''frm_VTA_Cliente.ctlCliente1.ConsultaCliente GrdBusCliente.Columns(0).Value
    frm_VTA_Cliente.ctlCliente1.XTipoFuncion = "Editar"
    frm_VTA_Cliente.ctlCliente1.CodDireccionCli = "" & GrdBusCliente.Columns("COD_DIRECCION_CLI").Value
    frm_VTA_Cliente.strCodigo = GrdBusCliente.Columns(0).Value
    frm_VTA_Cliente.CargarValores
    frm_VTA_Cliente.Show vbModal


   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'If KeyCode = vbKeyF3 Then

        Select Case KeyCode
            Case vbKeyF1
                txtNombres.SetFocus
            Case vbKeyF2
                Call Command3_Click
            Case vbKeyF3
                Call Command1_Click
            Case vbKeyF4
                Call Command4_Click
        End Select

End Sub

Private Sub Form_Load()
Out_CodigoCliente = ""
Out_NombreCliente = ""
Out_NumeroId = ""
Out_Tipo = ""
Out_Direccion = ""

    Command1.Default = False
    ParametroInicio
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo handle
    Set objCliente = Nothing
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Sub ParametroInicio()
On Error GoTo handle
    Set ctlCboTipCliente.RowSource = objCliente.ListaTipo
    ctlCboTipCliente.ListField = "DES"
    ctlCboTipCliente.BoundColumn = "COD"
    ctlCboTipCliente.BoundText = "*"
    Set cboTipoDocumentoCli.RowSource = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_TIPO_DOCUMENTO", 0, "")
    cboTipoDocumentoCli.ListField = "DES_DOCUMENTO_IDENTIDAD"
    cboTipoDocumentoCli.BoundColumn = "COD_DOCUMENTO_IDENTIDAD"
    ctlCboTipCliente.Enabled = False: cboTipoDocumentoCli.Enabled = False
    setteaGrilla
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub
Private Sub setteaGrilla()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    
    arrCampos = Array("COD_CLIENTE", "Nombre", "Numero", "FLG_TIPO_JURIDICA", "COD_DIRECCION_CLI")
    arrCaption = Array("Codigo", "Nombre", "Documento", "Tipo", "CodDireccion")
    arrAncho = Array(1000, 5000, 1300, 0, 0)
    arrAlineacion = Array(vbAlignNone, vbAlignNone, vbAlignNone, vbAlignLeft, vbAlignNone, dbgGeneral)
    GrdBusCliente.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    GrdBusCliente.Columns(3).Visible = False
    GrdBusCliente.Columns(4).Visible = False
    GrdBusCliente.Columns(3).AllowSizing = False
    GrdBusCliente.Columns(4).AllowSizing = False
    
End Sub

Private Sub BuscaCliente()
'On Error GoTo handle
    Dim strctlCboTipCliente As String
    Dim strcboTipoDocumento As String
    Dim strNumeroDocumento As String
    Dim strflgActivo As String
'''
'''    If Len(txtNombres.Text) <= 2 Then
'''        MsgBox "Ingrese como minímo tres letras para la busqueda", vbCritical, Caption: GrdBusCliente.Limpiar: Exit Sub
'''    End If
'''
    
    If Check2.Value = 1 Then
        If Trim(txtNumDocumento.Text) = "" Then MsgBox "Ingrese el Número del Documento", vbInformation, "Aviso": txtNumDocumento.SetFocus: Exit Sub
    End If
    strctlCboTipCliente = ctlCboTipCliente.BoundText
    If Check1.Value = 0 Then strctlCboTipCliente = ""
    
    strcboTipoDocumento = cboTipoDocumentoCli.BoundText
    strNumeroDocumento = txtNumDocumento.Text
    If Check2.Value = 0 Then strcboTipoDocumento = "": strNumeroDocumento = ""
    
    strflgActivo = chkActivo.Value
    Set GrdBusCliente.DataSource = objCliente.ListaCliente(txtNombres.Text, strctlCboTipCliente, strcboTipoDocumento, strNumeroDocumento, "") 'strflgActivo)
    If GrdBusCliente.ApproxCount > 0 Then GrdBusCliente.SetFocus
    
    If GrdBusCliente.ApproxCount = 0 Then MsgBox "No se encontro el criterio de Busqueda", vbInformation + vbOKOnly, App.ProductName
    
    Exit Sub
'handle:
'MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub GrdBusCliente_DblClick()
    If GrdBusCliente.ApproxCount = 0 Then Exit Sub
    Out_CodigoCliente = "" & GrdBusCliente.DataSource("COD_CLIENTE").Value
    Out_NombreCliente = "" & GrdBusCliente.DataSource("Nombre").Value
    Out_NumeroId = "" & GrdBusCliente.DataSource("Numero").Value
    Out_Tipo = "" & GrdBusCliente.DataSource("FLG_TIPO_JURIDICA").Value
    Out_Direccion = "" & GrdBusCliente.DataSource("direccion").Value
    Unload Me

End Sub


Private Sub GrdBusCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then GrdBusCliente_DblClick

End Sub

Private Sub txtNombres_GotFocus()
    Command1.Default = True
End Sub

Private Sub txtNombres_LostFocus()
    Command1.Default = False
End Sub

Private Sub txtNumDocumento_GotFocus()
    Command1.Default = True
End Sub

