VERSION 5.00
Begin VB.Form frm_DLV_CambioLocalDespacho 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio Local de Despacho"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3810
   Icon            =   "frm_DLV_CambioLocalDespacho.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   3810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Cambiar a "
      Height          =   1455
      Left            =   60
      TabIndex        =   12
      Top             =   2280
      Width           =   3675
      Begin vbp_Ventas.ctlDataCombo CboZona 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
      End
      Begin vbp_Ventas.ctlDataCombo CboLocal 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   960
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboCia 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cia :"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Zona"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Local"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1020
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos actuales"
      Height          =   1575
      Left            =   60
      TabIndex        =   7
      Top             =   600
      Width           =   3675
      Begin vbp_Ventas.ctlDataCombo cboCia2 
         Height          =   315
         Left            =   1080
         TabIndex        =   15
         Top             =   360
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         ColorDefault    =   -2147483633
         ColorDefault    =   -2147483633
         MatchEntry      =   1
         EnabledFoco     =   0   'False
         Enabled         =   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cia :"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   400
         Width           =   315
      End
      Begin VB.Label LblLocalRef2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   16
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Local Ref :"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1215
         Width           =   780
      End
      Begin VB.Label LblLocalRef 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   10
         Top             =   1140
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Zona :"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   795
         Width           =   465
      End
      Begin VB.Label LblZona 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1080
         TabIndex        =   8
         Top             =   720
         Width           =   1980
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   2580
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   1260
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Pedido 
      AutoSize        =   -1  'True
      Caption         =   "Pedido :"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   180
      Width           =   585
   End
   Begin VB.Label LblPedido 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1380
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frm_DLV_CambioLocalDespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objZona As New clsZona
Dim objProforma As New clsProforma
Public pZonax As String
Public pLocalx As String
Public pPedidox As String
Public pCiaRefx As String
Public pLocalRefx As String
Public pLocalSapfx As String

Private Sub cboCia_Change()
On Error GoTo Control

   CboLocal.Text = ""
   'Set CboLocal.RowSource = objZona.Lista_Locales(objUsuario.CodigoEmpresa, CboZona.BoundText)
   Set CboLocal.RowSource = objZona.Lista_Locales(cboCia.BoundText, CboZona.BoundText)
   CboLocal.BoundColumn = "COD_LOCAL"
   CboLocal.ListField = "LOCAL_DEX"

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cboLocal_KeyDown(KeyCode As Integer, Shift As Integer)
    sub_Tecla KeyCode
End Sub

Private Sub CboZona_KeyDown(KeyCode As Integer, Shift As Integer)
    sub_Tecla KeyCode
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    sub_Tecla KeyCode
End Sub

Private Sub Form_Load()
    Set cboCia.RowSource = gclsOracle.FN_Cursor("btlprod.pkg_local.fn_lista_marca", 0)
    cboCia.ListField = "Des"
    cboCia.BoundColumn = "Cod"
    
    Set cboCia2.RowSource = gclsOracle.FN_Cursor("btlprod.pkg_local.fn_lista_marca", 0)
    cboCia2.ListField = "Des"
    cboCia2.BoundColumn = "Cod"
    
    Set CboZona.RowSource = objZona.Lista
    CboZona.ListField = "DES_ZONA"
    CboZona.BoundColumn = "COD_ZONA"
    
    cboCia2.BoundText = pCiaRefx
    LblZona.Caption = pZonax
    LblLocalRef.Caption = pLocalRefx
    LblLocalRef2.Caption = pLocalSapfx
    LblPedido.Caption = pPedidox
End Sub

Private Sub cboZona_Click(Area As Integer)
On Error GoTo Control

   CboLocal.Text = ""
   'Set CboLocal.RowSource = objZona.Lista_Locales(objUsuario.CodigoEmpresa, CboZona.BoundText)
   Set CboLocal.RowSource = objZona.Lista_Locales(cboCia.BoundText, CboZona.BoundText)
   CboLocal.BoundColumn = "COD_LOCAL"
   CboLocal.ListField = "LOCAL_DEX"

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdGrabar_Click()
    On Error GoTo registra
    Dim a As String
    
    If CboLocal.BoundText = "" Then MsgBox "Seleccione el local a trasladar el pedido", vbCritical, Caption: Exit Sub
    a = objProforma.GrabarLocalDespacho(objUsuario.CodigoEmpresa, _
                                        Trim(LblPedido.Caption), _
                                        Trim(CboLocal.BoundText))
                    
    If a = "" Then
        'MsgBox "Se grabo con exito el local de despacho", vbInformation + vbOKOnly, "Grabar"
        Unload Me
    Else
        MsgBox a, vbCritical + vbOKOnly, "Atención"
    End If
Exit Sub
registra:
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub sub_Tecla(ByVal KeyCode As Integer)
If KeyCode = vbKeyEscape Then
    cmdCancelar_Click
End If
End Sub
