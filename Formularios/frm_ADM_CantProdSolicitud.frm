VERSION 5.00
Begin VB.Form frm_ADM_CantProdSolicitud 
   BorderStyle     =   0  'None
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      TabIndex        =   6
      Top             =   1080
      Width           =   840
   End
   Begin vbp_Ventas.ctlTextBox txtCantidad 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   855
      _ExtentX        =   1508
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   675
      Left            =   60
      TabIndex        =   3
      Top             =   300
      Width           =   4515
      Begin VB.TextBox txtDesProducto 
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
         Height          =   620
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   0
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      TabIndex        =   0
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label lblCantidad 
      Caption         =   "CANTIDAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1110
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frm_ADM_CantProdSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private ldblCantidadUnid As Double
Private lstrDesProducto As String
Private lblnOK As Boolean

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
    Call txtCantidad_KeyPress(13)
End Sub


Private Sub Form_Load()
    On Error GoTo ERROR
    lblnOK = False
    txtDesProducto.Text = lstrDesProducto
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub sub_sge_Cantidad(ByRef rdblCantidad As Double, ByVal vstrDesProducto As String)
    On Error GoTo ERROR
    txtCantidad.Text = rdblCantidad
    txtDesProducto.Text = vstrDesProducto
    lstrDesProducto = vstrDesProducto
    Screen.MousePointer = vbNormal
    Me.Show vbModal
    If lblnOK Then rdblCantidad = ldblCantidadUnid
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub


Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    On Error GoTo ERROR
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(Trim(txtCantidad.Text)) = 0 Then
            MsgBox "Cantidad no permitida", vbCritical, "Error"
            txtCantidad.Text = ""
            txtCantidad.SetFocus
            Exit Sub
        End If
        If IsNumeric(txtCantidad.Text) Then
            lblnOK = True
            ldblCantidadUnid = Val(txtCantidad.Text)
            Unload Me
        End If
    End If
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Activate()
    txtCantidad.Focus
End Sub

