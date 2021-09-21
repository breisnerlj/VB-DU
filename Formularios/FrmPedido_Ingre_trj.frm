VERSION 5.00
Begin VB.Form FrmPedido_Ingre_trj 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin vbp_Ventas.ctlTextBox ctlTxtTarjeta 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      PasswordChar    =   "*"
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
   Begin VB.Label lblSubTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Escanear tarjeta:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SOLICITE TARJETA DE PUNTOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4440
   End
End
Attribute VB_Name = "FrmPedido_Ingre_trj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvarTxtTarjeta As String
Private mvarEscaneoTarjeta As Boolean
Private mvarObligatorioEscaneo As Boolean

Private Sub ctlTxtTarjeta_KeyDown(KeyCode As Integer, Shift As Integer)
    If mvarObligatorioEscaneo Then
        If EsEscaneado(KeyCode, Shift) = True Then
            mvarEscaneoTarjeta = True
        Else
            mvarEscaneoTarjeta = False
        End If
    End If
End Sub

Private Sub ctlTxtTarjeta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mvarTxtTarjeta = Me.ctlTxtTarjeta.Text
        Unload Me
    End If
    If KeyAscii = 27 Then
        mvarTxtTarjeta = ""
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    Me.ctlTxtTarjeta.Focus
End Sub

Private Sub Form_Load()
    Me.ctlTxtTarjeta.Text = ""
    mvarEscaneoTarjeta = False
    mvarObligatorioEscaneo = True
End Sub

Public Function ObtenerTarjeta(ByVal vTitulo As String, _
                               Optional ByVal Tipo As ETipoDocumentoMonedero, _
                               Optional ByVal bEscanear As Boolean = True)
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Me.Caption = vTitulo
    mvarObligatorioEscaneo = bEscanear
    lblSubTitulo.Caption = ""
    
    Select Case Tipo
        Case ETipoDocumentoMonedero.eeDniCliente
            lblTitulo.Caption = "SOLICITE DNI/CE DEL CLIENTE"
            lblSubTitulo.Caption = IIf(mvarObligatorioEscaneo, "Escanear", "Ingresar") & " DNI/CE del cliente"
        Case ETipoDocumentoMonedero.eeTarjetaMonedero
            lblTitulo.Caption = "SOLICITE TARJETA MONEDERO"
            lblSubTitulo.Caption = IIf(mvarObligatorioEscaneo, "Escanear", "Ingresar") & " la Tarjeta"
    End Select
    
    Me.Show vbModal
    ObtenerTarjeta = mvarTxtTarjeta
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (mvarObligatorioEscaneo And Not mvarEscaneoTarjeta) And Me.ctlTxtTarjeta.Text <> "" Then
        MsgBox "Error: Debe " & lblSubTitulo.Caption & ".", vbCritical + vbOKOnly, App.ProductName
        Me.ctlTxtTarjeta.Text = ""
        mvarTxtTarjeta = ""
        Me.ctlTxtTarjeta.Focus
        Cancel = True
    End If
End Sub
