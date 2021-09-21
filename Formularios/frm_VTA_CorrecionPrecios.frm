VERSION 5.00
Begin VB.Form frm_VTA_CorrecionPrecios 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlDataCombo dbcAutoriza 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlDataCombo dbcMotivo 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   3420
      Picture         =   "frm_VTA_CorrecionPrecios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   4725
      Picture         =   "frm_VTA_CorrecionPrecios.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Descuento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   10
      Top             =   1920
      Width           =   5175
      Begin vbp_Ventas.ctlTextBox txtImporte 
         Height          =   315
         Left            =   2880
         TabIndex        =   5
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Tipo            =   6
         Alignment       =   1
         MaxLength       =   5
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
      Begin VB.OptionButton Tipo 
         Caption         =   "Importe"
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
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Tipo 
         Caption         =   "Porcentaje"
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
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin vbp_Ventas.ctlDataCombo ctlDataCombo2 
      Height          =   315
      Left            =   2280
      TabIndex        =   15
      Top             =   2520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
   End
   Begin vbp_Ventas.ctlDataCombo ctlDataCombo1 
      Height          =   315
      Left            =   2280
      TabIndex        =   16
      Top             =   1920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Autorización"
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
      Index           =   3
      Left            =   840
      TabIndex        =   20
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Motivo"
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
      Left            =   840
      TabIndex        =   19
      Top             =   960
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Motivo"
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
      Left            =   1080
      TabIndex        =   18
      Top             =   2040
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Autorización"
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
      Index           =   1
      Left            =   840
      TabIndex        =   17
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6240
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Shift+Enter"
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
      Index           =   12
      Left            =   3360
      TabIndex        =   14
      Top             =   4740
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Esc"
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
      Index           =   11
      Left            =   5070
      TabIndex        =   13
      Top             =   4740
      Width           =   390
   End
   Begin VB.Label lblTotalcDscto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   12
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblPrecioAct 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Total c/Dscto. S/."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   11
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Total S/."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   9
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Corrección de Precios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Index           =   4
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   2385
   End
End
Attribute VB_Name = "frm_VTA_CorrecionPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strCodigo As String
Dim strDescripcion As String
Dim strFlgFraccion As String
Dim intCantidad As Integer
Dim dblPrecio As Double
Dim TipoVenta As TipoVenta
Dim strRegalo As String
Dim TipoValor As eTipoValor
Dim objAutorizacion As New clsAutorizacion






Public Sub datos(ByVal pstrCodigo As String, _
                        ByVal pstrDescripcion As String, _
                        ByVal pstrFlgFraccion As String, _
                        ByVal pintCantidad As Integer, _
                        ByVal pdblPrecio As Double, _
                        ByVal pstrTipoVenta As String, _
                        ByVal pstrRegalo As String)
                        
On Error GoTo handle
strCodigo = pstrCodigo
strDescripcion = pstrDescripcion
strFlgFraccion = pstrFlgFraccion
intCantidad = pintCantidad
dblPrecio = pdblPrecio
TipoVenta = pstrTipoVenta
strRegalo = pstrRegalo
lblPrecioAct.Caption = dblPrecio
Tipo(0).Value = True

Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub


Private Sub cmdAceptar_Click()
On Error GoTo handle
    If Val(lblTotalcDscto.Caption) <= 0 Then
        MsgBox "Precio cálculado es incorrecto, verifique", vbCritical + vbOKOnly, "Atención"
        Exit Sub
    End If
    
    If dbcMotivo.Text = "" Then
        MsgBox "Ingrese el motivo de la corrección, verifique ", vbCritical + vbOKOnly, "Atención"
        Exit Sub
    End If
    
    If dbcAutoriza.Text = "" Then
        MsgBox "Ingrese el usuario que autoriza la corrección, verifique ", vbCritical + vbOKOnly, "Atención"
        Exit Sub
    End If
    
    objVenta.CorreccionPrecio strCodigo, _
                        strDescripcion, _
                        intCantidad, _
                        strFlgFraccion, _
                        Val(lblTotalcDscto.Caption), _
                        TipoVenta, _
                        strRegalo, _
                        TipoValor, _
                        Val(txtImporte.Text), dblPrecio, dbcMotivo.BoundText, dbcAutoriza.BoundText

    frmPedido.Cal_Montos
    frmPedido.grdPedido.Rebind
    Unload Me

Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub dbcAutoriza_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdCancelar_Click
End Sub

Private Sub dbcMotivo_Change()
On Error GoTo handle
        dbcAutoriza.Text = ""
        Set dbcAutoriza.RowSource = objAutorizacion.AutorizaUsuario(dbcMotivo.BoundText)
        dbcAutoriza.ListField = "NOMBRE"
        dbcAutoriza.BoundColumn = "COD_USUARIO"

Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub dbcMotivo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdCancelar_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
            If Shift = 1 Then cmdAceptar_Click
    Case vbKeyEscape
            cmdCancelar_Click
End Select

End Sub

Private Sub Form_Load()
On Error GoTo handle
        txtImporte.Text = ""
        Me.Caption = strCodigo & " - " & strDescripcion
        Set dbcMotivo.RowSource = objAutorizacion.Lista("")
        dbcMotivo.ListField = "DES_AUTORIZACION"
        dbcMotivo.BoundColumn = "COD_AUTORIZACION"

Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Tipo_Click(Index As Integer)
On Error GoTo handle
    If Index = 0 Then
        txtImporte.Tipo = Porcentaje
        CalculaXPorcentaje
        TipoValor = Valor_Porcentaje
    End If
    If Index = 1 Then
        txtImporte.Tipo = Real
        CalculaXImporte
        TipoValor = Valor_Importe
    End If

Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub txtImporte_Change()
On Error GoTo handle
    If Tipo(0).Value Then CalculaXPorcentaje
    If Tipo(1).Value Then CalculaXImporte

Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub



Private Sub CalculaXPorcentaje()
On Error GoTo handle
    lblTotalcDscto.Caption = Round(Val(lblPrecioAct.Caption) * (1 - Val(txtImporte.Text) / 100), 2)
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub


Private Sub CalculaXImporte()
On Error GoTo handle
    lblTotalcDscto.Caption = Round(Val(lblPrecioAct.Caption) - Val(txtImporte.Text), 2)
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub
