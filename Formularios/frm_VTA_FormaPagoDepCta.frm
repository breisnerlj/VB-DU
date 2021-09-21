VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_VTA_FormaPagoDepCta 
   BorderStyle     =   0  'None
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   4440
      Picture         =   "frm_VTA_FormaPagoDepCta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
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
      Left            =   5745
      Picture         =   "frm_VTA_FormaPagoDepCta.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   1095
   End
   Begin vbp_Ventas.ctlDataCombo dbcboCuenta 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   2820
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin MSMask.MaskEdBox mskFecEmi 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   3300
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin vbp_Ventas.ctlTextBox ctlTextBox1 
      Height          =   315
      Left            =   3180
      TabIndex        =   5
      Top             =   4200
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbp_Ventas.ctlDataCombo ctlDataCombo2 
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   4200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlTextBox ctlTextBox2 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   3750
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbp_Ventas.ctlGrilla ctlGrilla1 
      Height          =   1935
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   6255
      _ExtentX        =   13256
      _ExtentY        =   3413
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
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
      Left            =   4380
      TabIndex        =   18
      Top             =   6900
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
      Left            =   6097
      TabIndex        =   17
      Top             =   6900
      Width           =   390
   End
   Begin VB.Label Label8 
      Caption         =   "Nro. transacción :"
      Height          =   255
      Left            =   420
      TabIndex        =   16
      Top             =   3780
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Importe S/. : "
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
      Left            =   420
      TabIndex        =   15
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Caption         =   "35.25"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "F. transacción :"
      Height          =   255
      Left            =   420
      TabIndex        =   13
      Top             =   3330
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Monto : "
      Height          =   255
      Left            =   420
      TabIndex        =   12
      Top             =   4230
      Width           =   1575
   End
   Begin VB.Label lblTipoCambio 
      Alignment       =   1  'Right Justify
      Caption         =   "3.25"
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Tipo de cambio :"
      Height          =   255
      Left            =   420
      TabIndex        =   10
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Cuenta : "
      Height          =   255
      Left            =   420
      TabIndex        =   9
      Top             =   2850
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_VTA_FormaPagoDepCta.frx":0B14
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Forma de Pago - Depósito en cuenta"
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
      Left            =   420
      TabIndex        =   8
      Top             =   60
      Width           =   3855
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   6255
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   60
      X2              =   6255
      Y1              =   2595
      Y2              =   2595
   End
End
Attribute VB_Name = "frm_VTA_FormaPagoDepCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''nuevas variables
Public pstrDato As String
Public pstrDatoDes As String

Private Sub cmdAceptar_Click()
On Error GoTo Control
'Validar y pasar a la grilla de forma de pago
cmdCancelar_Click
frmPedido.Cal_Promo
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control
    psub_KeyDownAplicacion KeyCode, Shift
    Select Case KeyCode
        Case vbKeyReturn
            If Shift = 1 Then cmdAceptar_Click
        Case vbKeyEscape
            Unload Me
    End Select

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
SetteaFormulario Me
End Sub

