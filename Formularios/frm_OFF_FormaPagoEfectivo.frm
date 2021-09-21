VERSION 5.00
Begin VB.Form frm_OFF_FormaPagoEfectivo 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4260
      Picture         =   "frm_OFF_FormaPagoEfectivo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5565
      Picture         =   "frm_OFF_FormaPagoEfectivo.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Moneda"
      Height          =   1755
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6615
      Begin vbp_Ventas.ctlTextBox txtPagaCon 
         Height          =   375
         Left            =   2220
         TabIndex        =   1
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Tipo            =   4
         Alignment       =   1
         MaxLength       =   10
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
      Begin VB.Label lblImporte 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
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
         Left            =   2310
         TabIndex        =   9
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label LblTituloImporte 
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
         Left            =   600
         TabIndex        =   8
         Top             =   1230
         Width           =   1335
      End
      Begin VB.Label lblTipoCambio 
         Alignment       =   1  'Right Justify
         Caption         =   "3.25"
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
         Left            =   2310
         TabIndex        =   7
         Top             =   810
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de cambio :"
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
         Left            =   600
         TabIndex        =   6
         Top             =   810
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Paga con : "
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
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
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
      Left            =   4200
      TabIndex        =   11
      Top             =   6780
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
      Left            =   5910
      TabIndex        =   10
      Top             =   6780
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Forma de Pago - Efectivo"
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
      TabIndex        =   0
      Top             =   60
      Width           =   2670
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_OFF_FormaPagoEfectivo.frx":0B14
      Top             =   60
      Width           =   240
   End
End
Attribute VB_Name = "frm_OFF_FormaPagoEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public strCodF As String
Dim strCodMoneda As String
Dim strPagaCon As String
Dim dblImpTot As Double
Dim dblImpTotDol As Double

Private Sub cmdAceptar_Click()

If Val(txtPagaCon.Text) <= 0 Then MsgBox "El monto debe de ser mayor a cero ", vbCritical, App.ProductName: txtPagaCon.SetFocus: Exit Sub
If txtPagaCon.Text = "" Then MsgBox "Debe de ingresar un monto", vbCritical, App.ProductName: txtPagaCon.SetFocus: Exit Sub
On Error GoTo CtrlErr

    objOFFVenta.AgregaPagoVenta strCodF, _
                        strCodMoneda, _
                        Val(strPagaCon), _
                        dblImpTot, "", "", 0, Val(lblTipoCambio.Caption), 0, ""

    Unload Me
    
    frm_OFF_FormaPago.grdListaFP.Rebind
    
    frm_OFF_Principal.MostrarTotales
    
Exit Sub

CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo CtrlErr
    
    Dim tmpCtrl As Boolean, tmpAlt As Boolean
    
    tmpCtrl = (Shift And vbCtrlMask) > 0
    tmpAlt = (Shift And vbAltMask) > 0
    
    Select Case KeyCode
        Case vbKeyReturn
            If Shift = 1 Then Call cmdAceptar_Click
        ''''    DESDE ACA COPIA
        Case tmpCtrl And vbKeyQ And False
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyM And False
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyE And False
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyD
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyC
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case vbKeyF5
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case vbKeyF6 And False
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case vbKeyF7
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case vbKeyF8 And False
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyX And False
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyF
            cmdAceptar_Click
            frm_OFF_FormaPago.Form_KeyDown KeyCode, Shift
    End Select

    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_Load()
    setteaFormulario Me
End Sub

Private Sub lblTipoCambio_Change()
    lblTipoCambio.Caption = Format(lblTipoCambio.Caption, "#,###,##0.00")
End Sub

Private Sub txtPagaCon_Change()
    If txtPagaCon.Text = "" Then lblImporte.Caption = "0.00": Exit Sub
    strPagaCon = txtPagaCon.Text
    Select Case strCodF
           Case "0"
                lblTipoCambio.Caption = Format(1, "#,###,##0.00")
                dblImpTot = Round(Val(strPagaCon), 2)
                dblImpTotDol = 0
                lblImporte.Caption = Format(dblImpTot, "#,###,##0.00")
                strCodMoneda = "1"
           Case "1"
                dblImpTot = Round(Val(strPagaCon) * Val(lblTipoCambio.Caption), 2)
                dblImpTotDol = Format(Round(Val(strPagaCon), 2))
                lblImporte.Caption = Format(dblImpTot, "#,###,##0.00")
                strCodMoneda = "2"
    End Select

End Sub
