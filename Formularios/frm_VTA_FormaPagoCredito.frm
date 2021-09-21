VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_VTA_FormaPagoCredito 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox mskFecEmi 
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4440
      Picture         =   "frm_VTA_FormaPagoCredito.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5745
      Picture         =   "frm_VTA_FormaPagoCredito.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   1095
   End
   Begin vbp_Ventas.ctlTextBox txtNroDoc 
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
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
   End
   Begin vbp_Ventas.ctlGrilla grdDocDescuento 
      Height          =   1935
      Left            =   180
      TabIndex        =   0
      Top             =   780
      Width           =   6255
      _ExtentX        =   13256
      _ExtentY        =   3413
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox txtNombre 
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Top             =   4020
      Visible         =   0   'False
      Width           =   4155
      _ExtentX        =   9657
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
   Begin vbp_Ventas.ctlTextBox txtDNI 
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Top             =   4500
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      MaxLength       =   8
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
   Begin vbp_Ventas.ctlTextBox txtValor 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   4980
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      Alignment       =   1
      MaxLength       =   10
      TABAuto         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   " (F1)"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   180
      TabIndex        =   16
      Top             =   540
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Valor S/. :"
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
      Left            =   480
      TabIndex        =   14
      Top             =   5010
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha  Emisión :"
      Height          =   255
      Left            =   540
      TabIndex        =   13
      Top             =   3570
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Número : "
      Height          =   255
      Left            =   540
      TabIndex        =   12
      Top             =   3030
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_VTA_FormaPagoCredito.frx":0B14
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Forma de Pago - Venta al Crédito"
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
      TabIndex        =   11
      Top             =   60
      Width           =   3480
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   6840
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   60
      X2              =   6840
      Y1              =   2895
      Y2              =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Nombre : "
      Height          =   255
      Left            =   540
      TabIndex        =   10
      Top             =   4050
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "DNI : "
      Height          =   255
      Left            =   540
      TabIndex        =   9
      Top             =   4530
      Visible         =   0   'False
      Width           =   1575
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   6900
      Width           =   390
   End
End
Attribute VB_Name = "frm_VTA_FormaPagoCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim objFormaPago As New clsFormaPago
Dim odynR1 As oraDynaset
Dim odynr2 As oraDynaset
Dim strDato As String
Dim strDatoDes As String
Dim strMoneda As String
Dim dblImpTotal As Double
''nuevas variables
Public pstrDato As String
Public pstrDatoDes As String

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    SetteaFormulario Me
    SeteaGrilla
    
    'Set grdDocDescuento.DataSource = objDocPago.Lista
    Set odynR1 = objFormaPago.ListaHijo(pstrDato)
    Set grdDocDescuento.DataSource = odynR1
    strDato = "" & odynR1("COD_HIJO").Value
    strDatoDes = "" & odynR1("DES_HIJO").Value
    strMoneda = "" & odynR1("COD_MONEDA").Value
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo Control
            
'        If txtNroDoc.Text = "" Then MsgBox "Ingresa el numero de descuento", vbCritical, Caption: txtNroDoc.SetFocus: Exit Sub
'        If txtNombre.Text = "" Then MsgBox "Ingrese el nombre del cliente", vbCritical, Caption: txtNombre.SetFocus: Exit Sub
'        If txtDNI.Text = "" Then MsgBox "Ingrese el dni del cliente", vbCritical, Caption: txtDNI.SetFocus: Exit Sub
        If txtValor.Text = "" Then MsgBox "Ingrese el Importe del documento", vbCritical, App.ProductName: txtValor.SetFocus: Exit Sub
        
        objVenta.AgregaFormaPago pstrDato, _
                                 pstrDatoDes, _
                                 strDato, _
                                 strDatoDes, _
                                 dblImpTotal, _
                                 "", strMoneda, _
                                 "", "", _
                                 "", "", _
                                 0, "", _
                                 "", "", _
                                 "", txtNroDoc.Text, _
                                 "", "", _
                                 "", "", _
                                 mskFecEmi.Text, "", _
                                 "", txtNombre.Text, _
                                 txtDNI.Text, grdDocDescuento.Columns(0).Value
If objVenta.CodigoTipoVenta <> Cobro_Responsabilidad Then
    frmPedido.Cal_Promo
End If
                                 
    Unload Me
    '***************************************'
    'Arma el arreglo cada ez que se modifica'
      frm_VTA_FormaPago.SetFocus
      'frm_VTA_FormaPago.GrdListaFP.Array = objVenta.FormaPago
      frm_VTA_FormaPago.GrdListaFP.Rebind
    '***************************************'
    frmPedido.Cal_Montos
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control
    psub_KeyDownAplicacion KeyCode, Shift
    Select Case KeyCode
        Case vbKeyReturn
            If Shift = 1 Then cmdAceptar_Click
        Case vbKeyEscape
            cmdCancelar_Click
    End Select

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_HIJO", "DES_HIJO")
    arrCaption = Array("Codigo", "Vale")
    arrAncho = Array(900, 4500)
    arrAlineacion = Array(vbCenter, vbAlignLeft)
    grdDocDescuento.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    Dim i%
    For i = 0 To grdDocDescuento.Columns.Count - 1
        grdDocDescuento.Columns(i).Visible = False
    Next i
    grdDocDescuento.Columns("COD_HIJO").Visible = True
    grdDocDescuento.Columns("DES_HIJO").Visible = True
    grdDocDescuento.Columns(1).WrapText = True
    'grdDocDescuento.RowHeight = 1.5 * grdDocDescuento.RowHeight
    
End Sub

Private Sub grdDocDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

'Private Sub mskFecEmi_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Or KeyAscii = 9 Then SendKeys "{TAB}": KeyAscii = 0
'End Sub
'
'Private Sub mskFecEmi_Validate(Cancel As Boolean)
'    Cancel = Not fbln_Valida_Fecha("MM/yyyy", "Error en el Ingreso de fechas", mskFecEmi.Text)
'    If Cancel Then
'        MsgBox "Error en el Ingreso de fechas", vbExclamation, Caption
'        mskFecEmi.SetFocus
'    End If
'End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    txtDNI.Tipo = Entero
    If KeyAscii = 13 Or KeyAscii = 9 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    txtNombre.Tipo = Mayusculas
    If KeyAscii = 13 Or KeyAscii = 9 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
    txtNroDoc.Tipo = Entero
    If KeyAscii = 13 Or KeyAscii = 9 Then
      gclsOracle.Num_Intentos = 1
      'txtNroDoc.Text = objDocPago.validavale(txtNroDoc.Text)' comentado porke no existe el objeto
    End If
End Sub

Private Sub txtValor_Change()
    dblImpTotal = Val(txtValor.Text)
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    txtValor.Tipo = Real
    If KeyAscii = 13 Or KeyAscii = 9 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


