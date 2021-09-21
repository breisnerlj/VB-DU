VERSION 5.00
Begin VB.Form frm_VTA_FormaPagoNC 
   BorderStyle     =   0  'None
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6960
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
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Nota de Credito"
      Height          =   2655
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   5415
      Begin VB.Label Label9 
         Caption         =   "N. Documento"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label lblDocumento 
         BackColor       =   &H00DBFBFA&
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Local : "
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label lblLocal 
         BackColor       =   &H00DBFBFA&
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha : "
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   1260
         Width           =   1515
      End
      Begin VB.Label lblFecha 
         BackColor       =   &H00DBFBFA&
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   1260
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Importe S/. :"
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
         TabIndex        =   12
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label lblImporte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00DBFBFA&
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
         Left            =   2160
         TabIndex        =   11
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Referencia : "
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   2100
         Width           =   1515
      End
      Begin VB.Label lblReferencia 
         BackColor       =   &H00DBFBFA&
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   2100
         Width           =   2295
      End
   End
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
      Picture         =   "frm_VTA_FormaPagoNC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
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
      Picture         =   "frm_VTA_FormaPagoNC.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1095
   End
   Begin vbp_Ventas.ctlTextBox txtNroNC 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      Tipo            =   7
      MaxLength       =   11
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese el Numero de la Nota de Crédito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Width           =   2835
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_VTA_FormaPagoNC.frx":0B14
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Forma de Pago - Nota  de Credito"
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
      TabIndex        =   6
      Top             =   60
      Width           =   3525
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   6900
      Width           =   390
   End
   Begin VB.Label Label2 
      Caption         =   "Nota de Crédito : "
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   990
      Width           =   1635
   End
End
Attribute VB_Name = "frm_VTA_FormaPagoNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objDocumento As New clsDocumento
Dim objFormaPago As New clsFormaPago
Dim odynR1 As oraDynaset
Dim odynr2 As oraDynaset
Dim strMoneda As String
Dim flgKardex As String
''nuevas variables
Public pstrDato As String
Public pstrDatoDes As String

Private Sub Form_Load()
On Error GoTo handle
    Me.top = 0
    Me.left = 0
    SetteaFormulario Me
    Set odynr2 = objFormaPago.ListaHijo(pstrDato)
    strMoneda = "" & odynr2("COD_MONEDA").Value
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo Control
        
        If flgKardex = "0" Then MsgBox "No se puede usar una NC por DESCUENTO como forma de pago", vbCritical, App.ProductName: txtNroNC.SetFocus: Exit Sub
        If lblImporte.Caption = "" Then MsgBox "Importe Nota credito tiene que ser mayor a cero", vbCritical, App.ProductName: txtNroNC.SetFocus: Exit Sub
        If (lblImporte.Caption = "") Or (lblImporte.Caption = "0.00") Then MsgBox "Nota de Credito no existe", vbCritical, App.ProductName: txtNroNC.selection: Exit Sub
        
        'Comentado por Jahzeel para sacar la fecha y el documento
        'objVenta.AgregaFormaPago pstrDato, _
                                 pstrDatoDes, _
                                 pstrDato, _
                                 pstrDatoDes, _
                                 lblImporte.Caption, _
                                 "", strMoneda, "", "", _
                                 "", "", 0, "", _
                                 "", "", "", "", _
                                 "", "", "", "", _
                                 lblFecha.Caption, _
                                 lblDocumento.Caption, "", "", "", _
                                 "", lblLocal.Caption, lblReferencia.Caption

        

        objVenta.AgregaFormaPago pstrDato, _
                                 pstrDatoDes, _
                                 odynr2("COD_HIJO").Value, _
                                 odynr2("DES_HIJO").Value, _
                                 lblImporte.Caption, _
                                 "", strMoneda, "", "", _
                                 "", "", 0, "", _
                                 "", "", "", "", _
                                 "", "", "", "", _
                                 Format(lblFecha.Caption, "DD/MM/YYYY"), Trim(Replace(txtNroNC.Text, "-", "")), "", "", "", _
                                 "", lblLocal.Caption, lblReferencia.Caption


frmPedido.Cal_Promo
    Unload Me
    '***************************************'
    'Arma el arreglo cada ez que se modifica'
      frm_VTA_FormaPago.Show
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

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

'Private Sub txtNroNC_Change()
'On Error GoTo handle
'    ''txtNroNC.Tipo = Entero
'    Exit Sub
'handle:
'    MsgBox Err.Description, vbCritical, App.ProductName
'
'End Sub

Private Sub txtNroNC_KeyPress(KeyAscii As Integer)
On Error GoTo handle
    If KeyAscii = 13 Then
         Set odynR1 = objDocumento.ListaNotaCredito(Replace(txtNroNC.Text, "-", ""))
         If odynR1.RecordCount <= 0 Then MsgBox "Nota de Credito no existe", vbCritical, Caption: txtNroNC.SetFocus: txtNroNC.selection: Exit Sub
         lblDocumento.Caption = txtNroNC.Text
         lblLocal.Caption = odynR1("COD_LOCAL").Value
         lblFecha.Caption = odynR1("FCH_REGISTRA").Value
         lblImporte.Caption = odynR1("MTO_TOTAL").Value
         lblReferencia.Caption = odynR1("NUM_DOCUMENTO_REFER").Value
         flgKardex = odynR1("FLG_KARDEX").Value
    End If
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub
