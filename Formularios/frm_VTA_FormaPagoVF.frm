VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_VTA_FormaPagoVF 
   BorderStyle     =   0  'None
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4125
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox txtNumero 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   975
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">AAA-AAAAAAAAAA"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   3660
      Picture         =   "frm_VTA_FormaPagoVF.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   4980
      Picture         =   "frm_VTA_FormaPagoVF.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtValor 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin MSMask.MaskEdBox mskFec 
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
      TabIndex        =   1
      Top             =   1560
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frm_VTA_FormaPagoVF.frx":0B14
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Forma de Pago - Vale Fid."
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
      Left            =   720
      TabIndex        =   9
      Top             =   120
      Width           =   2685
   End
   Begin VB.Label Label4 
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
      Left            =   240
      TabIndex        =   8
      Top             =   2190
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha :"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1590
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Número :"
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
      Left            =   240
      TabIndex        =   6
      Top             =   990
      Width           =   1455
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
      Left            =   3570
      TabIndex        =   5
      Top             =   3540
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
      Left            =   5325
      TabIndex        =   4
      Top             =   3540
      Width           =   390
   End
End
Attribute VB_Name = "frm_VTA_FormaPagoVF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objFormaPago As New clsFormaPago
Dim objDocPago As New clsDocumentoPago
Dim odynR1 As oraDynaset
Dim odynr2 As oraDynaset
Dim strDato As String
Dim strDatoDes As String
Dim strMoneda As String
Dim dblImpTotal As Double
''nuevas variables
Public od As oraDynaset
Public pstrDato As String
Public pstrDatoDes As String


Private Sub Form_Load()
On Error GoTo handle
    Me.top = 0
    Me.left = 0
    setteaFormulario Me
    'SeteaGrilla
    'Set od = objDocPago.Lista(objUsuario.CodigoLocal)
    'Set grdDocDescuento.DataSource = od
    Set odynR1 = objFormaPago.ListaHijo(pstrDato)
    strDato = "" & odynR1("COD_HIJO").Value
    strDatoDes = "" & odynR1("DES_HIJO").Value
    strMoneda = "" & odynR1("COD_MONEDA").Value
    txtValor.Text = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "VALEFID")
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Public Sub aceptar()
'If validaDocumento = False Then Exit Sub

On Error GoTo Control
        Dim numero As String
        Dim Valida As String
        numero = Replace(TxtNumero.Text, "_", "")
        Valida = ValidaValeFid(numero)
        If Valida = "False" Then Exit Sub
        'If mskFecEmi.Mask = "##/##/####" Then MsgBox "Ingrese Fecha de emisión de documento", vbCritical, Caption: mskFecEmi.SetFocus: Exit Sub
        If IsDate(mskFec.Text) = False Then MsgBox "Ingrese Fecha de emisión de documento", vbCritical, App.ProductName: mskFec.SetFocus: Exit Sub
        'If txtNombre.Text = "" Then MsgBox "Ingrese el nombre del cliente", vbCritical, App.ProductName: txtNombre.SetFocus: Exit Sub
        'If txtDNI.Text = "" Then MsgBox "Ingrese el dni del cliente", vbCritical, App.ProductName: txtDNI.SetFocus: Exit Sub
        If txtValor.Text = "" Then MsgBox "Ingrese el Importe del documento", vbCritical, App.ProductName: txtValor.SetFocus: Exit Sub
        'MsgBox Mid(txtNumero.Text, 20, 22)
        'MsgBox Len(Mid(txtNumero.Text, 20, 22))
        ' Mid(numero, 20, 22),
        objVenta.AgregaFormaPago pstrDato, _
                                 pstrDatoDes, _
                                 strDato, _
                                 strDatoDes, _
                                 txtValor.Text, _
                                 "", IIf(strMoneda = "", "1", strMoneda), _
                                 "", "", _
                                 "", "", _
                                 0, numero, _
                                 "", "", _
                                 "", "", _
                                 "", "", _
                                 "", "", _
                                 mskFec.Text, "", _
                                 "", "", _
                                 "", strDato, _
                                 "", "", 0, "", "0", "", "", _
                                 Valida

                                 frmPedido.Cal_Promo
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
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Function ValidaValeFid(num As String) As String
    If objUsuario.CodLocalCallCenter <> "1DLV" Then MsgBox "Vale Fid. no aplica para este Call Center.", vbCritical, App.ProductName: ValidaValeFid = False: Exit Function
    If num = "--" Then MsgBox "Ingresa el numero de Vale Fid.", vbCritical, App.ProductName: ValidaValeFid = False: Exit Function
    If objUsuario.CodLocalCallCenter = "1DLV" Then 'INKAFARMA
        Dim oFP As New clsFarmaPuntos
        ValidaValeFid = oFP.ValidaValeFidInka(num)
    End If
End Function

Private Sub txtCriterio_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Or KeyAscii = 9 Then gclsOracle.FN_Cursor("ECVENTA.UNIFICADO_INKACLUB.CONSULTA_CLI",0,) KeyAsciii = 0
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then SendKeys "{TAB}": KeyAscii = 0
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

Private Sub cmdAceptar_Click()
    aceptar
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub


