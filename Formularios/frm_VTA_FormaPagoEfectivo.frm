VERSION 5.00
Begin VB.Form frm_VTA_FormaPagoEfectivo 
   BorderStyle     =   0  'None
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
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Moneda"
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   6615
      Begin vbp_Ventas.ctlTextBox txtPagaCon 
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         Tipo            =   4
         Alignment       =   1
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
         TabIndex        =   13
         Top             =   390
         Width           =   1575
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
         TabIndex        =   12
         Top             =   840
         Width           =   1575
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
         TabIndex        =   11
         Top             =   840
         Width           =   945
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
         TabIndex        =   10
         Top             =   1260
         Width           =   1335
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
         Top             =   1260
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5745
      Picture         =   "frm_VTA_FormaPagoEfectivo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4440
      Picture         =   "frm_VTA_FormaPagoEfectivo.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1095
   End
   Begin vbp_Ventas.ctlGrilla grdEfectivo 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3413
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione la Moneda    (F1)"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   2775
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
      TabIndex        =   6
      Top             =   6900
      Width           =   390
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
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_VTA_FormaPagoEfectivo.frx":0B14
      Top             =   60
      Width           =   240
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
      TabIndex        =   4
      Top             =   60
      Width           =   2670
   End
End
Attribute VB_Name = "frm_VTA_FormaPagoEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objFormaPago As New clsFormaPago
Dim strCodF As String
Dim dblImpTot As Double
Dim dblImpTotDol As String
Dim strTC As String
Dim strPagaCon As String
Public pblnOpc As Boolean
''nuevas variables
Public pstrDato As String
Public pstrDatoDes As String

Private Sub Form_Load()
    carga
End Sub
Public Sub carga(Optional ByVal Esshow As Boolean = False)
    If Esshow = True Then
        pstrDato = "001"
    pstrDatoDes = "EFECTIVO"
    
    

    End If
    
    'If pstrDato = "" Then MsgBox "ERROR"
    Me.top = 0
    Me.left = 0
    setteaFormulario Me
    SeteaGrilla
    If pblnOpc = False Then
        lblTipoCambio.Caption = objUsuario.TipoCambio
        strTC = objUsuario.TipoCambio
        Set grdEfectivo.DataSource = objFormaPago.ListaHijo(pstrDato)
    Else
        Set grdEfectivo.DataSource = objFormaPago.ListaHijo(frm_VTA_FormaPago.cCodFPadre)
    End If
    cmdAceptar.Default = False
    If Esshow = True Then
        Me.Show
        grdEfectivo.SetFocus
    End If
    
    If objVenta.CodModalidadVenta = Venta_Convenio Then
        txtPagaCon.Text = Format(frmPedido.lblcopago, "0.00")
    Else
        If Format(frmPedido.lblTotal, "0.00") <> "0.00" Then
            txtPagaCon.Text = Format(frmPedido.lblTotal, "0.00")
        End If
    End If
    
End Sub
Private Sub cmdAceptar_Click()
On Error GoTo Control
    'If (txtPagaCon.Text = "") Or (txtPagaCon.Text = "0") Then MsgBox "Ingrese el Monto a Pagar", vbExclamation, App.ProductName: txtPagaCon.SetFocus: Exit Sub
    If (txtPagaCon.Text = "" Or txtPagaCon.Text = "0") And objVenta.CodModalidadVenta <> Venta_Convenio Then MsgBox "Ingrese el Monto a Pagar", vbExclamation, App.ProductName: txtPagaCon.SetFocus: Exit Sub
    objVenta.AgregaFormaPago pstrDato, _
                             pstrDatoDes, _
                             strCodF, _
                             grdEfectivo.Columns("DES_HIJO").Value, _
                             dblImpTot, _
                             "", _
                             grdEfectivo.Columns("COD_MONEDA").Value, _
                             "", _
                             "", _
                             "", _
                             "", _
                             objUsuario.TipoCambio, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 0, "", "", dblImpTotDol
                             
    frmPedido.Cal_Promo
    Unload Me
    '***************************************'
    'Arma el arreglo cada vez que se modifica'
    frm_VTA_FormaPago.SetFocus
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
         Case vbKeyF1
            grdEfectivo.SetFocus
    End Select
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
End Sub

Private Sub grdEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub grdEfectivo_RegistroSeleccionado(ByVal DatoColumna0 As String)
    If grdEfectivo.ApproxCount <= 0 Then Exit Sub
    If pblnOpc = False Then
        strCodF = DatoColumna0
    Else
        strCodF = frm_VTA_FormaPago.cCodFHijo
    End If
    LblTituloImporte.Caption = "Importe S/. : "
    If pblnOpc = False Then
        txtPagaCon_Change
    End If
End Sub

Private Sub txtPagaCon_Change()
    If txtPagaCon.Text = "" Then lblImporte.Caption = "0.00": Exit Sub
    'If (txtPagaCon.Text = "" Or txtPagaCon.Text = "0") And objVenta.CodModalidadVenta <> Venta_Convenio Then Exit Sub
    strPagaCon = txtPagaCon.Text
    Select Case strCodF
           Case "001"
                dblImpTot = Round(Val(strPagaCon), 2)
                dblImpTotDol = "0.0"
                lblImporte.Caption = Format(dblImpTot, "#,###,##0.00")
                'lblTipoCambio.BackColor = RGB(242, 242, 242)
                'lblTipoCambio.ForeColor = RGB(0, 0, 0)
           Case "002"
                dblImpTot = Round(Val(strPagaCon) * Val(strTC), 2)
                dblImpTotDol = CStr(Format(Round(Val(strPagaCon), 2), "#,###,##0.00"))
                lblImporte.Caption = Format(dblImpTot, "#,###,##0.00")
                'lblTipoCambio.BackColor = RGB(227, 255, 213)
                'lblTipoCambio.ForeColor = RGB(0, 0, 0)
    End Select
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_HIJO", "DES_HIJO", "COD_MONEDA")
    arrCaption = Array("Codigo", "Efectivo", "Moneda")
    arrAncho = Array(900, 3000, 800)
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, vbAlignLeft)
    grdEfectivo.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    Dim i%
    For i = 0 To grdEfectivo.Columns.Count - 1
        grdEfectivo.Columns(i).Visible = False
    Next i
    grdEfectivo.Columns("COD_HIJO").Visible = False
    grdEfectivo.Columns("DES_HIJO").Visible = True
    grdEfectivo.Columns("COD_HIJO").AllowFocus = False
    grdEfectivo.Columns("DES_HIJO").ButtonText = True
    grdEfectivo.Styles(5).ForeColor = vbBlack
    grdEfectivo.Styles(5).Font.Bold = True
End Sub

Private Sub cmdCancelar_Click()
 Me.Cancelar
End Sub

Public Sub Cancelar()
    Unload Me
    'frmPedido.flgF6 = 0
    pblnOpc = False
    frmPedido.OptionsFocus
End Sub

Private Sub txtPagaCon_GotFocus()
    cmdAceptar.Default = True
End Sub

Private Sub txtPagaCon_LostFocus()
    cmdAceptar.Default = False
End Sub
