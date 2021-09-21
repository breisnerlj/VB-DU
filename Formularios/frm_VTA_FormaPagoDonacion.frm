VERSION 5.00
Begin VB.Form frm_VTA_FormaPagoDonacion 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4440
      Picture         =   "frm_VTA_FormaPagoDonacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5745
      Picture         =   "frm_VTA_FormaPagoDonacion.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   1095
   End
   Begin vbp_Ventas.ctlTextBox txtPagaCon 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
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
   Begin vbp_Ventas.ctlGrilla grdDonacion 
      Height          =   1935
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3413
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione la Institucion de la Donación (F1)"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   5175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   6660
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Forma de Pago - Donacion"
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
      Width           =   2790
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_VTA_FormaPagoDonacion.frx":0B14
      Top             =   60
      Width           =   240
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
      Left            =   360
      TabIndex        =   10
      Top             =   2790
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
      Left            =   360
      TabIndex        =   9
      Top             =   3240
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
      Left            =   2070
      TabIndex        =   8
      Top             =   3240
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
      Left            =   360
      TabIndex        =   7
      Top             =   3660
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
      Left            =   2070
      TabIndex        =   6
      Top             =   3660
      Width           =   975
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
      Left            =   6090
      TabIndex        =   4
      Top             =   6900
      Width           =   390
   End
End
Attribute VB_Name = "frm_VTA_FormaPagoDonacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objFormaPago As New clsFormaPago
Dim strCodF As String
Dim dblImpTot As Double
Dim strTC As String
Dim strPagaCon As String
Public pblnOpc As Boolean
''nuevas variables
Public pstrDato As String
Public pstrDatoDes As String


Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    SetteaFormulario Me
    SeteaGrilla
    If pblnOpc = False Then
        lblTipoCambio.Caption = Format(objUsuario.TipoCambio, "#,###,##0.00")
        strTC = objUsuario.TipoCambio
        Set grdDonacion.DataSource = objFormaPago.ListaHijo(pstrDato)
      Else
        Set grdDonacion.DataSource = objFormaPago.ListaHijo(frm_VTA_FormaPago.cCodFPadre)
    End If
    
    ' Redondeo de centimos
    If frmPedido.lblTotalPagar.Caption > 0 Then
        txtPagaCon.Text = Val("0.0" + Right(Format(frmPedido.lblTotalPagar.Caption, "##0.00"), 1))
    Else
        txtPagaCon.Text = -1 * Val("0.0" + Right(Format(frmPedido.lblTotalPagar.Caption, "##0.00"), 1))
    End If
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo Control

    If (txtPagaCon.Text = "") Or (txtPagaCon.Text = "0") Then MsgBox "Ingrese el Monto a Pagar", vbExclamation, "Ingrese Monto": txtPagaCon.SetFocus: Exit Sub
    objVenta.AgregaFormaPago pstrDato, _
                             pstrDatoDes, _
                             strCodF, _
                             grdDonacion.Columns("DES_HIJO").Value, _
                             -1 * dblImpTot, _
                             "", _
                             grdDonacion.Columns("COD_MONEDA").Value, _
                             "", _
                             "", _
                             "", _
                             "", _
                             objUsuario.TipoCambio
    frmPedido.Cal_Promo
    Unload Me
    '***************************************'
    'Arma el arreglo cada vez que se modifica'
      frm_VTA_FormaPago.SetFocus
      'frm_VTA_FormaPago.GrdListaFP.Array = objVenta.FormaPago
      frm_VTA_FormaPago.GrdListaFP.Rebind
    '***************************************'
    
        frmPedido.Cal_Montos
    
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

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
            grdDonacion.SetFocus
    End Select

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub grdDonacion_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub grdDonacion_RegistroSeleccionado(ByVal DatoColumna0 As String)
    If grdDonacion.ApproxCount <= 0 Then Exit Sub
    
    If pblnOpc = False Then
        strCodF = DatoColumna0
      Else
        strCodF = frm_VTA_FormaPago.cCodFHijo
    End If
    'Select Case strCodF
    '    Case "001"
    '        LblTituloImporte.Caption = "Importe S/. : "
    '    Case "002"
            LblTituloImporte.Caption = "Importe S/. : "
    'End Select
    If pblnOpc = False Then
        txtPagaCon_Change
    End If
End Sub

Private Sub txtPagaCon_Change()
    'If pblnOpc = False Then
        If txtPagaCon.Text = "" Then lblImporte.Caption = "0.00": Exit Sub
            strPagaCon = Val(txtPagaCon.Text)
            
            If grdDonacion.Columns("COD_MONEDA").Value = "1" Then
            
                        dblImpTot = strPagaCon
                        lblImporte.Caption = Format(strPagaCon, "#,###,##0.00")
                        lblTipoCambio.BackColor = RGB(242, 242, 242)
                        lblTipoCambio.ForeColor = RGB(0, 0, 0)
            Else
                        dblImpTot = (strPagaCon * strTC)
                        lblImporte.Caption = Format((strPagaCon * strTC), "#,###,##0.00")
                        lblTipoCambio.BackColor = RGB(227, 255, 213)
                        lblTipoCambio.ForeColor = RGB(0, 0, 0)
                
            End If
     'End If
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
    grdDonacion.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    Dim i%
    For i = 0 To grdDonacion.Columns.Count - 1
        grdDonacion.Columns(i).Visible = False
    Next i
    grdDonacion.Columns("COD_HIJO").Visible = True
    grdDonacion.Columns("DES_HIJO").Visible = True
    
    grdDonacion.Columns("COD_HIJO").AllowFocus = False
    grdDonacion.Columns("DES_HIJO").ButtonText = True
    'grdDonacion.RowHeight = 1.5 * grdDonacion.RowHeight
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
    pblnOpc = False
End Sub

Private Sub txtPagaCon_KeyPress(KeyAscii As Integer)
    txtPagaCon.Tipo = Real
End Sub


