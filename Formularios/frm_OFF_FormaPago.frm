VERSION 5.00
Begin VB.Form frm_OFF_FormaPago 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlGrillaArray grdListaFP 
      Height          =   2355
      Left            =   120
      TabIndex        =   2
      Top             =   3660
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4154
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlGrillaArray grdFormaPago 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4154
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4260
      Picture         =   "frm_OFF_FormaPago.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5565
      Picture         =   "frm_OFF_FormaPago.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   1095
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   6780
      Width           =   390
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   3420
      Width           =   180
   End
   Begin VB.Label Label2 
      Caption         =   "Formas de pago registradas : "
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Line Line2 
      X1              =   180
      X2              =   6840
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   180
      X2              =   6840
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   180
   End
   Begin VB.Label lblDocumento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "frm_OFF_FormaPago.frx":0B14
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Forma de Pago"
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
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   1590
   End
End
Attribute VB_Name = "frm_OFF_FormaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim strFormaPago As String

Private Sub SetGrid()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim columna As TrueDBGrid70.Column
  
    
    arrCampos = Array("", "")
    arrCaption = Array("Código", "Descripción")
    arrAncho = Array(900, 4500)
    arrAlineacion = Array(dbgCenter, dbgLeft)
    
    
    grdFormaPago.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    For Each columna In grdFormaPago.Columns
        columna.AllowSizing = False
        columna.Visible = False
    Next
    
    grdFormaPago.Columns(1).Visible = True
    
    grdFormaPago.MarqueeStyle = dbgHighlightRow
'    grdFormaPago.CambiaSeleccionadoBackColor &H800000
'    grdFormaPago.CambiaSeleccionadoForeColor &HFFFFFF
    
    
    grdFormaPago.Columns(1).ButtonText = True
    grdFormaPago.StylesFondo = True
    grdFormaPago.CambiaSeleccionadoForeColor (vbBlack)
        
    
    
    grdFormaPago.Columns(0).AllowFocus = False
    
    
    
    
    arrCampos = Array("", "", "", "", "", "", "", "", "", "")
    arrCaption = Array("Código", "Forma Pago", "Moneda", "Pago con", "Total S/.", "Vuelto", "Tarjeta", "Vencimiento", "# Cuotas", "T.Cambio S/.")
    arrAncho = Array(0, 3000, 0, 1100, 1000, 0, 0, 0, 0, 900)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgRight, dbgRight, dbgRight, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgRight)
    grdListaFP.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    For Each columna In grdListaFP.Columns
        columna.AllowSizing = False
        columna.Visible = False
    Next
    

    grdListaFP.Columns(objOFFVenta.ColDPDescripcion).Visible = True
    grdListaFP.Columns(objOFFVenta.ColDPMtoSoles).Visible = True
    grdListaFP.Columns(objOFFVenta.ColDPMtoSoles).NumberFormat = "#,###,##0.00"
    
    grdListaFP.MarqueeStyle = dbgHighlightRow
    grdListaFP.CambiaSeleccionadoBackColor &H800000
    grdListaFP.CambiaSeleccionadoForeColor &HFFFFFF
    
End Sub

Private Sub CargaFormaPago()
Dim objFormaPago As cls_OFF_FormaPago

On Error GoTo CtrlErr

    Set objFormaPago = New cls_OFF_FormaPago
    grdFormaPago.Array1 = objFormaPago.FormaPago
    grdFormaPago.Rebind
    Set objFormaPago = Nothing

Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
'    objOFFVenta.LimpiaPago
'    grdListaFP.Limpiar
'    frm_OFF_Principal.MostrarTotales
    Unload Me
End Sub

Private Sub Form_Activate()
    grdListaFP.Array1 = objOFFVenta.PagoVenta
    grdListaFP.Rebind
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo CtrlErr
    
    Dim tmpCtrl As Boolean, tmpAlt As Boolean
    
    tmpCtrl = (Shift And vbCtrlMask) > 0
    tmpAlt = (Shift And vbAltMask) > 0

    Select Case KeyCode
        Case vbKeyEscape
            cmdCancelar_Click
        Case vbKeyF1
            grdFormaPago.SetFocus
        Case vbKeyF2
            grdListaFP.SetFocus
        Case vbKeyReturn
            If Shift = 1 Then cmdAceptar_Click
        ''''    DESDE ACA COPIA
        Case tmpCtrl And vbKeyQ And False
            cmdAceptar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyM And False
            cmdAceptar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyE And False
            cmdAceptar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyD
            cmdAceptar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyC
            cmdAceptar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case vbKeyF5
            cmdAceptar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case vbKeyF6 And False
            cmdAceptar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case vbKeyF7
            cmdAceptar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case vbKeyF8 And False
            cmdAceptar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyX And False
            cmdAceptar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case tmpCtrl And vbKeyF
            cmdAceptar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
    End Select
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_Load()

    setteaFormulario Me
    
    Call SetGrid
    Call CargaFormaPago

End Sub

Private Sub grdFormaPago_ButtonClick(ByVal ColIndex As Integer)
    Select Case ColIndex
        Case 1
            grdFormaPago_DblClick
    End Select

End Sub

Private Sub grdFormaPago_DblClick()
On Error GoTo CtrlErr

    If grdFormaPago.ApproxCount = 0 Then Exit Sub
    
    frm_OFF_FormaPagoEfectivo.lblTipoCambio = objOFFUsuario.TipoCambio
    
    
    Select Case Trim(strFormaPago)
        Case "0"
            frm_OFF_FormaPagoEfectivo.strCodF = Trim(strFormaPago)
            frm_OFF_FormaPagoEfectivo.Label1(4).Caption = "Forma de Pago - Efectivo Soles"
            frm_OFF_FormaPagoEfectivo.Show vbModal
        Case "1"
            frm_OFF_FormaPagoEfectivo.strCodF = Trim(strFormaPago)
            frm_OFF_FormaPagoEfectivo.Label1(4).Caption = "Forma de Pago - Efectivo Dolares"
            frm_OFF_FormaPagoEfectivo.Show vbModal
        Case "2"
            frm_OFF_FormaPagoTarjeta.strCodF = Trim(strFormaPago)
            frm_OFF_FormaPagoTarjeta.Show vbModal
    End Select

Exit Sub

CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub grdFormaPago_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error GoTo CtrlErr

    If grdFormaPago.ApproxCount = 0 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyReturn
            Call grdFormaPago_DblClick
            
        
    End Select
    

Exit Sub
    
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub grdFormaPago_RegistroSeleccionado(ByVal DatoColumna0 As String)
    strFormaPago = DatoColumna0
End Sub

Private Sub grdListaFP_DblClick()
    ModificarFP
End Sub

Private Sub grdListaFP_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn
            ModificarFP
        Case vbKeyDelete
            On Error GoTo CtrlErr
            If grdListaFP.ApproxCount = 0 Then Exit Sub
            grdListaFP.Delete
            frm_OFF_Principal.MostrarTotales
CtrlErr:
            On Error GoTo 0
    
    End Select

End Sub

Sub ModificarFP()
    Dim row As Integer
    
    On Error GoTo CtlrErr
            
    If grdListaFP.ApproxCount = 0 Then Exit Sub

    row = grdListaFP.row
    
    Select Case Val(objOFFVenta.PagoVenta.Value(row, objOFFVenta.ColDPCodPago))
        Case 0, 1 'EFECTIVO O DOLARES
            With frm_OFF_FormaPagoEfectivo
                .strCodF = objOFFVenta.PagoVenta.Value(row, objOFFVenta.ColDPCodPago)
                .txtPagaCon.Text = objOFFVenta.PagoVenta.Value(row, objOFFVenta.ColDPMtoImporte)
                .lblImporte.Caption = objOFFVenta.PagoVenta.Value(row, objOFFVenta.ColDPMtoSoles)
                .lblTipoCambio = objOFFVenta.PagoVenta.Value(row, objOFFVenta.ColDPTipoCambio)
                .Show vbModal
            End With
        Case 2 'TARJETAS
            With frm_OFF_FormaPagoTarjeta
                .strCodF = objOFFVenta.PagoVenta.Value(row, objOFFVenta.ColDPCodPago)
                .txtNroTar.Text = objOFFVenta.PagoVenta.Value(row, objOFFVenta.ColDPNumTarjeta)
                .mskVencimiento.Text = objOFFVenta.PagoVenta.Value(row, objOFFVenta.ColDPFchVencimiento)
                .txtNroCuota.Text = objOFFVenta.PagoVenta.Value(row, objOFFVenta.ColDPNumCuotas)
                .cboTipoCuota.ListIndex = Val(objOFFVenta.PagoVenta.Value(row, objOFFVenta.ColDPTipoCuota)) - 1
                .txtNumAutorizacion.Text = objOFFVenta.PagoVenta.Value(row, objOFFVenta.ColDPNumAutorizacion)
                .txtImporte.Text = objOFFVenta.PagoVenta.Value(row, objOFFVenta.ColDPMtoImporte)
                .Show vbModal
            End With
    End Select
                
    Exit Sub
CtlrErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

