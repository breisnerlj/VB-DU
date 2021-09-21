VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_VTA_FormaPagoCheque 
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
      Caption         =   "Datos del Cheque"
      Height          =   3135
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   6615
      Begin vbp_Ventas.ctlTextBox txtNroChq 
         Height          =   315
         Left            =   2100
         TabIndex        =   1
         Top             =   240
         Width           =   2175
         _extentx        =   3836
         _extenty        =   556
         maxlength       =   18
         font            =   "frm_VTA_FormaPagoCheque.frx":0000
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
         Left            =   2100
         TabIndex        =   2
         Top             =   780
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
      Begin vbp_Ventas.ctlTextBox TxtMonto 
         Height          =   315
         Left            =   4020
         TabIndex        =   4
         Top             =   1260
         Width           =   1035
         _extentx        =   1826
         _extenty        =   556
         alignment       =   1
         font            =   "frm_VTA_FormaPagoCheque.frx":002C
      End
      Begin vbp_Ventas.ctlDataCombo ctlCboMoneda 
         Height          =   315
         Left            =   2100
         TabIndex        =   3
         Top             =   1260
         Width           =   1815
         _extentx        =   3201
         _extenty        =   556
         matchentry      =   1
      End
      Begin vbp_Ventas.ctlTextBox TxtObs 
         Height          =   375
         Left            =   2100
         TabIndex        =   5
         Top             =   2640
         Width           =   4215
         _extentx        =   8070
         _extenty        =   556
         font            =   "frm_VTA_FormaPagoCheque.frx":0058
      End
      Begin VB.Label Label2 
         Caption         =   "Nro. de cheque :"
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
         Left            =   480
         TabIndex        =   19
         Top             =   270
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
         Left            =   480
         TabIndex        =   18
         Top             =   2220
         Width           =   1575
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
         Left            =   2100
         TabIndex        =   17
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha  Emisión :"
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
         Left            =   480
         TabIndex        =   16
         Top             =   810
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Monto : "
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
         Left            =   480
         TabIndex        =   15
         Top             =   1290
         Width           =   1575
      End
      Begin VB.Label lblTipoCambio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00DBFBFA&
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
         Height          =   255
         Left            =   2100
         TabIndex        =   14
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label7 
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
         Left            =   480
         TabIndex        =   13
         Top             =   1740
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Observaciones :"
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
         Left            =   480
         TabIndex        =   12
         Top             =   2670
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4440
      Picture         =   "frm_VTA_FormaPagoCheque.frx":0084
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5760
      Picture         =   "frm_VTA_FormaPagoCheque.frx":060E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   1095
   End
   Begin vbp_Ventas.ctlGrilla grdBanco 
      Height          =   1575
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   6735
      _extentx        =   11880
      _extenty        =   2778
      menupopup       =   0   'False
      resalte         =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccion el Banco emisor del Cheque (F1)"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   3975
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
      TabIndex        =   10
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
      Left            =   6112
      TabIndex        =   9
      Top             =   6900
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_VTA_FormaPagoCheque.frx":0B98
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Forma de Pago - Cheque"
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
      Width           =   2625
   End
End
Attribute VB_Name = "frm_VTA_FormaPagoCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objBanco As New clsBanco
Dim objFormaPago As New clsFormaPago
Dim odynR1 As oraDynaset
Dim dblMto As Double
Dim dblImporte As Double
Dim dblTC As Double
Dim strDato As String
Dim strDatoDes As String
''nuevas variables
Public pstrDato As String
Public pstrDatoDes As String

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    SetteaFormulario Me
    Set grdBanco.DataSource = objBanco.Lista
    SeteaGrilla
    
    lblTipoCambio.Caption = objUsuario.TipoCambio
    
    Set ctlCboMoneda.RowSource = objFormaPago.ListaMoneda
    ctlCboMoneda.ListField = "DES"
    ctlCboMoneda.BoundColumn = "COD"
    ctlCboMoneda.BoundText = "1"
    
End Sub

Private Sub cmdAceptar_Click()
Dim strVencDia As String
Dim strVencMes As String
Dim strVencAño As String
On Error GoTo Control
    If txtNroChq.Text = "" Then MsgBox "Ingrese el numero del cheque", vbCritical, Caption: txtNroChq.selection: Exit Sub
    If Val(TxtMonto.Text) = 0 Then MsgBox "Ingrese el monto del cheque", vbCritical, App.ProductName: TxtMonto.SetFocus: Exit Sub
        
    strVencDia = Mid(mskFecEmi.Text, 1, 2)
    strVencMes = Mid(mskFecEmi.Text, 4, 2)
    strVencAño = Mid(mskFecEmi.Text, 7, 4)
    
    If strVencDia > 30 Then
        MsgBox "El dia ingresado no es valido", vbCritical, Caption
        mskFecEmi.SetFocus
        Exit Sub
    End If
    
    If strVencMes > 12 Then
        MsgBox "El mes ingresado no es valido", vbCritical, Caption
        mskFecEmi.SetFocus
        Exit Sub
    End If
    
    If (strVencAño < Format(objUsuario.sysdate, "yyyy")) Then
        MsgBox "El cheque Ingresada esta vencida", vbCritical, Caption
        mskFecEmi.SetFocus
        Exit Sub
    End If
        
        objVenta.AgregaFormaPago pstrDato, _
                                 pstrDatoDes, _
                                 strDato, _
                                 strDatoDes, _
                                 dblImporte, _
                                 "", _
                                 ctlCboMoneda.BoundText, _
                                 "", _
                                 grdBanco.Columns(0).Value, _
                                 "", "", _
                                 objUsuario.TipoCambio, _
                                 0, "", _
                                 "", "", _
                                 "", "", _
                                 "", "", _
                                 "", mskFecEmi.Text, _
                                 "", txtNroChq.Text, "", _
                                 "", "", _
                                 "", "", _
                                 "0", Trim(txtObs.Text)
                                 
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
            grdBanco.SetFocus
    End Select

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_BANCO", "DES_BANCO")
    arrCaption = Array("Codigo", "Banco")
    arrAncho = Array(900, 3000)
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft)
    grdBanco.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    Dim i%
    For i = 0 To grdBanco.Columns.Count - 1
        grdBanco.Columns(i).Visible = False
    Next i
    'grdBanco.Columns("COD_BANCO").Visible = False
    grdBanco.Columns("DES_BANCO").Visible = True
    'grdBanco.RowHeight = 1.5 * grdBanco.RowHeight
End Sub

Private Sub ctlCboMoneda_Change()
    If ctlCboMoneda.BoundText = "*" Then Exit Sub
    txtMonto_Change
End Sub

Private Sub mskFecEmi_Validate(Cancel As Boolean)
'    Cancel = Not fbln_Valida_Fecha("DD/MM/yyyy", "Error en el Ingreso de fechas", mskFecEmi.Text, mskFecEmi.Text)
'    If Cancel Then
'        MsgBox "Error en el Ingreso de fechas", vbExclamation, Caption
'        mskFecEmi.SetFocus
'    End If
    
    
    
    
End Sub

Private Sub txtMonto_Change()
    Dim vMon$
    vMon = ctlCboMoneda.BoundText
    If vMon = "1" Then
         lblTipoCambio.Caption = "0"
         dblTC = "0"
         dblMto = Val(TxtMonto.Text)
         dblImporte = dblMto
         lblImporte.Caption = dblImporte
         strDato = "4"
         strDatoDes = "SOLES"
       Else
         lblTipoCambio.Caption = objUsuario.TipoCambio
         dblTC = objUsuario.TipoCambio
         dblMto = Val(TxtMonto.Text)
         dblImporte = (dblMto * dblTC)
         lblImporte.Caption = dblImporte
         strDato = "5"
         strDatoDes = "DOLARES"
    End If
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
    TxtMonto.Tipo = Real
End Sub

Private Sub txtNroChq_KeyPress(KeyAscii As Integer)
    txtNroChq.Tipo = Entero
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub
