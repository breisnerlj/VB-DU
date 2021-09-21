VERSION 5.00
Begin VB.Form frm_VTA_ServicioSeguros 
   BorderStyle     =   0  'None
   Caption         =   "s"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5745
      Picture         =   "frm_VTA_ServicioSeguros.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4440
      Picture         =   "frm_VTA_ServicioSeguros.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1095
   End
   Begin vbp_Ventas.ctlTextBox txtSuministro 
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   3360
      Width           =   2775
      _ExtentX        =   4895
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
   Begin vbp_Ventas.ctlTextBox txtTC 
      Height          =   315
      Left            =   2340
      TabIndex        =   3
      Top             =   4320
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      ColorDefault    =   -2147483633
      ColorDefault    =   -2147483633
      Enabled         =   0   'False
      TABAuto         =   0   'False
      Bloqueado       =   -1  'True
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
   Begin vbp_Ventas.ctlGrilla grdDocumentos 
      Height          =   2775
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4895
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox txtMonto 
      Height          =   315
      Left            =   3960
      TabIndex        =   5
      Top             =   3840
      Width           =   1095
      _ExtentX        =   1931
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
   Begin vbp_Ventas.ctlTextBox txtImporte 
      Height          =   315
      Left            =   2340
      TabIndex        =   6
      Top             =   4800
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      ColorDefault    =   -2147483633
      ColorDefault    =   -2147483633
      Enabled         =   0   'False
      TABAuto         =   0   'False
      Bloqueado       =   -1  'True
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
   Begin vbp_Ventas.ctlDataCombo dbcMoneda 
      Height          =   315
      Left            =   2280
      TabIndex        =   15
      Top             =   3840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Servicios - Seguros"
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
      TabIndex        =   14
      Top             =   60
      Width           =   2100
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_VTA_ServicioSeguros.frx":0B14
      Top             =   60
      Width           =   240
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Numero de Operación:"
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
      Left            =   180
      TabIndex        =   13
      Top             =   3390
      Width           =   2040
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "T.C.:"
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
      Left            =   1815
      TabIndex        =   12
      Top             =   4350
      Width           =   405
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Monto:"
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
      Left            =   1620
      TabIndex        =   11
      Top             =   3870
      Width           =   600
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Importe:"
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
      Index           =   4
      Left            =   1080
      TabIndex        =   10
      Top             =   4830
      Width           =   720
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
      Index           =   5
      Left            =   6090
      TabIndex        =   9
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
      Index           =   6
      Left            =   4380
      TabIndex        =   8
      Top             =   6900
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "S/."
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
      Index           =   6
      Left            =   1980
      TabIndex        =   7
      Top             =   4830
      Width           =   240
   End
End
Attribute VB_Name = "frm_VTA_ServicioSeguros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objServicio As New clsServicio
Dim ObjFormaPago As New clsFormaPago
Public strCodigoPadre As String
Public strDescripcionPadre As String


Private Sub cmdAceptar_Click()
On Error GoTo handle

    If Val(txtMonto.Text) <= 0 Then
        MsgBox "Ingresar el importe", vbCritical, App.ProductName
        txtMonto.SetFocus
        Exit Sub
    End If


    objVenta.AgregaServicio strCodigoPadre, _
                        strDescripcionPadre, _
                        grdDocumentos.Columns(0).Value, _
                        grdDocumentos.Columns(1).Value, _
                        txtSuministro.Text, _
                        Servicio, _
                        txtMonto.Text, _
                        txtTC.Text, _
                        txtImporte.Text, _
                        grdDocumentos.Columns(2).Value, "", "", "0"

    frm_VTA_Servicios.grdServicios.Rebind
    Unload Me

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub
'
'Private Sub cmdCancelar_Click()
'    Unload Me
'End Sub



Private Sub cmdCancelar_Click()
Unload Me
End Sub

'Private Sub dbcMoneda_Click(Area As Integer)
'On Error GoTo handle
'    txtMonto_Change
'    Exit Sub
'handle:
'    MsgBox Err.Description, vbCritical, App.ProductName
'
'End Sub
'
'Private Sub dbcMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyEscape Then cmdCancelar_Click
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn And Shift = 1 Then cmdAceptar_Click
'End Sub

Private Sub Form_Load()
On Error GoTo handle
    SetteaFormulario Me
    txtTC.Text = objUsuario.TipoCambio
    Dim arrCampos, arrCaption, arrAncho, arrAlineacion
    arrCampos = Array("COD_SERVICIO", "DES_SERVICIO", "COD_PRODUCTO")
    arrCaption = Array("Codigo", "Servicio", "CodProducto")
    arrAncho = Array(1000, 4500, 0)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgGeneral)
    grdDocumentos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdDocumentos.Columns(2).Visible = False
    dbcMoneda.Enabled = False
    
    Set grdDocumentos.DataSource = objServicio.Lista("", strCodigoPadre)
    
        Set dbcMoneda.RowSource = ObjFormaPago.ListaMoneda
        dbcMoneda.ListField = "DES"
        dbcMoneda.BoundColumn = "COD"
        dbcMoneda.BoundText = objUsuario.Parametros("COD_MONEDA")(0, 2)
        Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
    
End Sub


'Private Sub grdDocumentos_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
'End Sub
'
'Private Sub grdDocumentos_RegistroSeleccionado(ByVal DatoColumna0 As String)
''lblComision.BackColor = RGB(252, 246, 207)
'
'lblComision.BackStyle = 1
'lblVencido.BackStyle = 1
''lblComision.BackColor = RGB(252, 246, 207)
'    With grdDocumentos
'        dbcMoneda.BoundText = .DataSource("COD_MONEDA").Value
'        If .DataSource("FLG_DOCUMENTO_VENCIDO").Value = "1" Then
'            lblVencido.Caption = "SI, Acepta recibos vencidos"
'        Else
'            lblVencido.Caption = "NO, acepta recibos vencidos"
'            '== Autor   Jahzeel López
'            '== Fecha   15/02/2007
'            '== Motivo  Se pone en True la propiedad Visible de la etiqueta lblVencido.
'            '==         para que no aparezca el mensaje, por indicación de etijero@btl.com.pe
'                lblVencido.Visible = False
'        End If
'
'        If .DataSource("FLG_COMISION_CLIENTE").Value = "1" Then
'            lblComision.Visible = True
'            If .DataSource("FLG_TIPO_VALOR_CLIE").Value = "0" Then
'                lblComision.Caption = "RECORDAR: Se cobra el S/." & .DataSource("IMP_VALOR_CLIE").Value & " al cliente"
'            Else
'                lblComision.Caption = "RECORDAR: Se cobra el " & .DataSource("IMP_VALOR_CLIE").Value & "% al cliente"
'            End If
'        Else
'            lblComision.Visible = False
'        End If
'    End With
'End Sub

Private Sub txtMonto_Change()
On Error GoTo handle
        If txtMonto.Text = "" Then txtImporte.Text = "0.00": Exit Sub
            Select Case dbcMoneda.BoundText
                Case "1"
                        txtImporte.Text = Val(txtMonto.Text)
                Case "2"
                        txtImporte.Text = Val(txtMonto.Text) * Val(txtTC.Text)
            End Select

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub
'
'Private Sub txtMonto_GotFocus()
'    cmdAceptar.Default = True
'End Sub
'
'Private Sub txtMonto_LostFocus()
'    cmdAceptar.Default = False
'End Sub


