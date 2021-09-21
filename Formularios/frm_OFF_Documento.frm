VERSION 5.00
Begin VB.Form frm_OFF_Documento 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin vbp_Ventas.ctlGrillaArray grdTipoDocumento 
      Height          =   1875
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   3307
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox txtDireccion 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   4260
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
      Tipo            =   8
      MaxLength       =   100
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
   Begin vbp_Ventas.ctlTextBox txtCliente 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   3840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
      Tipo            =   8
      MaxLength       =   100
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
   Begin vbp_Ventas.ctlTextBox txtNumDoc 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   3420
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      Tipo            =   3
      MaxLength       =   11
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
   Begin VB.ComboBox cboTipoCliente 
      Height          =   315
      ItemData        =   "frm_OFF_Documento.frx":0000
      Left            =   1560
      List            =   "frm_OFF_Documento.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2970
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5205
      Picture         =   "frm_OFF_Documento.frx":0025
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   3900
      Picture         =   "frm_OFF_Documento.frx":05AF
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   1095
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
      TabIndex        =   15
      Top             =   2640
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Persona :"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   3030
      Width           =   1035
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
      Index           =   8
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Documento"
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
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "frm_OFF_Documento.frx":0B39
      Top             =   120
      Width           =   240
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
      Left            =   5550
      TabIndex        =   11
      Top             =   6780
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
      Left            =   3840
      TabIndex        =   10
      Top             =   6780
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Razón Social:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   3900
      Width           =   990
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "RUC:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   3465
      Width           =   390
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Dirección:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   4335
      Width           =   720
   End
End
Attribute VB_Name = "frm_OFF_Documento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub cboTipoCliente_Click()
    If left(cboTipoCliente.Text, 1) = "0" Then
        lblTitle(1).Caption = "DNI:"
        lblTitle(2).Caption = "Nombres:"
        objOFFCliente.Tipo = "0"
    Else
        lblTitle(1).Caption = "RUC:"
        lblTitle(2).Caption = "Razón Social:"
        objOFFCliente.Tipo = "1"
    End If

End Sub

Private Sub cboTipoCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo CtrlErr
    If grdTipoDocumento.ApproxCount > 0 Then
        objOFFVenta.TipoDocumento = grdTipoDocumento.Columns(0).Value
        objOFFVenta.NumDocumento = grdTipoDocumento.Columns(2).Value
    Else
        Exit Sub
    End If
    
    If objOFFCliente.Tipo = "1" And objOFFVenta.TipoDocumento = COD_TIPO_BOL Then
    
        Err.Raise vbObjectError + 513, "frm_OFF_Documento", "Tiene que seleccionar Factura para tipo de cliente Persona Jurídica"
        Exit Sub
    End If
    
    objOFFCliente.Ruc = txtNumDoc.Text
    If objOFFCliente.Tipo = "1" And objOFFCliente.ValidaRuc(objOFFCliente.Ruc) <> 0 Then
        Err.Raise vbObjectError + 513, "frm_OFF_Documento", "El número de RUC no es Válido"
        Exit Sub
    End If
    objOFFCliente.Nombre = txtCliente.Text
    
    If objOFFCliente.Tipo = "1" And objOFFCliente.Nombre = "" Then
        Err.Raise vbObjectError + 513, "frm_OFF_Documento", "Indicar la razón social"
        Exit Sub
    End If
    
    objOFFCliente.Direccion = txtDireccion.Text
    
    If objOFFCliente.Tipo = "0" And Len(objOFFCliente.Ruc) > 8 Then
        Err.Raise vbObjectError + 513, "frm_OFF_Documento", "el número de DNI es incorrecto"
        Exit Sub
    End If
    
    
    
    frm_OFF_Principal.Siguiente
    
    
    
    Unload Me 'Me.Hide
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    grdTipoDocumento.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo CtrlErr
    
    Dim tmpCtrl As Boolean, tmpAlt As Boolean
    
    tmpCtrl = (Shift And vbCtrlMask) > 0
    tmpAlt = (Shift And vbAltMask) > 0

    Select Case KeyCode
        Case vbKeyReturn
            If Shift = 1 Then Call cmdAceptar_Click
        Case vbKeyF1
            grdTipoDocumento.SetFocus
        Case vbKeyF2
            cboTipoCliente.SetFocus
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
        Case vbKeyF5 And False
            cmdAceptar_Click
            frm_OFF_Principal.Form_KeyDown KeyCode, Shift
        Case vbKeyF6
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
    Call CargarTipoDocumento

    grdTipoDocumento.MoveFirst
    Do While Not grdTipoDocumento.EOF
        If grdTipoDocumento.Columns(0).Value = objOFFVenta.TipoDocumento Then
            Exit Do
        End If
        grdTipoDocumento.MoveNext
    Loop
    cboTipoCliente.ListIndex = IIf(Len(objOFFCliente.Tipo) = 0, 0, objOFFCliente.Tipo)
    txtNumDoc.Text = objOFFCliente.Ruc
    txtCliente.Text = objOFFCliente.Nombre
    txtDireccion.Text = objOFFCliente.Direccion

End Sub


Private Sub CargarTipoDocumento()
Dim objDocumento As cls_OFF_Documento
Dim RowFound As Long

On Error GoTo CtrlErr

    Set objDocumento = New cls_OFF_Documento
    
    
    grdTipoDocumento.Array1 = objDocumento.ListaTipoDocumento(1)
    
    grdTipoDocumento.Rebind
    
    RowFound = objDocumento.ListaTipoDocumento.Find(1, 0, objOFFVenta.TipoDocumento, XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
    
    If RowFound >= 0 Then grdTipoDocumento.Bookmark = RowFound
    
    
    Set objDocumento = Nothing

Exit Sub

CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub



Private Sub SetGrid()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim columna As TrueDBGrid70.Column
  
    arrCampos = Array("", "", "", "", "")
    arrCaption = Array("Código", "Descripción", "Correlativo", "Líneas", "Ancho")
    arrAncho = Array(1000, 3000, 0, 0, 0)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft)
    grdTipoDocumento.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    For Each columna In grdTipoDocumento.Columns
        columna.AllowSizing = False
        columna.Visible = False
    Next
    
    grdTipoDocumento.Columns(0).Visible = True
    grdTipoDocumento.Columns(1).Visible = True
    
    grdTipoDocumento.MarqueeStyle = dbgHighlightRow
    grdTipoDocumento.CambiaSeleccionadoBackColor &H800000
    grdTipoDocumento.CambiaSeleccionadoForeColor &HFFFFFF
    
    
End Sub


Private Sub grdTipoDocumento_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo CtrlErr

    Select Case KeyCode
        Case vbKeyReturn
            If Shift = 1 Then
                Call cmdAceptar_Click
            Else
                cboTipoCliente.SetFocus
            End If
    
    End Select
    

Exit Sub

CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub



Public Sub Limpia()
    txtNumDoc.Text = ""
    txtCliente.Text = ""
    txtDireccion.Text = ""
    
End Sub

Private Sub grdTipoDocumento_RegistroSeleccionado(ByVal DatoColumna0 As String)
    
    If (grdTipoDocumento.Columns(0).Value = "FAC") Or (grdTipoDocumento.Columns(0).Value = "TKF") Then
        cboTipoCliente.ListIndex = 1
    Else
        cboTipoCliente.ListIndex = 0
    End If
    
    'objVenta.Out_Tipo = ctlCboTipCliente.BoundText
End Sub
