VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLoteFechVen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lote y fecha de vencimiento"
   ClientHeight    =   6960
   ClientLeft      =   7305
   ClientTop       =   2295
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   5760
   Begin vbp_Ventas.ctlDataCombo CboNroLote 
      Height          =   315
      Left            =   1560
      TabIndex        =   16
      Top             =   2880
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin VB.CommandButton cmdAgregar 
      Height          =   495
      Left            =   5040
      Picture         =   "frmLoteFechVen.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmdEliminar 
      Height          =   495
      Left            =   5040
      Picture         =   "frmLoteFechVen.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3360
      Width           =   615
   End
   Begin VB.CheckBox chkVariosLotes 
      Caption         =   "&Varios Lotes"
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
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox chkFraccionamiento 
      Caption         =   "&Fraccionamiento"
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
      Left            =   2460
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1980
      Width           =   2055
   End
   Begin MSMask.MaskEdBox txtFechaVencimiento 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin vbp_Ventas.ctlTextBox txtNroLote 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   2880
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      MaxLength       =   20
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
   Begin vbp_Ventas.ctlTextBox txtCantidad 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Tipo            =   3
      MaxLength       =   4
      EnabledFoco     =   0   'False
      TABAuto         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
   End
   Begin vbp_Ventas.ctlGrillaArray grdLoteFchVen 
      Height          =   2895
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3960
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5106
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad del producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   11
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmLoteFechVen.frx":0B14
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000040&
      Height          =   555
      Left            =   1560
      TabIndex        =   10
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label lblCodigo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Código :"
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
      Left            =   120
      TabIndex        =   8
      Top             =   750
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripción :"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1170
   End
   Begin VB.Label lblIndicador 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3600
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nro de Lote"
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
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1050
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Vencimiento"
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
      Left            =   120
      TabIndex        =   4
      Top             =   3510
      Width           =   1110
   End
End
Attribute VB_Name = "frmLoteFechVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lCodigoProducto As String
Dim lLoteObligatorio, lFechaObligatorio  As String
Dim objProducto As New clsProducto


Public Sub CargaDatos(ByVal LoteObligatorio As String, ByVal FechaObligatorio As String, codigoProducto As String)
    lCodigoProducto = codigoProducto
    lLoteObligatorio = LoteObligatorio
    lFechaObligatorio = FechaObligatorio
    Me.Show vbModal
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub CboNroLote_Change()
On Error GoTo CtrlErr
    If CboNroLote.Text = "NO TIENE N LOTE" Then
        txtNroLote.Visible = True
        txtNroLote.Text = " "
        txtNroLote.SetFocus
        CboNroLote.Visible = False
    Else
        If CboNroLote.Text <> "" Then
            If lFechaObligatorio = "1" Then
                txtFechaVencimiento.Text = CboNroLote.BoundText
            End If
            txtNroLote.Text = CboNroLote.Text
            txtNroLote.Enabled = True
            CboNroLote.Visible = True
            If lFechaObligatorio = "0" Then
                If (MsgBox("¿Desea agregar el lote " + txtNroLote.Text + " ?", vbCritical + vbYesNo) = vbYes) Then
                    Unload Me
                End If
            End If
        Else
            MsgBox "Seleccione un Nro de Lote"
        End If
    End If
Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub



'Private Sub chkVariosLotes_Click()
'    If chkVariosLotes.Value = 1 Then
'        cmdAgregar.Visible = True
'        cmdEliminar.Visible = True
'        frmLoteFechVen.Height = 7410
'    Else
'        cmdAgregar.Visible = False
'        cmdEliminar.Visible = False
'        frmLoteFechVen.Height = 4365
'    End If
'End Sub

Private Sub Form_Load()
'Dim objProducto As New clsProducto
    With objVenta
        CboNroLote.Visible = False
        lblCodigo.Caption = .Producto(lCodigoProducto, 0)
        lblDescripcion.Caption = .Producto(lCodigoProducto, 1)
        lblIndicador.Caption = "" & objProducto.IndicadorReceta(.Producto(lCodigoProducto, 21))
        txtCantidad.Locked = True
        txtCantidad.TabStop = False
        txtCantidad.Text = .Producto(lCodigoProducto, 3)
        chkFraccionamiento = .Producto(lCodigoProducto, 6)
        chkFraccionamiento.Enabled = False
        txtNroLote.Text = .Producto(lCodigoProducto, 22)
        
        Label4.Visible = False
        txtNroLote.Visible = True
        Label6.Visible = False
        
        txtFechaVencimiento.Visible = False
        
        ''** Esta línea se comento porque el valor que le esta mandado al objeto txtFechaVecimiento es un 0 o 1 flag
        ''** y se cae ya que el objeto espera un valor de tipo fecha - comentado 18/02/2010 Cristhian Rueda
        
        txtFechaVencimiento.Text = IIf(IsDate(.Producto(lCodigoProducto, 23)), .Producto(lCodigoProducto, 23), txtFechaVencimiento.Text)
        
        If lLoteObligatorio = "1" Then
            Label4.Visible = True
            'txtNroLote.Visible = True
    
            Set CboNroLote.RowSource = objProducto.ListaLote(lblCodigo.Caption, objUsuario.CodigoLocal) 'objUsuario.ListaUsuarioDLV
            CboNroLote.ListField = "NLOTE"
            CboNroLote.BoundColumn = "FVENC"
            CboNroLote.BoundText = "*"
            CboNroLote.Visible = True
            txtNroLote.Visible = False
    
        End If
        If lFechaObligatorio = "1" Then
            Label6.Visible = True
            txtFechaVencimiento.Visible = True
        End If
    End With
    
    chkVariosLotes.Value = 0
    cmdAgregar.Visible = False
    cmdEliminar.Visible = False
    frmLoteFechVen.Height = 4365
Set objProducto = Nothing
End Sub

Sub salir()
    
    If lLoteObligatorio = "1" And txtNroLote.Text = "" Then
        MsgBox "Debe de Ingresar el número de lote", vbCritical, App.ProductName
        'txtNroLote.Focus
        Exit Sub
    End If
    If lFechaObligatorio = "1" And Replace(Replace(txtFechaVencimiento.Text, "_", ""), "/", "") = "" Then
        MsgBox "Debe de Ingresar la fecha de vencimiento", vbCritical, App.ProductName
        txtFechaVencimiento.SetFocus
        Exit Sub
    End If
        objVenta.Producto(lCodigoProducto, 22) = txtNroLote.Text
        objVenta.Producto(lCodigoProducto, 23) = txtFechaVencimiento.Text
    Unload Me
End Sub

Private Sub txtCantidad_GotFocus()
    If objVenta.CodigoTipoVenta = Guias_Remision Then
        'ctlGrillaArray3.Array1 = muestraArray(objVenta.ProductoLote, lblCodigo.Caption)
        If objProducto.DevIndicadorLote(lblCodigo.Caption) = "N" Then
            CboNroLote.Visible = False
            txtNroLote.Visible = True
            txtNroLote.Text = "NOTHING"
            txtNroLote.Enabled = False
        End If
    End If
End Sub

Private Sub txtFechaVencimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then salir
End Sub

Private Sub txtFechaVencimiento_KeyPress(KeyAscii As Integer)
Dim dfecha As Date
On Error GoTo CtrlErr

    dfecha = ("1/" & month(objUsuario.sysdate) & "/" & year(objUsuario.sysdate))
    If KeyAscii = 13 Then
        If txtFechaVencimiento.Text = "__/__/____" Then
            Call salir
        Else
            If Not IsDate(txtFechaVencimiento.Text) Or CDate(txtFechaVencimiento.Text) < dfecha Then
                MsgBox "La fecha de vencimiento es incorrecta", vbCritical + vbOKOnly, App.ProductName
                KeyAscii = 0
            Else
                Call salir
            End If
        End If
    
    End If
    
    
Exit Sub

CtrlErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName


End Sub

Private Sub txtNroLote_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And txtFechaVencimiento.Visible = False Then salir
End Sub

Private Sub CboNroLote_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And txtFechaVencimiento.Visible = False Then salir
End Sub

