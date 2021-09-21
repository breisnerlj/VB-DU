VERSION 5.00
Begin VB.Form frm_OFF_CantidadProducto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cantidad del producto"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   Icon            =   "frm_OFF_CantidadProducto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlTextBox txtPrecio 
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   2460
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Tipo            =   4
      Alignment       =   1
      Enabled         =   0   'False
      MaxLength       =   8
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
   End
   Begin vbp_Ventas.ctlTextBox txtCantidad 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Tipo            =   3
      Alignment       =   1
      MaxLength       =   5
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
      Left            =   2520
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "CTRL+P : Cambiar de precio"
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
      Index           =   1
      Left            =   2700
      TabIndex        =   10
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Precio :"
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
      Left            =   180
      TabIndex        =   9
      Top             =   2490
      Width           =   675
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000040&
      Height          =   555
      Left            =   1440
      TabIndex        =   8
      Top             =   1260
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
      Left            =   1440
      TabIndex        =   7
      Top             =   780
      Width           =   1095
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
      Left            =   540
      TabIndex        =   6
      Top             =   180
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "frm_OFF_CantidadProducto.frx":1C9A2
      Top             =   60
      Width           =   480
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
      Left            =   180
      TabIndex        =   5
      Top             =   810
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
      Left            =   180
      TabIndex        =   4
      Top             =   1260
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad :"
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
      Left            =   180
      TabIndex        =   3
      Top             =   2010
      Width           =   900
   End
   Begin VB.Label lblIndicador 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3660
      TabIndex        =   2
      Top             =   180
      Width           =   2055
   End
End
Attribute VB_Name = "frm_OFF_CantidadProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private objProducto As New cls_OFF_Producto

Private strFlgFraccionamiento As String
Private intCtdProducto As Integer
Private dblPrcUnitario As Double
Private dblMtoIgv As Double
Private dblMtoExonerado As Double
Private dblMtoSubtotal As Double
Private strFlgModPrecio As String





Public Property Get FlgFraccionamiento() As String
    FlgFraccionamiento = strFlgFraccionamiento
End Property
Public Property Let FlgFraccionamiento(ByVal newValue As String)
    strFlgFraccionamiento = newValue
    CalcularTotales
End Property

Public Property Get CtdProducto() As Integer
    CtdProducto = intCtdProducto
End Property
Public Property Let CtdProducto(ByVal newValue As Integer)
    intCtdProducto = newValue
    txtCantidad.Text = CtdProducto
    CalcularTotales
End Property

Public Property Get PrcUnitario() As Double
    PrcUnitario = dblPrcUnitario
End Property
Public Property Let PrcUnitario(ByVal newValue As Double)
    dblPrcUnitario = newValue
    txtPrecio.Text = PrcUnitario
    CalcularTotales
End Property

Public Property Get MtoIgv() As Double
    MtoIgv = dblMtoIgv
End Property
Public Property Let MtoIgv(ByVal newValue As Double)
    dblMtoIgv = newValue
End Property

Public Property Get MtoExonerado() As Double
    MtoExonerado = dblMtoExonerado
End Property
Public Property Let MtoExonerado(ByVal newValue As Double)
    dblMtoExonerado = newValue
End Property

Public Property Get MtoSubTotal() As Double
    MtoSubTotal = dblMtoSubtotal
End Property
Public Property Let MtoSubTotal(ByVal newValue As Double)
    dblMtoSubtotal = newValue
End Property

Public Property Get FlgModPrecio() As String
    FlgModPrecio = strFlgModPrecio
End Property
Public Property Let FlgModPrecio(ByVal newValue As String)
    strFlgModPrecio = newValue
End Property





Private Sub chkFraccionamiento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtCantidad_KeyDown KeyCode, Shift
    End If
End Sub

Private Sub chkFraccionamiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub chkFraccionamiento_Validate(Cancel As Boolean)
On Error GoTo Handle
    
    If chkFraccionamiento.Value = 0 Then
        FlgFraccionamiento = "U"
    Else
        FlgFraccionamiento = "F"
    End If
    

    
    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim bolEnabled As Boolean
On Error GoTo Handle
    
    Dim tmpCtrl As Boolean, tmpAlt As Boolean
    tmpCtrl = (Shift And vbCtrlMask) > 0
    tmpAlt = (Shift And vbAltMask) > 0

    Select Case KeyCode
        
        Case vbKeyEscape
            Unload Me
        
        Case tmpCtrl And vbKeyP
            If Val(txtCantidad.Text) = 0 Then
                MsgBox "La cantidad tiene que ser mayor que CERO", vbCritical, App.ProductName: txtCantidad.SetFocus
                Exit Sub
            End If
        
            frm_OFF_Autorizacion.Show vbModal
            
            
            txtPrecio.Enabled = False
            
            
            If frm_OFF_Autorizacion.bolAutorizacionOK Then
                txtPrecio.Enabled = True
                txtPrecio.SetFocus
            End If
    
    End Select

    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0
If objProducto.FlgFracciona = "N" Then
     chkFraccionamiento.Enabled = False
Else
    If UCase(Chr(KeyAscii)) = "F" And chkFraccionamiento.Visible Then _
       chkFraccionamiento.Value = IIf(chkFraccionamiento.Value = 1, 0, 1)
End If
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    
    lblCodigo.Caption = objProducto.Codigo
    lblDescripcion.Caption = objProducto.Descripcion
    lblIndicador.Caption = objProducto.ConReceta
    
    PrcUnitario = objProducto.PrecioPublico
    
    If objProducto.FlgFracciona = "S" Then
        chkFraccionamiento.Enabled = True
    Else
        chkFraccionamiento.Value = 0
        chkFraccionamiento.Enabled = False
    End If
    
    chkFraccionamiento_Validate False
    FlgModPrecio = "N"
    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub





Public Sub CargarDesdeProducto()
On Error GoTo Handle

    lblCodigo.Caption = objProducto.Codigo
    lblDescripcion.Caption = objProducto.Descripcion
    lblIndicador.Caption = objProducto.ConReceta

    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub CalcularTotales()
On Error GoTo Handle

    If FlgFraccionamiento = "F" Then
        MtoSubTotal = Round(PrcUnitario * CtdProducto / IIf(objProducto.CtdFracciona <= 0, 1, objProducto.CtdFracciona), 2)
    Else
        MtoSubTotal = Round(PrcUnitario * CtdProducto, 2)
    End If
    
    If objProducto.PctIgv = 0 Then
        MtoIgv = 0
    Else
        'Autor : Arturo Escate
        'Fecha :07/108/2008
        'Proposito : Se corrigio el calculo del impuesto ya que este estav mal
        'MtoIgv = Round(MtoSubTotal / objProducto.PctIgv, 2)
        MtoIgv = Round(MtoSubTotal - (MtoSubTotal / (1 + (objProducto.PctIgv / 100))), 2)
    End If
    
    MtoExonerado = IIf(objProducto.PctIgv = 0, MtoSubTotal, 0)

    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub txtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            ValidaDatos
    End Select

End Sub

Private Sub txtCantidad_Validate(Cancel As Boolean)
On Error GoTo Handle

    Dim intCantidad As Integer

    intCantidad = Val(txtCantidad.Text)
    
    If intCantidad < 0 Then
        txtCantidad.Text = "0"
        intCantidad = 0
    End If
    
    CtdProducto = intCantidad

    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub AgregarItem()
On Error GoTo Handle

    Dim strDescripcion As String
    Dim xdb As New XArrayDB

    'validamos
    If objProducto.Codigo = "" Then
        Exit Sub
    End If
    
    If CtdProducto <= 0 Then
        Exit Sub
    End If
    If PrcUnitario < 0 Then
        Exit Sub
    End If
    
    strDescripcion = UCase(Mid(objProducto.Descripcion, 1, 1)) + LCase(Mid(objProducto.Descripcion, 2))
    
    frm_OFF_Principal.grdDetalleVenta.Array1 = objOFFVenta.AgregaDetalleVenta( _
                                                    objProducto.Codigo, _
                                                    strDescripcion, _
                                                    FlgFraccionamiento, _
                                                    CtdProducto, _
                                                    objProducto.PctDescuento, _
                                                    PrcUnitario, _
                                                    objProducto.PrecioPublico, _
                                                    MtoIgv, _
                                                    MtoExonerado, _
                                                    MtoSubTotal, _
                                                    FlgModPrecio, _
                                                    objProducto.FlgRegalo, _
                                                    objProducto.PctIgv, _
                                                    objProducto.PartidaArancelaria, _
                                                    objProducto.CtdFracciona, _
                                                    objOFFUsuario.UsuModPrecio, _
                                                    objProducto.DescripcionCorta, _
                                                    objProducto.PrecioPublico, _
                                                    objProducto.ConReceta, _
                                                    objProducto.FlgFracciona)
    frm_OFF_Principal.MostrarTotales
    frm_OFF_Principal.grdDetalleVenta.Rebind
    
    
    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub


Private Sub txtPrecio_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn
            ValidaDatos
    End Select


End Sub

Private Sub txtPrecio_Validate(Cancel As Boolean)
On Error GoTo Handle

    Dim ldblPrecio As Double
    
    ldblPrecio = Val(txtPrecio.Text)
    
    If ldblPrecio < 0 Then
        txtPrecio.Text = "0"
        ldblPrecio = 0
    End If
    
    PrcUnitario = ldblPrecio
    
    If PrcUnitario <> objProducto.PrecioPublico Then
        FlgModPrecio = "S"
    End If

    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub


Public Sub Datos(ByVal pstrCodigo As String, _
            ByVal pstrDescripcion As String, _
            ByVal pdblPctDescuento As Double, _
            ByVal pdblPrecioPublico As Double, _
            ByVal pstrFlgRegalo As String, _
            ByVal pdblPctIGV As Double, _
            ByVal pstrPartidaArancelaria As String, _
            ByVal pintCtdFracciona As Integer, _
            ByVal pstrDescripcionCorta As String, _
            ByVal pstrConReceta As String, _
            ByVal pstrFlgFracciona As String)

            
On Error GoTo CtrlErr

    objProducto.Codigo = pstrCodigo
    objProducto.Descripcion = pstrDescripcion
    objProducto.PctDescuento = pdblPctDescuento
    objProducto.PrecioPublico = pdblPrecioPublico
    objProducto.FlgRegalo = pstrFlgRegalo
    objProducto.PctIgv = pdblPctIGV
    objProducto.PartidaArancelaria = pstrPartidaArancelaria
    objProducto.CtdFracciona = pintCtdFracciona
    objProducto.DescripcionCorta = pstrDescripcionCorta
    objProducto.ConReceta = pstrConReceta
    objProducto.FlgFracciona = pstrFlgFracciona

    Me.Show vbModal

Exit Sub

CtrlErr:

    MsgBox Err.Description, vbCritical, App.ProductName
End Sub



Private Sub ValidaDatos()
            chkFraccionamiento_Validate False
            txtPrecio_Validate False
            txtCantidad_Validate False
            AgregarItem
            Unload Me
End Sub
