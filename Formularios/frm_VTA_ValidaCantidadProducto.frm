VERSION 5.00
Begin VB.Form frm_VTA_ValidaCantidadProducto 
   Caption         =   "Validacion por Producto"
   ClientHeight    =   2205
   ClientLeft      =   6150
   ClientTop       =   6420
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   3735
   Begin VB.CommandButton cmdAcepta 
      Caption         =   "Acepta"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin vbp_Ventas.ctlTextBox txtUsuario 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Alignment       =   2
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
   Begin vbp_Ventas.ctlTextBox txtPassword 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Alignment       =   2
      PasswordChar    =   "*"
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
   Begin vbp_Ventas.ctlTextBox txtCodBarra 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Alignment       =   2
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
   Begin VB.Label lblDescripcion 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   75
   End
   Begin VB.Label lblCodigo 
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblCodBarra 
      AutoSize        =   -1  'True
      Caption         =   "C Barra :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      Caption         =   "Usuario :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   945
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1155
   End
End
Attribute VB_Name = "frm_VTA_ValidaCantidadProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objProducto As New clsProducto
Dim oraDato As oraDynaset
Dim rsUsuario As oraDynaset

Dim strFlgLocal, strFlgQuimico, strFlgOtros, strFlgUsuario, strFlgPassword, strFlgScanner As String
    
Public Sub subDatos(strCodProd As String, strDesProd As String, pstrFlgLocal As String, pstrFlgQuimico As String, pstrFlgOtros As String, pstrFlgPassword As String, pstrFlgScanner As String)
    
    
    lblCodigo.Caption = strCodProd
    lblDescripcion.Caption = strDesProd
    
    strFlgLocal = pstrFlgLocal
    strFlgQuimico = pstrFlgQuimico
    strFlgOtros = pstrFlgOtros
'    strFlgUsuario = pstrFlgUsuario
    strFlgPassword = pstrFlgPassword
    strFlgScanner = pstrFlgScanner
    
    txtUsuario.Visible = False
    txtPassword.Visible = False
    txtCodBarra.Visible = False
    lblUsuario.Visible = False
    lblPassword.Visible = False
    lblCodBarra.Visible = False
    
'    If strFlgLocal = "1" Then
    If strFlgQuimico = "1" Then
        strFlgUsuario = "1"
    End If
    
    If strFlgOtros = "1" Then
        strFlgUsuario = "1"
    End If
    
    If (strFlgOtros = "" And strFlgQuimico = "") Then
        strFlgUsuario = "0"
    End If
    
        If strFlgUsuario = "1" Then
            txtUsuario.Visible = True
            lblUsuario.Visible = True
            If strFlgPassword = "1" Then
                txtPassword.Visible = True
                lblPassword.Visible = True
                If strFlgScanner = "1" Then
                    txtCodBarra.Visible = True
                    lblCodBarra.Visible = True
                    Me.Height = 2715
                Else
                    Me.Height = 2340
                End If
            Else
                If strFlgScanner = "1" Then
                    txtCodBarra.Visible = True
                    lblCodBarra.Visible = True
                    txtCodBarra.top = txtCodBarra.top - 380
                    lblCodBarra.top = lblCodBarra.top - 380
                    Me.Height = 2340
                Else
                    Me.Height = 1965
                End If
            End If
        Else
            If strFlgScanner = "1" Then
                txtCodBarra.Visible = True
                lblCodBarra.Visible = True
                txtCodBarra.top = txtCodBarra.top - 770
                lblCodBarra.top = lblCodBarra.top - 770
                Me.Height = 1965
            End If
        End If
    


    'Me.Caption = pstrCaption
    Me.top = Me.top / 1.5
    Me.left = Me.left / 1.5
    Me.Show vbModal

End Sub

Private Sub cmdAcepta_Click()
'On Error GoTo handle
On Error GoTo ERROR
    Dim n As Integer
    n = 0
    If txtUsuario.Visible = True Then
        If txtUsuario.Text = "" Then
            MsgBox "Debe ingresar su usuario", vbCritical, "Aviso"
            txtUsuario.Text = ""
            txtUsuario.SetFocus
            Exit Sub
        End If
    End If
    
    If txtPassword.Visible = True Then
        If txtPassword.Text = "" Then
            MsgBox "Debe ingresar su password", vbCritical, "Aviso"
            txtPassword.Text = ""
            txtPassword.SetFocus
            Exit Sub
        End If
    End If
    
    If txtCodBarra.Visible = True Then
        If txtCodBarra.Text = "" Then
            MsgBox "Debe scanner el producto", vbCritical, "Aviso"
            txtCodBarra.Text = ""
            txtCodBarra.SetFocus
            Exit Sub
        End If
    End If
    
    
    'CASO QUIMICO LOCAL
    'Falta validar si quimico del local
'        If strFlgQuimico = "1" And strFlgOtros = "0" And strFlgPassword = "0" And strFlgScanner = "0" Then
'            n = gclsOracle.FN_Valor("BTLPROD.pkg_usuario.FN_ES_QUIMICO", txtUsuario.Text)
'            If n = 0 Then
'                MsgBox "Es necesario que el Quimico ingrese su codigo", vbCritical
'                txtUsuario.SetFocus
'                Exit Sub
'            Else
'                Unload Me
'            End If
'        End If
    
    'CASO QUIMICO EMPRESA
        If strFlgQuimico = "1" And strFlgOtros = "0" And strFlgPassword = "0" And strFlgScanner = "0" Then
            n = gclsOracle.FN_Valor("BTLPROD.pkg_usuario.FN_ES_QUIMICO", txtUsuario.Text)
            If n = 0 Then
                MsgBox "Es necesario que el Quimico ingrese su codigo", vbCritical
                txtUsuario.Text = ""
                txtUsuario.SetFocus
                Exit Sub
            Else
                Unload Me
            End If
        End If
    
    'CASO OTROS
    If strFlgQuimico = "0" And strFlgOtros = "1" And strFlgPassword = "0" And strFlgScanner = "0" Then
        n = gclsOracle.FN_Valor("CMR.PKG_USUARIO.FN_USUARIO_VALIDO", txtUsuario.Text)
        If n = 1 Then
            Unload Me
        Else
            MsgBox "Codigo de usuario no existe o esta inactivo", vbCritical
            txtUsuario.Text = ""
            txtUsuario.SetFocus
            Exit Sub
        End If
    End If
    
    
    'CASO QUIMICO y PASSWORD
        If strFlgQuimico = "1" And strFlgOtros = "0" And strFlgPassword = "1" And strFlgScanner = "0" Then
            n = gclsOracle.FN_Valor("BTLPROD.pkg_usuario.FN_ES_QUIMICO", txtUsuario.Text)
            If n = 0 Then
                MsgBox "Es necesario que el Quimico ingrese su codigo", vbCritical
                txtUsuario.Text = ""
                txtUsuario.SetFocus
                Exit Sub
            Else
                Set rsUsuario = gclsOracle.FN_Cursor("CMR.PKG_USUARIO.FN_LOGIN_PROD", 0, txtUsuario.Text, txtPassword.Text)
                If LTrim(RTrim(rsUsuario(0))) <> LTrim(RTrim(txtUsuario.Text)) Then
                    MsgBox "Contrasena errada", vbCritical
                    txtPassword.Text = ""
                    txtPassword.SetFocus
                    Exit Sub
                Else
                    Unload Me
                End If
            End If
        End If
    
    'CASO OTROS y PASSWORD
        If strFlgQuimico = "0" And strFlgOtros = "1" And strFlgPassword = "1" And strFlgScanner = "0" Then
            Set rsUsuario = gclsOracle.FN_Cursor("CMR.PKG_USUARIO.FN_LOGIN_PROD", 0, txtUsuario.Text, txtPassword.Text)
            If LTrim(RTrim(rsUsuario(0))) <> LTrim(RTrim(txtUsuario.Text)) Then
                MsgBox "DATOS ERRADOS", vbCritical
                txtPassword.Text = ""
                txtPassword.SetFocus
                Exit Sub
            Else
                Unload Me
            End If
        End If
    
    'CASO QUIMICO Y SCANNER
    If strFlgQuimico = "1" And strFlgOtros = "0" And strFlgPassword = "0" And strFlgScanner = "1" Then
        n = gclsOracle.FN_Valor("BTLPROD.pkg_usuario.FN_ES_QUIMICO", txtUsuario.Text)
        If n = 0 Then
            MsgBox "Es necesario que el Quimico ingrese su codigo", vbCritical
            txtUsuario.Text = ""
            txtUsuario.SetFocus
            Exit Sub
        Else
            If LTrim(RTrim(gclsOracle.FN_Valor("Nuevo.pkg_codigo_barra.FN_PRODUCTO", txtCodBarra.Text))) = LTrim(RTrim(lblCodigo.Caption)) Then
                Unload Me
            Else
                MsgBox "Error al ingresar el codigo de barra", vbCritical
                txtCodBarra.Text = ""
                txtCodBarra.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    'CASO OTROS Y SCANNER
    If strFlgQuimico = "0" And strFlgOtros = "1" And strFlgPassword = "0" And strFlgScanner = "1" Then
        n = gclsOracle.FN_Valor("CMR.PKG_USUARIO.FN_USUARIO_VALIDO", txtUsuario.Text)
        If n = 1 Then
            If LTrim(RTrim(gclsOracle.FN_Valor("Nuevo.pkg_codigo_barra.FN_PRODUCTO", txtCodBarra.Text))) = LTrim(RTrim(lblCodigo.Caption)) Then
                Unload Me
            Else
                MsgBox "Error al ingresar el codigo de barra", vbCritical
                txtCodBarra.Text = ""
                txtCodBarra.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Codigo de usuario no existe o esta inactivo", vbCritical
            txtUsuario.Text = ""
            txtUsuario.SetFocus
            Exit Sub
        End If
    End If
    
    'CASO SCANNER
        If (strFlgQuimico = "0" Or strFlgQuimico = "") And (strFlgOtros = "0" Or strFlgOtros = "") And strFlgPassword = "0" And strFlgScanner = "1" Then
            If LTrim(RTrim(gclsOracle.FN_Valor("Nuevo.pkg_codigo_barra.FN_PRODUCTO", txtCodBarra.Text))) = LTrim(RTrim(lblCodigo.Caption)) Then
                Unload Me
            Else
                MsgBox "Error al ingresar el codigo de barra", vbCritical
                txtCodBarra.Text = ""
                txtCodBarra.SetFocus
                Exit Sub
            End If
        End If
        
    'CASO QUIMICO, PASSWORD Y SCANNER
    If strFlgQuimico = "1" And strFlgOtros = "0" And strFlgPassword = "1" And strFlgScanner = "1" Then
        n = gclsOracle.FN_Valor("BTLPROD.pkg_usuario.FN_ES_QUIMICO", txtUsuario.Text)
        If n = 0 Then
            MsgBox "Es necesario que el Quimico ingrese su codigo", vbCritical
            txtUsuario.Text = ""
            txtUsuario.SetFocus
            Exit Sub
        Else
            Set rsUsuario = gclsOracle.FN_Cursor("CMR.PKG_USUARIO.FN_LOGIN_PROD", 0, txtUsuario.Text, txtPassword.Text)
            If LTrim(RTrim(rsUsuario(0))) <> LTrim(RTrim(txtUsuario.Text)) Then
                MsgBox "Contrasena errada", vbCritical
                txtPassword.Text = ""
                txtPassword.SetFocus
                Exit Sub
            Else
                If LTrim(RTrim(gclsOracle.FN_Valor("Nuevo.pkg_codigo_barra.FN_PRODUCTO", txtCodBarra.Text))) = LTrim(RTrim(lblCodigo.Caption)) Then
                    Unload Me
                Else
                    MsgBox "Error al ingresar el codigo de barra", vbCritical
                    txtCodBarra.Text = ""
                    txtCodBarra.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
    
    'CASO OTROS, PASSWORD Y SCANNER
    If strFlgQuimico = "0" And strFlgOtros = "1" And strFlgPassword = "1" And strFlgScanner = "1" Then
        Set rsUsuario = gclsOracle.FN_Cursor("CMR.PKG_USUARIO.FN_LOGIN_PROD", 0, txtUsuario.Text, txtPassword.Text)
        If LTrim(RTrim(rsUsuario(0))) <> LTrim(RTrim(txtUsuario.Text)) Then
            MsgBox "DATOS ERRADOS", vbCritical
            txtPassword.Text = ""
            txtPassword.SetFocus
            Exit Sub
        Else
            If LTrim(RTrim(gclsOracle.FN_Valor("Nuevo.pkg_codigo_barra.FN_PRODUCTO", txtCodBarra.Text))) = LTrim(RTrim(lblCodigo.Caption)) Then
                Unload Me
            Else
                MsgBox "Error al ingresar el codigo de barra", vbCritical
                txtCodBarra.Text = ""
                txtCodBarra.SetFocus
                Exit Sub
            End If
        End If
    End If
    'strFlgLocal, strFlgQuimico, strFlgOtros, strFlgPassword, strFlgScanner

Exit Sub
ERROR:
    Err.Raise vbObjectError, "clsOracle", Err.Description
'handle:
'    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub txtCodBarra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAcepta_Click
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAcepta_Click
    End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAcepta_Click
    End If
End Sub
