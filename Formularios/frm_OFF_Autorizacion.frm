VERSION 5.00
Begin VB.Form frm_OFF_Autorizacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorización"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "frm_OFF_Autorizacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlTextBox txtPassword 
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Top             =   900
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
      MaxLength       =   5
      PasswordChar    =   "*"
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
   Begin vbp_Ventas.ctlTextBox txtUsuario 
      Height          =   315
      Left            =   1500
      TabIndex        =   2
      Top             =   480
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
      Tipo            =   3
      MaxLength       =   5
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
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuario"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   540
      Width           =   540
   End
End
Attribute VB_Name = "frm_OFF_Autorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public bolAutorizacionOK As Boolean


Private Sub cmdAceptar_Click()
Dim bolBienCodigo As Boolean
Dim bolBienPassword As Boolean

On Error GoTo CtrlErr
    
    If Len(Trim(txtUsuario.Text)) = 0 Then
        MsgBox "Debe ingresar el Usuario", vbExclamation, App.ProductName
        Exit Sub
    End If
    If Len(Trim(txtPassword.Text)) = 0 Then
        MsgBox "Debe ingresar el Password", vbExclamation, App.ProductName
        Exit Sub
    End If
    
    bolAutorizacionOK = False
    bolBienCodigo = False
    bolBienPassword = False
    
    bolBienCodigo = ValidaUsuario(Trim(txtUsuario.Text))
    'If Not bolBienCodigo Then txtUsuario.SetFocus: Exit Sub
    bolBienPassword = ValidaPassword(txtPassword.Text)
    'If Not bolBienPassword Then txtPassword.SetFocus: Exit Sub
    
    If Not bolBienCodigo Or Not bolBienPassword Then
        MsgBox "El usuario no existe o no está asignado a este Local", vbCritical, App.ProductName
        Exit Sub
    End If
    
    If bolBienCodigo And bolBienPassword Then
        bolAutorizacionOK = True
        objOFFUsuario.UsuModPrecio = txtUsuario.Text
        Unload Me
    End If
    
Exit Sub

CtrlErr:

    MsgBox Err.Number & "-" & Err.Description, vbCritical, App.ProductName
    
    
End Sub

Private Function ValidaUsuario(ByVal pstrCodUsuario As String) As Boolean

Dim strBuscar As String
Dim cnn As ADODB.Connection
Dim strSQL As String
Dim rsUsuario As New ADODB.Recordset
Dim bolEncontro As Boolean

On Error GoTo CtrlErr

    bolEncontro = False
    
    If Len(Trim(pstrCodUsuario)) > 5 Then
        Exit Function
    End If
    
'    Set cnn = New ADODB.Connection
'    cnn.Open gstrConexion
    
'    strSQL = "select * from usuario.txt"
'    rsUsuario.Open strSQL, cnn, adOpenStatic, adLockReadOnly, adCmdText
        
    rsUsuario.CursorLocation = adUseClient
    rsUsuario.Open strUsuariosXML, gstrConexion, adOpenStatic, adLockReadOnly
    rsUsuario.Filter = " COD_USUARIO = '" & pstrCodUsuario & "'"
    If rsUsuario.RecordCount > 0 Then bolEncontro = True
    
    ValidaUsuario = bolEncontro

    Exit Function
CtrlErr:
    MsgBox Err.Number & "-" & Err.Description, vbCritical, App.ProductName
End Function

Private Sub cmdCancelar_Click()
    objOFFUsuario.UsuModPrecio = ""
    Unload Me
End Sub

Private Function ValidaPassword(ByVal pstrPassword As String) As Boolean
Dim strAdminPass As String
Dim objArchivoIni As cls_ArchivoIni
Dim bolPasswordCorrecto As Boolean

On Error GoTo CtrlErr

    bolPasswordCorrecto = False

    Set objArchivoIni = New cls_ArchivoIni
    strAdminPass = objArchivoIni.LeerIni(gstrIni, "general", "ADMIN_PASS", "")
    
    If strAdminPass = pstrPassword Then bolPasswordCorrecto = True

    Set objArchivoIni = Nothing

    ValidaPassword = bolPasswordCorrecto

Exit Function

CtrlErr:

    MsgBox Err.Number & "-" & Err.Description, vbCritical, App.ProductName
End Function
