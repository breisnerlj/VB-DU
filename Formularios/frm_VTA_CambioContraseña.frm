VERSION 5.00
Begin VB.Form frm_VTA_CambioContraseña 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Contraseña"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   Icon            =   "frm_VTA_CambioContraseña.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3555
   StartUpPosition =   1  'CenterOwner
   Begin vbp_Ventas.ctlTextBox txtPassword 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Alignment       =   2
      MaxLength       =   8
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
   Begin vbp_Ventas.ctlTextBox txtPasswordNew 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Alignment       =   2
      MaxLength       =   8
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
   Begin vbp_Ventas.ctlTextBox txtRePasswordNew 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Alignment       =   2
      MaxLength       =   8
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
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   3360
      Y1              =   610
      Y2              =   610
   End
   Begin VB.Line Line 
      Index           =   0
      X1              =   120
      X2              =   3360
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblRePasswordNew 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña nueva :"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1290
      Width           =   1395
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña actual :"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   210
      Width           =   1380
   End
   Begin VB.Label lblPasswordNew 
      Caption         =   "Contraseña nueva :"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   810
      Width           =   1695
   End
End
Attribute VB_Name = "frm_VTA_CambioContraseña"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strCodigo As String
Private strPassword As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Public Sub Mostrar(ByVal strCodigoUsuario As String, ByVal strPasswordUsuario As String)
    strCodigo = strCodigoUsuario
    strPassword = strPasswordUsuario
    Me.Show vbModal
End Sub



Private Sub txtRePasswordNew_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vstrMensaje As String
On Error GoTo handle
    
    Select Case KeyCode
        Case vbKeyReturn

            If Len(txtPassword.Text) > 8 Then MsgBox "La contraseña no puede ser mayor a 8 caracteres.", _
                                                     vbOKOnly + vbExclamation, "Error": txtPassword.SetFocus: Exit Sub

            If Len(txtPasswordNew.Text) > 8 Then MsgBox "La nueva contraseña no puede ser mayor a 8 carateres.", _
                                                        vbOKOnly + vbExclamation, "Error": txtPasswordNew.SetFocus: Exit Sub

            If Len(txtRePasswordNew.Text) > 8 Then MsgBox "La verificación de la nueva contraseña no puede ser mayor a 8 carateres.", _
                                                          vbOKOnly + vbExclamation, "Error": txtRePasswordNew.SetFocus: Exit Sub

            
            'validando que el minimo sea 4
            If Len(txtPasswordNew.Text) < 4 Then
                MsgBox "La nueva contraseña debe ser mayor a 4 caracteres.", vbOKOnly + vbExclamation, "Error"
                txtPasswordNew.SetFocus
                Exit Sub
            End If
            
            'Validando el Password
            If Trim(txtPassword.Text) = "" Then
                MsgBox "Ingrese clave actual", vbCritical, App.ProductName
                txtPassword.SetFocus
                Exit Sub
            End If
            
            'Validando el Password Nuevo
            If Trim(txtPasswordNew.Text) = "" Then
                MsgBox "Ingrese la nueva clave", vbCritical, App.ProductName
                txtPasswordNew.SetFocus
                Exit Sub
            End If
            
            'Validando el Password Nuevo
            If Trim(txtRePasswordNew.Text) = "" Then
                MsgBox "Ingrese otra vez la nueva clave ", vbCritical, App.ProductName
                txtRePasswordNew.SetFocus
                Exit Sub
            End If

            If Trim(txtPasswordNew.Text) <> Trim(txtRePasswordNew.Text) Then
                MsgBox "La nueva clave y su verificación no son iguales.", vbCritical, App.ProductName
                txtPasswordNew.SetFocus
                Exit Sub
            End If

            If txtPassword.Text <> strPassword Then
                MsgBox "La contraseña actual no es correcta.", vbCritical, App.ProductName
                Exit Sub
            End If
            
            'validando que la clave nueva no sea igual a la anterior
            If txtPassword.Text = txtPasswordNew.Text Then
                MsgBox "La nueva contraseña no puede ser igual a la anterior.", vbOKOnly + vbExclamation, "Error"
                txtPasswordNew.SetFocus
                Exit Sub
            End If

            vstrMensaje = objUsuario.GrabaContraseña(strCodigo, _
                                                     txtPassword.Text, _
                                                     txtPasswordNew.Text)
                                                           
            If vstrMensaje <> "" Then
                MsgBox vstrMensaje, vbOKOnly + vbExclamation, "Mensaje"
            Else
                MsgBox "Se cambió correctamente la contraseña.", vbOKOnly + vbInformation, "Mensaje"
                Unload Me
            End If

    End Select
Exit Sub

handle:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error " & Err.Number
End Sub
