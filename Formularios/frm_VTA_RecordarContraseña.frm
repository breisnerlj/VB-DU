VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_VTA_RecordarContraseña 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recordatorio de  contraseña"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3180
   Icon            =   "frm_VTA_RecordarContraseña.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3180
   StartUpPosition =   1  'CenterOwner
   Begin vbp_Ventas.ctlTextBox txtDNI 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Tipo            =   4
      Alignment       =   2
      MaxLength       =   8
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
   Begin MSMask.MaskEdBox mskFchNacimiento 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   3000
      Y1              =   1330
      Y2              =   1330
   End
   Begin VB.Line Line 
      Index           =   0
      X1              =   120
      X2              =   3000
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Su clave es :"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1140
   End
   Begin VB.Label lblFchNacimiento 
      Alignment       =   2  'Center
      Caption         =   "Ingrese Fecha Nacimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblDNI 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese su D.N.I."
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
      Left            =   120
      TabIndex        =   1
      Top             =   210
      Width           =   1470
   End
   Begin VB.Label lblContraseña 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   2895
   End
End
Attribute VB_Name = "frm_VTA_RecordarContraseña"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strCodigo As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub mskFchNacimiento_GotFocus()
    mskFchNacimiento.BackColor = txtDNI.ColorFoco
End Sub

Private Sub mskFchNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Respuesta As String
On Error GoTo handle
    
    Select Case KeyCode
        Case vbKeyReturn

            If txtDNI.Text = "" Or Len(txtDNI.Text) < 8 Then MsgBox "Debe ingresar correctamente su número de DNI.", vbOKOnly + vbExclamation, "Error": txtDNI.SetFocus: Exit Sub

            If Len(mskFchNacimiento.ClipText) < 8 Then MsgBox "Debe ingresar correctamente su fecha de nacimiento.", vbOKOnly + vbExclamation, "Error": mskFchNacimiento.SetFocus: Exit Sub

            'Validando el dia
            If Mid(mskFchNacimiento.ClipText, 1, 2) > 31 Then
                MsgBox "El día ingresado no es valido", vbCritical, App.ProductName
                mskFchNacimiento.SetFocus
                Exit Sub
            End If
            
            'Validando el mes
            If Mid(mskFchNacimiento.ClipText, 3, 2) > 12 Then
                MsgBox "El mes ingresado no es valido", vbCritical, App.ProductName
                mskFchNacimiento.SetFocus
                Exit Sub
            End If
            
            'Validando el año
            If Val(Mid(mskFchNacimiento.ClipText, 5, 4)) > Year(Now) Then
                MsgBox "El año ingresado no es valido", vbCritical, App.ProductName
                mskFchNacimiento.SetFocus
                Exit Sub
            End If
        
            Respuesta = objUsuario.EvocarContraseña(strCodigo, txtDNI.Text, mskFchNacimiento.Text)
            
            If Respuesta = "0" Then
                lblContraseña.FontSize = 7
                lblContraseña.BackColor = &H80000004
                lblContraseña.Caption = "Los datos resqueridos no son validos ... Verifique"
            Else
                lblContraseña.FontSize = 12
                lblContraseña.Caption = Respuesta
                lblContraseña.BackColor = &H80000018
            End If
    End Select
Exit Sub

handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub mskFchNacimiento_LostFocus()
    mskFchNacimiento.BackColor = txtDNI.ColorDefault
End Sub

Public Sub Mostrar(ByVal strCodigoUsuario As String)
    strCodigo = strCodigoUsuario
    Me.Show vbModal
End Sub
