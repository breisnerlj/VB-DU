VERSION 5.00
Begin VB.Form frm_VTA_ValidaCotizacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Validación"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3600
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   2160
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin vbp_Ventas.ctlTextBox txtValidaUsuario 
      Height          =   375
      Left            =   1620
      TabIndex        =   0
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Alignment       =   2
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
   Begin vbp_Ventas.ctlTextBox txtValidaPassword 
      Height          =   375
      Left            =   1620
      TabIndex        =   1
      Top             =   720
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
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   3
      Top             =   300
      Width           =   945
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   2
      Top             =   780
      Width           =   1155
   End
End
Attribute VB_Name = "frm_VTA_ValidaCotizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim odynValidaCotiza As oraDynaset

Private Sub cmdCancelar_Click()
   gstrValidaCotizacion = "2"
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error GoTo handle
    Select Case KeyCode
        Case vbKeyReturn
            If txtValidaUsuario.Text = "" Then Screen.MousePointer = vbDefault: Exit Sub
            If txtValidaPassword.Text = "" Then Screen.MousePointer = vbDefault: Exit Sub
            
            Set odynValidaCotiza = objUsuario.fnValidaCotiza(Trim(txtValidaUsuario.Text), Trim(txtValidaPassword.Text), objUsuario.CodigoAplicacion)
            gstrValidaCotizacion = "0"
            
            If odynValidaCotiza.RecordCount > 0 Then
                gstrValidaCotizacion = "1"
                Unload Me
            Else
                gstrValidaCotizacion = "0"
            End If
    End Select
    Exit Sub
handle:
 MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
  End Sub


