VERSION 5.00
Begin VB.Form frm_VTA_ObservaAutorizacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Observaciones de Autorización"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlTextBox ctlTextBox1 
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1085
      Tipo            =   3
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbp_Ventas.ctlTextBox txtObservacion 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1508
      Tipo            =   2
      MaxLength       =   400
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   6960
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Command2"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Si tiene el Número de Solicitud por favor ingresela:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese el motivo por el cual esta solicitando la aprobación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "frm_VTA_ObservaAutorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OutObservacion As String
Public OutNumeroId As String

Private Sub Command1_Click()
    OutObservacion = txtObservacion.Text
    OutNumeroId = ctlTextBox1.Text
    Unload Me
End Sub

Private Sub Command2_Click()
    OutObservacion = ""
    OutNumeroId = ""
    Unload Me
End Sub

Private Sub Form_Load()
    OutObservacion = ""
    OutNumeroId = ""
End Sub
