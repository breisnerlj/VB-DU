VERSION 5.00
Begin VB.Form frm_DLV_EntregaTercero 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos de Entrega a Tercero"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "frm_DLV_EntregaTercero.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlTextBox TxtAuxNomb 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      Enabled         =   0   'False
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
   Begin vbp_Ventas.ctlTextBox TxtAuxDirecc 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      Enabled         =   0   'False
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
   Begin vbp_Ventas.ctlTextBox TxtAuxRefren 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      Enabled         =   0   'False
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
   Begin vbp_Ventas.ctlTextBox TxtAuxTelef 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      Enabled         =   0   'False
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
   Begin vbp_Ventas.ctlTextBox TxtAuxDistrito 
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   1080
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      Enabled         =   0   'False
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Distrito"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Telefono"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Referencia"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Dirección"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombres"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "frm_DLV_EntregaTercero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pAuxNomb As String
Public pAuxDirecc As String
Public pAuxRefer As String
Public pAuxTelefono As String
Public pAuxDistrito As String

Private Sub Form_Load()
    
    TxtAuxNomb.Text = pAuxNomb
    TxtAuxDirecc.Text = pAuxDirecc
    TxtAuxRefren.Text = pAuxRefer
    TxtAuxTelef.Text = pAuxTelefono
    TxtAuxDistrito.Text = pAuxDistrito

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
