VERSION 5.00
Begin VB.Form frmMensajePromo 
   BackColor       =   &H001C25DA&
   BorderStyle     =   0  'None
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox LblMensajePromo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001C25DA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   1815
      Left            =   158
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "frmMensajePromo.frx":0000
      Top             =   1440
      Width           =   8175
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3518
      TabIndex        =   0
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "!!GANASTE!!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0006F0FF&
      Height          =   1575
      Left            =   675
      TabIndex        =   1
      Top             =   120
      Width           =   8535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "!!GANASTE!!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00752B11&
      Height          =   1575
      Left            =   795
      TabIndex        =   2
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "frmMensajePromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pRetMensaje As String

Private Sub Form_Load()
    LblMensajePromo.Text = pRetMensaje
End Sub

Private Sub cmdSalir_Click()
If MsgBox("¿Esta seguro que desea salir?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
    Unload Me
End If
End Sub
