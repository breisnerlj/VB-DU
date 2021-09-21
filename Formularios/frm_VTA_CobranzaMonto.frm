VERSION 5.00
Begin VB.Form frm_VTA_CobranzaMonto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importe a pagar"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   Icon            =   "frm_VTA_CobranzaMonto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   60
      TabIndex        =   3
      Top             =   -60
      Width           =   4155
      Begin vbp_Ventas.ctlTextBox txtMonto 
         Height          =   315
         Left            =   2520
         TabIndex        =   0
         Top             =   420
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Alignment       =   1
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
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Importe a pagar S/. : "
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
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   457
         Width           =   1845
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   1680
      TabIndex        =   1
      Top             =   1140
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   3000
      TabIndex        =   2
      Top             =   1140
      Width           =   1215
   End
End
Attribute VB_Name = "frm_VTA_CobranzaMonto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
On Error GoTo Control
cmdCancelar_Click

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control
    Select Case KeyCode
        Case vbKeyReturn
            
        Case vbKeyEscape
            cmdCancelar_Click
    End Select

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName

End Sub

