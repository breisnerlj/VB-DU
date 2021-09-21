VERSION 5.00
Begin VB.Form frmAprobacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Codigo de Aprobacion"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   3540
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Codigo de Seguridad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton cmdValidar 
         Caption         =   "Validar Codigo"
         Height          =   615
         Left            =   1800
         TabIndex        =   1
         Top             =   960
         Width           =   1335
      End
      Begin vbp_Ventas.ctlTextBox txtCodigo 
         Height          =   495
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         Tipo            =   3
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   135
   End
End
Attribute VB_Name = "frmAprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ObjValidacion As New clsAprobacion
Dim strCodigoAutoriza As String
Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub cmdValidar_Click()
    If txtCodigo.Text = "" Then
        MsgBox "Debe ingresar el codigo de Autorización", vbCritical, App.ProductName
        txtCodigo.Focus
    Else
        strCodigoAutoriza = Trim(txtCodigo.Text)
        Unload Me
    End If
End Sub
Public Function Carga() As String
        Me.Show vbModal
        Carga = strCodigoAutoriza
End Function


