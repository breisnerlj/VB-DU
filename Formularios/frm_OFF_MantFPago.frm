VERSION 5.00
Begin VB.Form frm_OFF_MantFPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frm_OFF_MantFPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   2175
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4575
      Begin vbp_Ventas.ctlTextBox txtDescripcion 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   960
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Tipo            =   2
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
      Begin vbp_Ventas.ctlTextBox txtCodigo 
         Height          =   375
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Tipo            =   3
         MaxLength       =   1
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
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1050
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   450
         Width           =   540
      End
   End
End
Attribute VB_Name = "frm_OFF_MantFPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCodFormaPago As String
Public strDesFormaPago As String
Public bolCancelar As Boolean



Private Sub cmdAceptar_Click()

On Error GoTo Handle
    strCodFormaPago = txtCodigo.Text
    strDesFormaPago = txtDescripcion.Text
    bolCancelar = False
    Unload Me
Exit Sub

Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub cmdCancelar_Click()
    bolCancelar = True
    Unload Me
End Sub

Private Sub Form_Load()

bolCancelar = True

End Sub



Public Sub Datos(ByVal pstrCodFormaPago As String, _
            ByVal pstrDesFormaPago As String, _
            ByVal pstrCaption As String)


On Error GoTo Handle

    txtCodigo.Text = pstrCodFormaPago
    txtDescripcion.Text = pstrDesFormaPago
    Me.Caption = pstrCaption
    
Exit Sub

Handle:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


