VERSION 5.00
Begin VB.Form frm_OFF_MantMoneda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frm_OFF_MantMoneda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   2055
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   4575
      Begin vbp_Ventas.ctlTextBox txtSmbMoneda 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
      Begin vbp_Ventas.ctlTextBox txtDescripcion 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
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
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Simbolo:"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   1380
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   900
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   420
         Width           =   540
      End
   End
End
Attribute VB_Name = "frm_OFF_MantMoneda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCodMoneda As String
Public strDesMoneda As String
Public strSmbMoneda As String
Public bolCancelar As Boolean



Private Sub cmdAceptar_Click()

On Error GoTo Handle
    strCodMoneda = txtCodigo.Text
    strDesMoneda = txtDescripcion.Text
    strSmbMoneda = txtSmbMoneda.Text
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



Public Sub Datos(ByVal pstrCodMoneda As String, _
            ByVal pstrDesMoneda As String, _
            ByVal pstrSmbMoneda As String, _
            ByVal pstrCaption As String)


On Error GoTo Handle

    txtCodigo.Text = pstrCodMoneda
    txtDescripcion.Text = pstrDesMoneda
    txtSmbMoneda.Text = pstrSmbMoneda
    Me.Caption = pstrCaption
    
Exit Sub

Handle:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub
