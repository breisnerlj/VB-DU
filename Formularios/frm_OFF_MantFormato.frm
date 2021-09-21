VERSION 5.00
Begin VB.Form frm_OFF_MantFormato 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frm_OFF_MantFormato.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   2535
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   4575
      Begin vbp_Ventas.ctlTextBox txtCtdAncho 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Tipo            =   4
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
      Begin vbp_Ventas.ctlTextBox txtDesFormato 
         Height          =   315
         Left            =   1440
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
      Begin vbp_Ventas.ctlTextBox txtCodFormato 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         Tipo            =   3
         MaxLength       =   3
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
      Begin vbp_Ventas.ctlTextBox txtCtdAlto 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Tipo            =   4
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Alto:"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   1860
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ancho:"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   1380
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   900
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   420
         Width           =   540
      End
   End
End
Attribute VB_Name = "frm_OFF_MantFormato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCodFormato As String
Public strDesFormato As String
Public strCtdAncho As String
Public strCtdAlto As String
Public bolCancelar As Boolean



Private Sub cmdAceptar_Click()

On Error GoTo Handle
    strCodFormato = txtCodFormato.Text
    strDesFormato = txtDesFormato.Text
    strCtdAncho = txtCtdAncho.Text
    strCtdAlto = txtCtdAlto.Text
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



Public Sub Datos(ByVal pstrCodFormato As String, _
            ByVal pstrDesFormato As String, _
            ByVal pintCtdAncho As String, _
            ByVal pintCtdAlto As String, _
            ByVal pstrCaption As String)


On Error GoTo Handle

    txtCodFormato.Text = pstrCodFormato
    txtDesFormato.Text = pstrDesFormato
    txtCtdAncho.Text = pintCtdAncho
    txtCtdAlto.Text = pintCtdAlto
    Me.Caption = pstrCaption
    
Exit Sub

Handle:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


