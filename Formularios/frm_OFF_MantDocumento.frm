VERSION 5.00
Begin VB.Form frm_OFF_MantDocumento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frm_OFF_MantDocumento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   840
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   3135
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4575
      Begin vbp_Ventas.ctlTextBox txtAncho 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   2160
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Tipo            =   3
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
      Begin vbp_Ventas.ctlTextBox txtNumLineas 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   1680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Tipo            =   3
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
      Begin vbp_Ventas.ctlTextBox txtNumDoc 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Tipo            =   7
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
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
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
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Tipo            =   2
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ancho:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2220
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "# Líneas:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1740
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1260
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   780
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   300
         Width           =   540
      End
   End
End
Attribute VB_Name = "frm_OFF_MantDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCodDocumento As String
Public strDesDocumento As String
Public strNumDocumento As String
Public strNumLinDocumento As String
Public strAnchoDocumento As String
Public bolCancelar As Boolean



Private Sub cmdAceptar_Click()

On Error GoTo Handle
    strCodDocumento = txtCodigo.Text
    strDesDocumento = txtDescripcion.Text
    strNumDocumento = txtNumDoc.Text
    strNumLinDocumento = txtNumLineas.Text
    strAnchoDocumento = txtAncho.Text
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



Public Sub Datos(ByVal pstrCodDocumento As String, _
            ByVal pstrDesDocumento As String, _
            ByVal pstrNumDocumento As String, _
            ByVal pstrNumLinDocumento As String, _
            ByVal pstrAnchoDocumento As String, _
            ByVal pstrCaption As String)


On Error GoTo Handle

    txtCodigo.Text = pstrCodDocumento
    txtDescripcion.Text = pstrDesDocumento
    txtNumDoc.Text = pstrNumDocumento
    txtNumLineas.Text = pstrNumLinDocumento
    txtAncho.Text = pstrAnchoDocumento
    Me.Caption = pstrCaption
    
Exit Sub

Handle:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub



