VERSION 5.00
Begin VB.Form frm_ADM_EntregaProducto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Cantidad"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlTextBox txtCantidad 
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Tipo            =   3
      Alignment       =   2
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
   Begin VB.ComboBox cboFecVen 
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox cboLote 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "F. Vencim."
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Lote"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label lblDescripcion 
      Caption         =   "Label2"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripción"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   840
   End
   Begin VB.Label lblCodigo 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frm_ADM_EntregaProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub carga(ByVal rs As oraDynaset)
lblCodigo.Caption = "" & rs(11).Value
lblDescripcion.Caption = "" & rs("DES_PRODUCTO").Value
cboLote.Clear
While Not rs.EOF
    cboLote.AddItem "" & rs("NUM_LOTE").Value
    rs.MoveNext
Wend
    Me.Show vbModal
End Sub
