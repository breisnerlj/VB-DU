VERSION 5.00
Begin VB.Form frm_VTA_Vanttive 
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmCliente 
      Caption         =   "Cliente"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   6240
         TabIndex        =   17
         Text            =   "Text10"
         Top             =   2650
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   6240
         TabIndex        =   16
         Text            =   "Text9"
         Top             =   2230
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   6240
         TabIndex        =   15
         Text            =   "Text8"
         Top             =   1810
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Text            =   "Text7"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   1080
         TabIndex        =   13
         Text            =   "Text6"
         Top             =   2650
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   2230
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1080
         TabIndex        =   11
         Text            =   "12345678901"
         Top             =   1810
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   500
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "frm_VTA_Vanttive.frx":0000
         Top             =   1265
         Width           =   6375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   375
         Left            =   6960
         TabIndex        =   9
         Top             =   300
         Width           =   480
      End
      Begin VB.TextBox Text2 
         Height          =   500
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "frm_VTA_Vanttive.frx":0006
         Top             =   720
         Width           =   6375
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   300
         Width           =   5775
      End
      Begin VB.Label Label10 
         Caption         =   "Línea de Crédito Disponible"
         Height          =   255
         Left            =   4150
         TabIndex        =   21
         Top             =   2715
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Monto en Documentos por Cobrar"
         Height          =   255
         Left            =   3720
         TabIndex        =   20
         Top             =   2320
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Teléfono"
         Height          =   255
         Left            =   5500
         TabIndex        =   19
         Top             =   1875
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "DNI"
         Height          =   255
         Left            =   3000
         TabIndex        =   18
         Top             =   1870
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "F. Pago"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   2720
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Límite Credito"
         Height          =   375
         Left            =   430
         TabIndex        =   5
         Top             =   2210
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "NIT"
         Height          =   255
         Left            =   650
         TabIndex        =   4
         Top             =   1870
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Dirección"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1400
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre / R. Social"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   750
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Dato"
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   375
         Width           =   375
      End
   End
End
Attribute VB_Name = "frm_VTA_Vanttive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

