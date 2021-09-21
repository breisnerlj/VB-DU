VERSION 5.00
Begin VB.Form frm_DLV_Reporte0 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reportar falta de Stock"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Frame frame_producto 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   5895
      Begin VB.Label lblLaboratorio 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   1425
         Width           =   3855
      End
      Begin VB.Label lblDescripcion 
         Height          =   420
         Left            =   1680
         TabIndex        =   9
         Top             =   915
         Width           =   4020
      End
      Begin VB.Label lblCodigo 
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   420
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Laboratorio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1420
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   920
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   420
         Width           =   975
      End
   End
   Begin VB.Frame frame_stock 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.Label lbl_Stock 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lbl_strStock 
         Caption         =   "Stock del Producto: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frm_DLV_Reporte0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strLocal As String
Public codLocal As String
Public desProducto As String
Public codProducto As String
Public stockProducto As String
Public strLaboratorio As String
Public bool As Boolean
'Public vtaUnid As String
Dim objProducto As clsProducto
Dim objLocal As clsLocal

Private Sub btnAceptar_Click()
'On Error GoTo Err
    Dim response As String
    Dim codLocalPos As String
    Dim codProdPos As String
    Dim Tipo As String
    Set objProducto = New clsProducto
    
    codLocalPos = codLocal 'objLocal.GetCodPosu(codLocal)
    codProdPos = objProducto.GetCodPosu(codProducto)
    Tipo = "DLV"
    response = objProducto.ReporteProductoStockZero("001", codLocalPos, codProdPos, desProducto, "DLV", gvarUSUARIO)
    bool = True
    Unload Me
'Err:
'    Err.Raise Err.Number, "ReporteZero", Err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control
'    psub_KeyDownAplicacion KeyCode, Shift
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'    Debug.Print "ABC"
'End Sub

Private Sub Form_Load()
    lbl_Stock.Caption = stockProducto
    lblCodigo.Caption = codProducto
    lblDescripcion.Caption = desProducto
    lblLaboratorio.Caption = strLaboratorio
'    lblUnidad.Caption = vtaUnid
End Sub

