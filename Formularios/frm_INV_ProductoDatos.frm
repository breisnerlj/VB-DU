VERSION 5.00
Begin VB.Form frm_INV_ProductoDatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relacion de Productos"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   Icon            =   "frm_INV_ProductoDatos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      Begin vbp_Ventas.ctlGrilla GrdBusProducto 
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6588
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
End
Attribute VB_Name = "frm_INV_ProductoDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public irsProducto As oraDynaset
Public istrCodProducto As String
Public istrDesProducto As String
Public iarrCabecera As Variant

Private Sub Form_Load()
    Dim i As Integer
    Dim lintAncho As Integer

    Set GrdBusProducto.DataSource = irsProducto
    
    istrCodProducto = ""
    istrDesProducto = ""
    
    For i = LBound(iarrCabecera) To UBound(iarrCabecera)
        If i <= GrdBusProducto.Columns.Count - 1 Then
            GrdBusProducto.Columns(i).Caption = iarrCabecera(i)
            lintAncho = Len(GrdBusProducto.Columns(i).Value)
            lintAncho = IIf(Len(iarrCabecera(i)) > lintAncho, Len(iarrCabecera(i)), lintAncho)
            GrdBusProducto.Columns(i).Width = 120 * lintAncho
        End If
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub GrdBusProducto_DblClick()
    If GrdBusProducto.ApproxCount <= 0 Then Exit Sub
    istrCodProducto = GrdBusProducto.Columns(0).Value
    istrDesProducto = GrdBusProducto.Columns(1).Value
    Unload Me
End Sub

Private Sub GrdBusProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            GrdBusProducto_DblClick
    End Select
End Sub


