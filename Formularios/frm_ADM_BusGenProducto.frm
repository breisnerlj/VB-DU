VERSION 5.00
Begin VB.Form frm_ADM_BusGenProducto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Productos"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlGrilla grdProductos 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5953
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
End
Attribute VB_Name = "frm_ADM_BusGenProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strCodProdGen As String
Public strDesProdGen As String
Dim objProducto As New clsProducto
Dim objVenta As New clsVenta


Public Sub LoadForm(ByVal vstrCodProducto As String, _
                    ByVal vstrCodLocal As String)

    SeteaGrilla
    Set grdProductos.DataSource = objProducto.ConsultaLocal("", _
                                                            "", _
                                                            "", _
                                                            "", _
                                                            "", _
                                                            "", _
                                                            vstrCodProducto, _
                                                            "", _
                                                            vstrCodLocal, _
                                                            "1")

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
    Unload Me
End Select
End Sub

Private Sub grdProductos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grdProductos_DblClick
    End If
End Sub

Private Sub grdProductos_DblClick()
    On Error GoTo CtrlErr
    If grdProductos.ApproxCount <= 0 Then Exit Sub
    'objVenta.CodProducto_BusGen = grdProductos.Columns("DESCRIPCION").Value
    'objVenta.DesProducto_BusGen = grdProductos.Columns("CODIGO").Value
strCodProdGen = grdProductos.Columns("CODIGO").Value
strDesProdGen = grdProductos.Columns("DESCRIPCION").Value
    Unload Me
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub SeteaGrilla()
On Error GoTo handle
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim Columna As TrueDBGrid70.Column
                      
    arrCampos = Array("CODIGO", "DESCRIPCION")
    arrCaption = Array("Código", "Descripción")
    arrAncho = Array(900, 4500)
    arrAlineacion = Array(dbgCenter, dbgLeft)
                          
    grdProductos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub


