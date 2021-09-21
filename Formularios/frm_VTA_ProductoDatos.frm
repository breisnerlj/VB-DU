VERSION 5.00
Begin VB.Form frm_VTA_ProductoDatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relacion de Productos"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   9435
   Icon            =   "frm_VTA_ProductoDatos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   9435
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9405
      Begin vbp_Ventas.ctlGrilla GrdBusProducto 
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6588
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
End
Attribute VB_Name = "frm_VTA_ProductoDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objKardex As New clsKardex
Dim objProducto As New clsProducto
Public pCodProd As String

Private Sub Form_Load()
    
    'Set GrdBusProducto.DataSource = objKardex.ListaProducto(frm_VTA_RepKardex.pstrCodProd)
    SeteaGrilla
    Set GrdBusProducto.DataSource = objProducto.ConsultaLocal("", _
                                                              "", _
                                                              "", _
                                                              "", _
                                                              "", _
                                                              "", _
                                                              frm_VTA_RepKardex.pstrCodProd, _
                                                              "", _
                                                              objUsuario.CodigoLocal, _
                                                              "1")

    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub grdBusProducto_DblClick()
    If GrdBusProducto.ApproxCount <= 0 Then Exit Sub
    pCodProd = GrdBusProducto.Columns(0).Value
    frm_VTA_RepKardex.lblCod_Producto.Caption = GrdBusProducto.Columns(0).Value
    frm_VTA_RepKardex.lblProducto.Caption = GrdBusProducto.Columns(1).Value
    frm_VTA_RepKardex.lblEstado.Caption = GrdBusProducto.Columns(2).Value
    frm_VTA_RepKardex.lblLab.Caption = objProducto.DevLabLineaProd(pCodProd, "0") 'GrdBusProducto.Columns(3).Value
    frm_VTA_RepKardex.lblLinea.Caption = objProducto.DevLabLineaProd(pCodProd, "1") 'GrdBusProducto.Columns(4).Value
    
    frm_VTA_RepKardex.lblCodSap = objProducto.DevCodSap(GrdBusProducto.Columns(0).Value)
    
    
    frm_VTA_RepKardex.Buscar
    Unload Me
End Sub

Private Sub grdBusProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            grdBusProducto_DblClick
    End Select
End Sub

Private Sub SeteaGrilla()
On Error GoTo handle
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim Columna As TrueDBGrid70.Column

  
                      
                      
    arrCampos = Array("CODIGO", "DESCRIPCION", _
                      "ESTADO", "STOCK")
                      
                      
    arrCaption = Array("Código", "Descripción", _
                       "Estado", "Stock")
    
    arrAncho = Array(1000, 5000, _
                     1050, 1550)
                     
    arrAlineacion = Array(dbgCenter, dbgLeft, _
                          dbgCenter, dbgCenter)
                          
    GrdBusProducto.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    
    
    
    
    For Each Columna In GrdBusProducto.Columns
        Columna.AllowSizing = False
    
    Next
    

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub


