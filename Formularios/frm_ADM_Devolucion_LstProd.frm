VERSION 5.00
Begin VB.Form frm_ADM_Devolucion_LstProd 
   Caption         =   "Lista de Productos"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlGrilla ctlGrilla1 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10398
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
End
Attribute VB_Name = "frm_ADM_Devolucion_LstProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objentrega As New clsEntrega
Dim strIdEntrega As String
Dim glMomento As String

Public Sub recibe(ByVal idEntrega As String, ByVal moment As String)
strIdEntrega = idEntrega
glMomento = moment
SeteaGrilla
Set Me.ctlGrilla1.DataSource = objentrega.ListaProdDev(strIdEntrega)
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
    arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO", "DES_LABORATORIO")
    arrCaption = Array("Codigo", "Producto", "Laboratorio")
    arrAncho = Array(800, 3500, 3500)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft)
    ctlGrilla1.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If

If KeyCode = vbKeyReturn Then
    'objentrega.GrabaDetDev strIdEntrega, "" & Me.ctlGrilla1.Columns(0).Value, "1", "", "", glMomento
    frm_ADM_ProdDevCant.lblProducto.Caption = "" & Me.ctlGrilla1.Columns(1).Value
    frm_ADM_ProdDevCant.idEntrega = strIdEntrega
    frm_ADM_ProdDevCant.codProducto = "" & Me.ctlGrilla1.Columns(0).Value
    frm_ADM_ProdDevCant.moment = glMomento
    frm_ADM_ProdDevCant.Show vbModal
    Unload Me
End If
End Sub

