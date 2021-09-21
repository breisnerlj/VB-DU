VERSION 5.00
Begin VB.Form frm_ADM_ProductoDatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relacion de Productos"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   9435
   ControlBox      =   0   'False
   Icon            =   "frm_ADM_ProductoDatos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9405
      Begin vbp_Ventas.ctlGrilla grdBusProducto 
         Height          =   3735
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6588
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
End
Attribute VB_Name = "frm_ADM_ProductoDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objProducto As New clsProducto
Private strDato As String
Private varSalida(1 To 3) As String

Private Property Get Dato() As String
    Dato = strDato
End Property

Public Property Let Dato(ByVal vNewValue As String)
    strDato = vNewValue
End Property

Public Property Get Salida(ByVal Index As Integer) As Variant
    Salida = varSalida(Index)
End Property

Private Property Let Salida(ByVal Index As Integer, ByVal vNewValue As Variant)
    varSalida(Index) = vNewValue
End Property

Private Sub Form_Load()
    On Error GoTo CtrlErr
    SeteaGrilla
    Set grdBusProducto.DataSource = objProducto.ConsultaLocal("", _
                                                              "", _
                                                              "", _
                                                              "", _
                                                              "", _
                                                              "", _
                                                              Dato, _
                                                              "ACT", _
                                                              objUsuario.CodigoLocal, _
                                                              "1")
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo CtrlErr
    
    Select Case KeyCode
        Case vbKeyEscape
            Salida(1) = ""
            Salida(2) = ""
            Salida(3) = ""
            Unload Me
    End Select
    
    Exit Sub
    
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objProducto = Nothing
End Sub

Private Sub grdBusProducto_DblClick()
    
    On Error GoTo CtrlErr
    
    If grdBusProducto.ApproxCount <= 0 Then Exit Sub
    Salida(1) = grdBusProducto.Columns("CODIGO").Value
    Salida(2) = grdBusProducto.Columns("DESCRIPCION").Value
    Salida(3) = grdBusProducto.Columns("ESTADO").Value
    Unload Me
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub grdBusProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo CtrlErr
    
    Select Case KeyCode
        Case vbKeyReturn
            grdBusProducto_DblClick
    End Select
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub SeteaGrilla()

  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim columna As TrueDBGrid70.Column

  
                      
                      
    arrCampos = Array("CODIGO", "DESCRIPCION", _
                      "ESTADO", "STOCK")
                      
                      
    arrCaption = Array("Código", "Descripción", _
                       "Estado", "Stock")
    
    arrAncho = Array(1000, 5000, _
                     1050, 1550)
                     
    arrAlineacion = Array(dbgCenter, dbgLeft, _
                          dbgCenter, dbgCenter)
                          
    grdBusProducto.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    
    For Each columna In grdBusProducto.Columns
        columna.AllowSizing = False
    Next
    
End Sub

Public Function ConsultaProd(vstrDato As String) As oraDynaset
    Set ConsultaProd = gclsOracle.FN_Cursor("BTLPROD.PKG_SMM.FN_CONSULTA_PROD", 0, vstrDato)
End Function

