VERSION 5.00
Begin VB.Form frm_DLV_LstTipoDireccion 
   Caption         =   "Lista de Direcciones"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   Icon            =   "frm_DLV_LstTipoDireccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   6600
   Begin vbp_Ventas.ctlGrilla grdDireccDlv 
      Height          =   3900
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   6580
      _ExtentX        =   11615
      _ExtentY        =   6879
   End
   Begin vbp_Ventas.ctlToolBar ToolBar_AsisMoto 
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1058
      ModoBotones     =   3
      EnabledEfecto   =   0   'False
   End
End
Attribute VB_Name = "frm_DLV_LstTipoDireccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objDirecc As New clsDireccDLV

Private Sub Form_Load()
    SetteaFormulario Me
    SeteaGrila
    Set grdDireccDlv.DataSource = objDirecc.ListaDirecc
        
End Sub

Sub SeteaGrila()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_TIPO_DIRECCION", "DES_TIPO_DIRECCION", "FLG_ACTIVO")
                      
    arrCaption = Array("Codigo", "Descripción", "Activo")
                       
    arrAncho = Array(900, 2800, 800)
                     
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, vbAlignLeft)
                          
    grdDireccDlv.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
End Sub

Private Sub ToolBar_AsisMoto_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    Select Case Index
        Case 1
            'Nuevo'
            frm_DLV_GrbTipoDireccion.pstrCod = ""
            frm_DLV_GrbTipoDireccion.Show vbModal
        Case 2
            'Modificar'
            If Not grdDireccDlv.ApproxCount <= 0 Then
                frm_DLV_GrbTipoDireccion.pstrCod = grdDireccDlv.Columns("COD_TIPO_DIRECCION").Value
                frm_DLV_GrbTipoDireccion.pstrDes = grdDireccDlv.Columns("DES_TIPO_DIRECCION").Value
                frm_DLV_GrbTipoDireccion.pstrActivo = grdDireccDlv.Columns("FLG_ACTIVO").Value

                frm_DLV_GrbTipoDireccion.Show vbModal
            End If
        Case 4
            Set grdDireccDlv.DataSource = objDirecc.ListaDirecc
            SeteaGrila
        Case 5
            grdDireccDlv.MostrarImprimir
        Case 6
            grdDireccDlv.MostrarExcel
        Case 7
            grdDireccDlv.MostrarEmail
        Case 8
            Unload Me
        Case Else
            MsgBox "Esta opción se encuentra Deshabilitada", vbExclamation, App.ProductName
    End Select
End Sub
