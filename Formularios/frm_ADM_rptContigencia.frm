VERSION 5.00
Begin VB.Form frm_ADM_rptContigencia 
   BorderStyle     =   0  'None
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlToolBar ToolBar 
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
   Begin vbp_Ventas.ctlGrilla grdLog 
      Height          =   6420
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   11324
   End
End
Attribute VB_Name = "frm_ADM_rptContigencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strSecError As String
Dim objContingencia As New cls_OFF_Sincronizacion
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
ToolBar.Buttons(1).Visible = False
ToolBar.Buttons(2).Visible = False
ToolBar.Buttons(3).Visible = False
ToolBar.Buttons(4).Visible = False
   setteaFormulario Me
    SeteaGrilla
    Set grdLog.DataSource = objContingencia.ListaErrores(strSecError)
End Sub

Private Sub ToolBar_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    Select Case Index
        Case 1
               grdLog.MostrarImprimir
            Case 2
               grdLog.MostrarExcel
            Case 3
               grdLog.MostrarEmail
             Case 4
                Unload Me
    End Select
End Sub
Sub SeteaGrilla()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
  
    arrCampos = Array("ITEM_ERROR", "DES_ERROR")
                      
    arrCaption = Array("N°", "Error")
                       
    arrAncho = Array(500, 6000)
                     
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft)
                          
    grdLog.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub

