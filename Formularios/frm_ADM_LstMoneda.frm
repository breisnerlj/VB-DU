VERSION 5.00
Begin VB.Form frm_ADM_LstMoneda 
   Caption         =   "Mantenimiento de Moneda"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   Icon            =   "frm_ADM_LstMoneda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   6585
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
   Begin vbp_Ventas.ctlGrilla grdMoneda 
      Height          =   3900
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   6879
   End
End
Attribute VB_Name = "frm_ADM_LstMoneda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objMoneda As New clsMoneda
Public pBlnMuestra As Boolean

Private Sub Form_Load()
    SetteaFormulario Me
    SeteaGrilla
    Set grdMoneda.DataSource = objMoneda.Lista
    
End Sub

Sub SeteaGrilla()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_MONEDA", "DES_MONEDA", "SMB_MONEDA", _
                      "DES_LG_MONEDA", "FLG_ACTIVO")
                      
    arrCaption = Array("Codigo", "Moneda", "Simbolo", _
                       "Longitud", "Activo")
                       
    arrAncho = Array(900, 1800, 900, _
                     1800, 900)
                     
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft)
                          
    grdMoneda.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub


Private Sub ToolBar_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
        Select Case Index
            Case "1"
                pBlnMuestra = False
                frm_ADM_Moneda.Show
            Case "2"
                pBlnMuestra = True
                frm_ADM_Moneda.Show
            Case "3"
                Set grdMoneda.DataSource = objMoneda.Lista
            Case "4"
                Set grdMoneda.DataSource = objMoneda.Lista
            
            Case "5"
                 grdMoneda.MostrarImprimir
                    
            Case "6"
                grdMoneda.MostrarExcel
                 
            Case "7"
                grdMoneda.MostrarEmail
                
            Case "8"
                Unload Me
        End Select
End Sub
