VERSION 5.00
Begin VB.Form frm_ADM_DetDocEmitidos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   Icon            =   "frm_ADM_DetDocEmitidos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Documentos Emitidos"
      Height          =   5055
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin vbp_Ventas.ctlGrilla grdDetDocEmitidos 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   8281
         Resalte         =   0   'False
      End
   End
End
Attribute VB_Name = "frm_ADM_DetDocEmitidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objDocumento As New clsDocumento
Dim TipoDocumento As String
Public Sub Carga_Form(ByVal vstrcia As String, _
                      ByVal vstrCodLocal As String, _
                      ByVal vstrCodTipoDoc As String, _
                      ByVal vstrNumDocIni As String, _
                      ByVal vstrNumDocFin As String, _
                      ByVal vstrCodLiquidacion As String)

    
    SeteaGrilla
    frm_ADM_DetDocEmitidos.Caption = Trim(vstrCodTipoDoc) & "  " & " de liquidación Nº" & "  " & Trim(frm_VTA_Detalle_Liquidacion.pCodLiq)
    Set grdDetDocEmitidos.DataSource = objDocumento.ListaDetDocEmitidos(vstrcia, _
                                                                        vstrCodLocal, _
                                                                        vstrCodTipoDoc, _
                                                                        vstrNumDocIni, _
                                                                        vstrNumDocFin, _
                                                                        vstrCodLiquidacion)
    TipoDocumento = vstrCodTipoDoc
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  
    arrCampos = Array("NUM_DOCUMENTO", _
                      "FCH_EMISION", _
                      "COD_ESTADO", _
                      "MODALIDAD", _
                      "CONVENIO", _
                      "CLIENTE", _
                      "RUC", _
                      "MTO_TOTAL")
                      
    arrCaption = Array("Documento", _
                       "Fecha", _
                       "Estado", _
                       "Modalidad", _
                       "Convenio", _
                       "Cliente", _
                       "Ruc", _
                       "Total")
    
    arrAncho = Array(1200, _
                     1200, _
                     700, _
                     1300, _
                     1500, _
                     1800, _
                     1200, _
                     900)
    
    arrAlineacion = Array(vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft)
    
    grdDetDocEmitidos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub grdDetDocEmitidos_DblClick()
    frm_ADM_PreviewDoc.Datos objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, TipoDocumento, grdDetDocEmitidos.DataSource("NUM_DOCUMENTO"), "", ""
End Sub
