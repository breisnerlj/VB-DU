VERSION 5.00
Begin VB.Form frm_VTA_Detalle_Liquidacion 
   Caption         =   "Detalle de Liquidación"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   Icon            =   "frm_VTA_Detalle_Liquidacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdDetLiquidacion 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Actualizar"
      Height          =   555
      Left            =   5760
      Picture         =   "frm_VTA_Detalle_Liquidacion.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Detalle de Liquidación"
      Top             =   6200
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalle de Venta por Modalidad"
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   6975
      Begin vbp_Ventas.ctlGrilla ctlGrdVentaxMod 
         Height          =   2775
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4895
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Secuencia por Documento"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin vbp_Ventas.ctlGrilla ctlGrdSecDoc 
         Height          =   2415
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4260
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   6960
      Y1              =   2880
      Y2              =   2880
   End
End
Attribute VB_Name = "frm_VTA_Detalle_Liquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objLiquidacion As New clsLiquidacion
Public pCodLiq As String
Dim strSecxDoc As String
Dim strVtaxMod As String

Private Sub CmdDetLiquidacion_Click()
    On Error GoTo CtrlErr
    Set ctlGrdSecDoc.DataSource = objLiquidacion.Det_Liquidacion_Venta(objUsuario.CodigoEmpresa, _
                                                                           objUsuario.CodigoLocal, _
                                                                           pCodLiq, _
                                                                           strSecxDoc)
                                                                           
    Set ctlGrdVentaxMod.DataSource = objLiquidacion.Det_Liquidacion_Venta(objUsuario.CodigoEmpresa, _
                                                                              objUsuario.CodigoLocal, _
                                                                              pCodLiq, _
                                                                              strVtaxMod)
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error"
End Sub

Private Sub ctlGrdSecDoc_DblClick()
    If ctlGrdSecDoc.ApproxCount <= 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    frm_ADM_DetDocEmitidos.Carga_Form objUsuario.CodigoEmpresa, _
                                      objUsuario.CodigoLocal, _
                                      ctlGrdSecDoc.Columns("DOCUMENTO").Value, _
                                      ctlGrdSecDoc.Columns("SEC_MINIMO").Value, _
                                      ctlGrdSecDoc.Columns("SEC_MAXIMO").Value, _
                                      Trim(pCodLiq)
    frm_ADM_DetDocEmitidos.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
    Unload Me
End Select
End Sub

Private Sub Form_Load()
On Error GoTo Control

    setteaFormulario Me
    SeteaGrilla_SecxDoc
    SeteaGrilla_VtaxMod
    With frm_VTA_Detalle_Liquidacion
    
        .Caption = "Detalle de Liquidación" & "  " & "Nº" & "  " & Trim(pCodLiq)
        
        strSecxDoc = objLiquidacion.SecxDoc
        strVtaxMod = objLiquidacion.VtaxModalidad
        
        Set ctlGrdSecDoc.DataSource = objLiquidacion.Det_Liquidacion_Venta(objUsuario.CodigoEmpresa, _
                                                                           objUsuario.CodigoLocal, _
                                                                           pCodLiq, _
                                                                           strSecxDoc)
                                                                           
        Set ctlGrdVentaxMod.DataSource = objLiquidacion.Det_Liquidacion_Venta(objUsuario.CodigoEmpresa, _
                                                                              objUsuario.CodigoLocal, _
                                                                              pCodLiq, _
                                                                              strVtaxMod)
    End With

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub SeteaGrilla_SecxDoc()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  
    arrCampos = Array("DOCUMENTO", _
                      "SEC_MINIMO", _
                      "SEC_MAXIMO")
                      
    arrCaption = Array("Tipo Doc.", _
                       "Sec Minímo", _
                       "Sec Maxímo")
    
    arrAncho = Array(1500, _
                     1800, _
                     1800)
    
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft, vbAlignLeft, _
                          vbAlignLeft)
    
    ctlGrdSecDoc.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
End Sub

Private Sub SeteaGrilla_VtaxMod()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_MODALIDAD_VENTA", _
                      "DES_MODALIDAD_VENTA", _
                      "TOTAL")
                      
    arrCaption = Array("Codigo", _
                       "Modalidad", _
                       "Total")
    
    arrAncho = Array(900, _
                     2200, _
                     1000)
    
    arrAlineacion = Array(vbAlignLeft, _
                          vbAlignLeft, _
                          dbgRight)
    
    ctlGrdVentaxMod.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    ctlGrdVentaxMod.Columns("TOTAL").NumberFormat = "Standard"

End Sub
