VERSION 5.00
Begin VB.Form frm_VTA_Liquidacion_FP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidación de Formas de Pagos"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   Icon            =   "frm_VTA_Liquidacion_FP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdActualiza 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Actualizar"
      Height          =   555
      Left            =   5880
      Picture         =   "frm_VTA_Liquidacion_FP.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Detalle de Liquidación"
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      Begin vbp_Ventas.ctlGrilla grdDetLiqFP 
         Height          =   5655
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   9975
      End
   End
   Begin VB.Label lblNumItems 
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      Height          =   225
      Left            =   1440
      TabIndex        =   7
      Top             =   6465
      Width           =   855
   End
   Begin VB.Label lblEtiqNumItems 
      AutoSize        =   -1  'True
      Caption         =   "Items = >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   6480
      Width           =   795
   End
   Begin VB.Label lblEtiqueta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   6120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label LblTotSist 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0.00"
      Height          =   225
      Left            =   4080
      TabIndex        =   4
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label51 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total =>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   3
      Top             =   6480
      Width           =   720
   End
End
Attribute VB_Name = "frm_VTA_Liquidacion_FP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim odyn As oraDynaset
Public pLiquidacion As String
Public pCodFormaPago As String
Public pCodHijo As String

Private Sub Form_Load()
    setteaFormulario Me
    SeteaGrilla
     
    On Error GoTo CtrlErr
    grdDetLiqFP.Columns(0).Visible = False: grdDetLiqFP.Columns(1).Visible = False ': grdDetLiqFP.Columns(2).Visible = False
    Me.Caption = "Liquidación por Forma de Pago - Nº" & " " & pLiquidacion
    
    Set odyn = objLiquidacion.ListaDetLiqFormaPago(objUsuario.CodigoEmpresa, _
                                                   objUsuario.CodigoLocal, _
                                                   pLiquidacion, _
                                                   pCodFormaPago, _
                                                   pCodHijo)
    Call Total(odyn)
    Set grdDetLiqFP.DataSource = odyn
                                                       
    Exit Sub
CtrlErr:
    MsgBox "No tiene registros de formas de pago", vbCritical, App.FileDescription
End Sub

Private Sub cmdActualiza_Click()
    On Error GoTo CtrlErr
    Screen.MousePointer = vbHourglass
    Set odyn = objLiquidacion.ListaDetLiqFormaPago(objUsuario.CodigoEmpresa, _
                                                   objUsuario.CodigoLocal, _
                                                   pLiquidacion, _
                                                   pCodFormaPago, _
                                                   pCodHijo)
    Call Total(odyn)
    Set grdDetLiqFP.DataSource = odyn
    Screen.MousePointer = vbDefault
    Exit Sub
CtrlErr:
    MsgBox "No tiene registros de formas de pago", vbCritical, App.FileDescription

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub grdDetLiqFP_DblClick()
    If grdDetLiqFP.ApproxCount <= 0 Then Exit Sub
    If grdDetLiqFP.Columns("COD_TIPO_DOCUMENTO").Value = objUsuario.TipoDocBol Or _
       grdDetLiqFP.Columns("COD_TIPO_DOCUMENTO").Value = objUsuario.TipoDocFac Or _
       grdDetLiqFP.Columns("COD_TIPO_DOCUMENTO").Value = objVenta.TipoDocTKB Or _
       grdDetLiqFP.Columns("COD_TIPO_DOCUMENTO").Value = objVenta.TipoDocTKF Then
       
        Screen.MousePointer = vbHourglass
        frm_ADM_PreviewDoc.Datos objUsuario.CodigoEmpresa, _
                                 objUsuario.CodigoLocal, _
                                 grdDetLiqFP.Columns("COD_TIPO_DOCUMENTO").Value, _
                                 grdDetLiqFP.Columns("NUM_DOCUMENTO").Value, _
                                 "", ""
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_FORMA_PAGO", _
                      "COD_HIJO", _
                      "DES_HIJO", _
                      "COD_TIPO_DOCUMENTO", _
                      "NUM_DOCUMENTO", _
                      "TOTAL_DOC", _
                      "TOTAL_PAGADO")
                      
    arrCaption = Array("CodFPago", _
                       "CodHijo", _
                       "Forma Pago", _
                       "Tipo Doc.", _
                       "Numero", _
                       "Total Doc.", _
                       "Total Pagado")
    
    arrAncho = Array(800, _
                     800, _
                     1800, _
                     1000, _
                     1100, _
                     1200, _
                     1200)
    
    arrAlineacion = Array(vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft, _
                          vbAlignLeft)
    
    grdDetLiqFP.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdDetLiqFP.Columns(0).Merge = True
    grdDetLiqFP.Columns(1).Merge = True
    grdDetLiqFP.Columns(2).Merge = True
End Sub

Private Sub Total(ByVal odynClonSist As oraDynaset)
Dim dblTotSist As Double
Dim strEtiquetaTotal As String
Dim intNumItems As Integer

    dblTotSist = 0
    intNumItems = 0
    odynClonSist.MoveFirst
        strEtiquetaTotal = odynClonSist("DES_HIJO").Value
        While Not odynClonSist.EOF
              dblTotSist = dblTotSist + odynClonSist("TOTAL_PAGADO").Value
              intNumItems = intNumItems + 1
           odynClonSist.MoveNext
        Wend
    odynClonSist.MoveFirst
    'lblEtiqueta.Caption = strEtiquetaTotal
    LblTotSist.Caption = Format(dblTotSist, "#,###,##0.00")
    lblNumItems.Caption = intNumItems

End Sub

'Private Sub grdDetLiqFP_RegistroSeleccionado(ByVal DatoColumna0 As String)
'    lblEtiqueta.Caption = grdDetLiqFP.Columns(2).Value
'End Sub
