VERSION 5.00
Begin VB.Form frm_ADM_Sobrantes 
   Caption         =   "Reporte de Faltantes/Sobrantes"
   ClientHeight    =   6750
   ClientLeft      =   4065
   ClientTop       =   2640
   ClientWidth     =   10185
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10185
   Begin VB.CommandButton Command1 
      Caption         =   "[F1] Prox. Vencimiento"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "[Esc] Salir"
      Height          =   375
      Left            =   9000
      TabIndex        =   7
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdF11 
      Caption         =   "[F11] Imprimir"
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "[F2] Deterioradas"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sobrantes"
      Height          =   3015
      Left            =   0
      TabIndex        =   2
      Top             =   3120
      Width           =   10215
      Begin vbp_Ventas.ctlGrilla ctlgrdSobrantes 
         Height          =   2415
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4260
         Resalte         =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Faltantes"
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin vbp_Ventas.ctlGrilla ctlgrdFaltantes 
         Height          =   2415
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4260
         Resalte         =   0   'False
      End
   End
   Begin vbp_Ventas.ctlGrilla ctlGrilla1 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4260
      Resalte         =   0   'False
   End
End
Attribute VB_Name = "frm_ADM_Sobrantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objEntrega As New clsEntrega
Public idEntrega As String

Private Sub cmdEsc_Click()
    frm_ADM_Entrega.Form_Load
    frm_ADM_Entrega.grdRecepcion.DataSource.FindFirst "ID_ENTREGA='" & Trim(idEntrega) & "'"
    Unload Me
End Sub

Private Sub cmdF11_Click()
    Me.ctlGrilla1.MostrarImprimir
End Sub

Private Sub cmdF2_Click()
 objEntrega.LimpiaDetDev (idEntrega)
 frm_ADM_ProdDevolucion.strIdEntrega = idEntrega
 frm_ADM_ProdDevolucion.strTipo = "DET"
 frm_ADM_ProdDevolucion.Show vbModal
End Sub

Private Sub Command1_Click()
 objEntrega.LimpiaDetDev (idEntrega)
 frm_ADM_ProdDevolucion.strIdEntrega = idEntrega
 frm_ADM_ProdDevolucion.strTipo = "VCMTO"
 frm_ADM_ProdDevolucion.Show vbModal
End Sub



Private Sub ctlgrdFaltantes_DblClick()
'        frm_ADM_ProdLoteSobrFalt.lblProducto.Caption = "" & Me.ctlGrilla1.Columns(1).Value
'        'Agregar el indicador a la funcion de tomas
'        'MsgBox Me.ctlgrdFaltantes.Columns(5).Value
'        'Modificado por MLEVANO 23/11/2012
'        If Me.ctlgrdFaltantes.Columns(5).Value = "" Then
'            Call frm_ADM_ProdLoteSobrFalt.cargaDatos(idEntrega, Me.ctlGrilla1.Columns(0).Value)
'        End If
End Sub

Private Sub ctlgrdSobrantes_DblClick()
        frm_ADM_ProdLoteSobrFalt.lblProducto.Caption = "" & Me.ctlgrdSobrantes.Columns(1).Value
        frm_ADM_ProdLoteSobrFalt.lblNEntrega.Caption = "" & Me.ctlgrdSobrantes.Columns(7).Value
        frm_ADM_ProdLoteSobrFalt.lblGuia.Caption = "" & Me.ctlgrdSobrantes.Columns(6).Value
        
        'Agregar el indicador a la funcion de tomas
        'Tomas modifico su funcion... buscar donde agregar el indicador
        'MsgBox Me.ctlgrdFaltantes.Columns(5).Value
        'por MLEVANO 23/11/2012
        If Me.ctlgrdSobrantes.Columns(5).Value = "" And Me.ctlgrdSobrantes.Columns(6).Value = "" Then
            Call frm_ADM_ProdLoteSobrFalt.cargaDatos(idEntrega, Me.ctlgrdSobrantes.Columns(0).Value)
            Unload Me
        End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    cmdEsc_Click
End If
If KeyCode = vbKeyF11 Then
    cmdF11_Click
End If
If KeyCode = vbKeyF2 Then
    cmdF2_Click
End If
If KeyCode = vbKeyF1 Then
    Command1_Click
End If
End Sub

Private Sub Form_Load()
Me.cargaDetalle "Faltante", Me.ctlgrdFaltantes
Me.cargaDetalle "Sobrante", Me.ctlgrdSobrantes
Me.cargaDetalle "@", Me.ctlGrilla1
Dim odyn As oraDynaset
Dim flgV, flgD As String
Set odyn = objEntrega.BuscaFlgGeneraGuia(idEntrega)
If odyn.RecordCount > 0 Then
    flgD = "" & odyn(0).Value
    flgV = "" & odyn(1).Value
End If
If flgD = "1" Then
    Me.cmdF2.Enabled = False
Else
    Me.cmdF2.Enabled = True
End If
If flgV = "1" Then
    Me.Command1.Enabled = False
Else
    Me.Command1.Enabled = True
End If
End Sub

Sub cargaDetalle(Tipo As String, grilla As ctlGrilla)
'Set Me.ctlgrdFaltantes.DataSource = objEntrega.ListaDiferencias(identrega, tipo)
Set grilla.DataSource = objEntrega.ListaDiferencias(idEntrega, Tipo)
SeteaGrilla grilla
End Sub

Sub SeteaGrilla(grilla As ctlGrilla)
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim arrFoco As Variant

    arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO", "DES_LABORATORIO", "TIPO", "DIFERENCIA", "LOTE", "NUM_FACTURA_SAP", "NUM_ENTREGA", "NUM_GUIA")
    arrCaption = Array("Codigo", "Descripcion", "Laboratorio", "Condicion", "Dif.", "Lote", "Factura", "NEntrega", "NGuia")
    arrAncho = Array(1000, 2500, 2500, 1000, 500, 800, 800, 800, 800)
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, dbgCenter, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight)

    grilla.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grilla.Columns(0).Merge = False
    grilla.Columns(1).Merge = False
    grilla.Columns(2).Merge = False
    grilla.Columns(3).Merge = False
    grilla.Columns(4).Merge = False
    grilla.Columns(5).Merge = False
    grilla.Columns(6).Merge = False
    grilla.Columns(7).Merge = False
    grilla.Columns(8).Merge = False
    
    grilla.Columns(1).BackColor = vbInfoBackground
    grilla.Columns(0).BackColor = vbInfoBackground
    grilla.Columns(7).Visible = False
    grilla.Columns(8).Visible = False

End Sub

Public Sub recibe()
    Unload frm_ADM_ProdDevGuia
    Unload frm_ADM_ProdDevolucion
    Form_Load
End Sub
