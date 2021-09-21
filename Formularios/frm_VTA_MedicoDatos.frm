VERSION 5.00
Begin VB.Form frm_VTA_MedicoDatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda de Medicos"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   Icon            =   "frm_VTA_MedicoDatos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   435
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin vbp_Ventas.ctlGrilla GrdBusMedico 
         Height          =   4815
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   8493
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   435
      Left            =   4680
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   6000
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
   End
End
Attribute VB_Name = "frm_VTA_MedicoDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objMedico As New clsMedico

Private Sub cmdNuevo_Click()
    
    Call frmGrabaMedico.Datos("", "", "", "", "", "", "0", "", "-(Nuevo)")
    Call ListarMedicos
    
End Sub

Private Sub Form_Activate()
'    GrdBusMedico.SetFocus
    Call ListarMedicos
End Sub

Private Sub Form_Load()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    arrCampos = Array("COD_MEDICO", "NOM_MEDICO", "DES_DIRECCION", "NUM_CMP")
    arrCaption = Array("Codigo", "Medico", "Dirección", "Nº CMP")
    arrAncho = Array(1000, 2200, 2500, 800)
    arrAlineacion = Array(vbAlignNone, vbAlignNone, vbAlignLeft, vbAlignNone)
    GrdBusMedico.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub

Private Sub cmdAceptar_Click()
    Call GrdBusMedico_DblClick
    Unload Me
End Sub

Private Sub GrdBusMedico_DblClick()
    frm_VTA_RecetarioM.TxtMedico.Text = "" & GrdBusMedico.Columns("NUM_CMP").Value
    frm_VTA_RecetarioM.LblMedico.Caption = "" & GrdBusMedico.Columns("NOM_MEDICO").Value
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub GrdBusMedico_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            GrdBusMedico_DblClick
    End Select
    
    
End Sub


Private Sub ListarMedicos()
    Set GrdBusMedico.DataSource = objMedico.Lista(frm_VTA_RecetarioM.pstrDatoMedico)
End Sub
