VERSION 5.00
Begin VB.Form frm_VTA_ProveedorDatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda de Proveedores"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frm_VTA_ProveedorDatos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   6000
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
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
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin vbp_Ventas.ctlGrilla GrdBusProveedor 
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
End
Attribute VB_Name = "frm_VTA_ProveedorDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objProveedor As New clsProveedor

Private Sub Form_Activate()
    'GrdBusProveedor.SetFocus
End Sub

Private Sub Form_Load()
    Set GrdBusProveedor.DataSource = objProveedor.ListaRegMagistral(frm_VTA_RecetarioM.pstrDatoProv, frm_VTA_RecetarioM.pstrFlgRM, objUsuario.CodigoLocal)
    
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    arrCampos = Array("RUC_PROVEEDOR", "NOM_PROVEEDOR", "DIR_PROVEEDOR")
    arrCaption = Array("Ruc", "Proveedor", "Dirección")
    arrAncho = Array(1000, 2200, 2500, 2000, 0)
    arrAlineacion = Array(vbAlignNone, vbAlignLeft, vbAlignLeft)
    GrdBusProveedor.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub

Private Sub cmdAceptar_Click()
    Call GrdBusProveedor_DblClick
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub GrdBusProveedor_DblClick()
    frm_VTA_RecetarioM.TxtProveedor.Text = "" & GrdBusProveedor.Columns("RUC_PROVEEDOR").Value
    frm_VTA_RecetarioM.LblNomprov = "" & GrdBusProveedor.Columns("NOM_PROVEEDOR").Value
    Unload Me
End Sub


Private Sub GrdBusProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            GrdBusProveedor_DblClick
    End Select
End Sub


