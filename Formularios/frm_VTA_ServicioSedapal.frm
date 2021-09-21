VERSION 5.00
Begin VB.Form frm_VTA_ServicioSedapal 
   BorderStyle     =   0  'None
   ClientHeight    =   7245
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlGrillaArray GrdSuministros 
      Height          =   1815
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3201
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.ComboBox cboCriterio 
      Height          =   315
      ItemData        =   "frm_VTA_ServicioSedapal.frx":0000
      Left            =   120
      List            =   "frm_VTA_ServicioSedapal.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4440
      Picture         =   "frm_VTA_ServicioSedapal.frx":0041
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5745
      Picture         =   "frm_VTA_ServicioSedapal.frx":05CB
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   1095
   End
   Begin vbp_Ventas.ctlTextBox txtCriterio 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   1920
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   556
      Tipo            =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbp_Ventas.ctlGrilla grdDocumentos 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1720
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Suministro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   930
   End
   Begin VB.Label lblSuministro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   3480
      Width           =   4695
   End
   Begin VB.Label lblDireccion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   3000
      Width           =   4695
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Dirección :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   945
   End
   Begin VB.Label lblCliente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   2520
      Width           =   4695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Buscar por "
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Shift+Enter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   6
      Left            =   4380
      TabIndex        =   9
      Top             =   6900
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   5
      Left            =   6097
      TabIndex        =   8
      Top             =   6900
      Width           =   390
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   660
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_VTA_ServicioSedapal.frx":0B55
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Servicios - Sedapal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Index           =   4
      Left            =   420
      TabIndex        =   6
      Top             =   60
      Width           =   2055
   End
End
Attribute VB_Name = "frm_VTA_ServicioSedapal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objServicio As New clsServicio
Public strCodigoPadre As String
Public strDescripcionPadre As String
Dim objHComVPos As Object

Private Sub cboCriterio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdCancelar_Click
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo handle
    objVenta.AgregaServicio strCodigoPadre, _
                            strDescripcionPadre, _
                            grdDocumentos.DataSource("COD_SERVICIO").Value, _
                            grdDocumentos.DataSource("DES_SERVICIO").Value, _
                            lblSuministro.Caption, _
                            Servicio, _
                            GrdSuministros.Columns("Monto").Value, _
                            objUsuario.TipoCambio, _
                            Val(GrdSuministros.Columns("Monto").Value), _
                            grdDocumentos.Columns("COD_PRODUCTO").Value, _
                            GrdSuministros.Columns("N° Referencia").Value, _
                            lblSuministro.Caption, "0"
    frm_VTA_Servicios.grdServicios.Rebind
    Unload Me
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdBuscar_Click()
On Error GoTo handle
    Dim objSedapal As New clsServicio
        objSedapal.ConectaSedapal cboCriterio.ListIndex, txtCriterio.Text
      ' GrdSuministros.Array1 = objSedapal.Recibos
        lblCliente.Caption = objSedapal.Cliente
        lblDireccion.Caption = objSedapal.Direccion
        lblSuministro.Caption = objSedapal.Suministro
    Set objSedapal = Nothing
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = 1 Then cmdAceptar_Click
End Sub

Private Sub Form_Load()
On Error GoTo handle
    SetteaFormulario Me
            Dim arrCampos, arrCaption, arrAncho, arrAlineacion
    arrCampos = Array("COD_SERVICIO", "DES_SERVICIO", "COD_PRODUCTO")
    arrCaption = Array("Codigo", "Servicio", "CodProducto")
    arrAncho = Array(1000, 4500, 0)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgGeneral)
    grdDocumentos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdDocumentos.Columns(2).Visible = False
    Set objHComVPos = CreateObject("HComVPos.CC_Transaction")
    Set grdDocumentos.DataSource = objServicio.Lista("", strCodigoPadre)
    'txtCriterio.Text = "28123073"
       'GrdSuministros.TipoArray = True
       'Dim arrCampos, arrCaption, arrAncho, arrAlineacion As Variant
       arrCampos = Array("", "", "", "", "")
       arrCaption = Array("N° Referencia", "Fecha", "Monto", "xxx", "xxx")
       arrAncho = Array("1500", "1000", "800", "800", "800")
       arrAlineacion = Array(vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft)
       GrdSuministros.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
       GrdSuministros.Columns(3).Visible = False
       GrdSuministros.Columns(4).Visible = False

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub


Private Sub grdDocumentos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub
