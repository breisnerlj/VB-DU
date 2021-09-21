VERSION 5.00
Begin VB.Form frm_VTA_Canje 
   BorderStyle     =   0  'None
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_VTA_Canje.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      Picture         =   "frm_VTA_Canje.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5145
      Picture         =   "frm_VTA_Canje.frx":0894
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   1095
   End
   Begin vbp_Ventas.ctlGrilla grdPromocion 
      Height          =   1455
      Left            =   60
      TabIndex        =   0
      Top             =   795
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2566
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlGrilla grdCanje 
      Height          =   3315
      Left            =   60
      TabIndex        =   1
      Top             =   2760
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5847
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
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
      Left            =   3780
      TabIndex        =   10
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
      Left            =   5565
      TabIndex        =   9
      Top             =   6900
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_VTA_Canje.frx":0E1E
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Canje de promoción"
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
      Index           =   3
      Left            =   420
      TabIndex        =   8
      Top             =   60
      Width           =   2130
   End
   Begin VB.Label Label1 
      Caption         =   "Producto a canjear :"
      Height          =   255
      Index           =   2
      Left            =   420
      TabIndex        =   7
      Top             =   2355
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F2"
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
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   2340
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Promoción : "
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   5
      Top             =   495
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F1"
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
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   255
   End
End
Attribute VB_Name = "frm_VTA_Canje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objPromocion As New clsPromocion
Dim objProducto As New clsProducto

Private Sub cmdAceptar_Click()
Dim Indicador As String
Dim PctComi As Double

    If grdCanje.ApproxCount = 0 Then Exit Sub
    PctComi = objProducto.pctComision(IIf(IsNull(grdCanje.Columns(0).Value), "", grdCanje.Columns(0).Value), objUsuario.CodigoLocal, Format(ptmTipoPrecio, "000"))
    Indicador = objProducto.CodIndicadorReceta(grdCanje.Columns(0).Value)
    objVenta.AgregaProducto grdCanje.Columns("COD_PRODUCTO").Value, _
                            grdCanje.Columns("DES_PRODUCTO").Value, _
                            grdCanje.Columns("CTD_CANJE").Value, _
                            grdCanje.Columns("FLG_FRACCION").Value, _
                            0, _
                            objVenta.CodigoTipoVenta, _
                            Producto_Regalo, , , , , , _
                            Indicador, PctComi
    'frmPedido.AgregaProducto grdCanje.Columns(0).Value, grdCanje.Columns(1).Value, grdCanje.Columns(2).Value, 0
    frmPedido.grdPedido.Rebind
   Unload Me
    'Me.Hide
End Sub
Private Sub grdCanje_DblClick()

'''''''''Dim dblPctComision As Double
'''''''''    If grdPromocion.ApproxCount = 0 Then Exit Sub
'''''''''    dblPctComision = objProducto.PctComision(grdCanje.Columns(0).Value, objUsuario.CodigoLocal, Format(ptmTipoPrecio, "000"))
'''''''''    objVenta.AgregaProducto grdCanje.Columns(0).Value, grdCanje.Columns(1).Value, grdCanje.Columns(2).Value, 0, 0, objVenta.CodigoTipoVenta, Producto_Normal, dblPctComision
'''''''''    frmPedido.Cal_Montos
'''''''''    frmPedido.grdPedido.Rebind
End Sub
'''''''''Private Sub grdCanje_KeyDown(KeyCode As Integer, Shift As Integer)
'''''''''    Select Case KeyCode
'''''''''        Case vbKeyReturn
'''''''''            grdCanje_DblClick
'''''''''    End Select
'''''''''End Sub
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

Private Sub grdPromocion_RegistroSeleccionado(ByVal DatoColumna0 As String)
    If grdPromocion.ApproxCount = 0 Then Exit Sub
    Set grdCanje.DataSource = objPromocion.ListaCanjeProducto(grdPromocion.DataSource("COD_PROMOCION"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPromocion = Nothing 'Descarga el Objeto Promocion cuando sale del Formulario
End Sub

Private Sub Form_Load()
    SetteaFormulario Me
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    
    arrCampos = Array("COD_PROMOCION", "DES_PROMOCION")
    arrCaption = Array("Codigo", "Promoción")
    arrAncho = Array(1800, 3800)
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft)
    
    grdPromocion.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    Set grdPromocion.DataSource = objPromocion.ListaLocal("", objUsuario.CodigoEmpresa, objUsuario.CodigoLocal)
    
    arrCampos = Array("COD_PRODUCTO", "CTD_CANJE", "DES_PRODUCTO", "FLG_FRACCION")
    arrCaption = Array("Codigo", "Ctd", "Regalo", "F")
    arrAncho = Array(700, 300, 3800, 200)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgLeft, dbgCenter)
    grdCanje.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdCanje.Columns(0).WrapText = True
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
    'objVenta.CancelarVenta
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control
    psub_KeyDownAplicacion KeyCode, Shift
    Select Case KeyCode
        Case vbKeyF1
            grdPromocion.SetFocus
        Case vbKeyF2
            grdCanje.SetFocus
        Case vbKeyReturn
            If Shift = 1 Then cmdAceptar_Click
        Case vbKeyEscape
            cmdCancelar_Click
    End Select

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

