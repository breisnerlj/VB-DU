VERSION 5.00
Begin VB.Form frm_DLV_PromocionesXproductos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Promociones"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin vbp_Ventas.ctlGrilla GrillaProductoPromocion 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5318
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Salir [ESC]"
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
      Index           =   2
      Left            =   7680
      TabIndex        =   1
      Top             =   3105
      Width           =   1155
   End
End
Attribute VB_Name = "frm_DLV_PromocionesXproductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objProducto As New clsProducto
Dim oraBusqueda As oraDynaset
Private strCodigoProducto As String
Private PromocionEncontradoXproducto As String
Dim cantidadPromociones As String

Private Sub Form_Load()
    On Error GoTo Control
    
    SetteaGrd
    sub_Buscar
   
    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.number
End Sub

Sub sub_Buscar()
    Dim strFaccionamiento As String
        
    On Error GoTo handle
        
    Set oraBusqueda = objProducto.ConsultaPromocionProducto(strCodigoProducto, objUsuario.CodLocalCallCenter, mdiPrincipal.ctlCliente1.LocalAsignado) 'GNIBIN 20210127 Proyecto MultiMarca | ECASTILLO 28.01.2021 se descomenta
'    'INI  GNIBIN 20210127 Proyecto MultiMarca
'    Dim strCia As String
'    strCia = mdiPrincipal.ctlCliente1.sCia
'
'    Select Case strCia
'        Case "94", "93", "92", "1DLV"
'            strCia = "1DLV"
'        Case Else
'            strCia = "0DLV"
'    End Select
'
'    Set oraBusqueda = objProducto.ConsultaPromocionProducto(strCodigoProducto, strCia, mdiPrincipal.ctlCliente1.LocalAsignado)
'    'FIN  GNIBIN 20210127 Proyecto MultiMarca
    
    PromocionEncontradoXproducto = oraBusqueda.RecordCount
    
    Set GrillaProductoPromocion.DataSource = oraBusqueda
        
    SetteaGrd
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub SetteaGrd()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    
    GrillaProductoPromocion.RowHeight = 0
    GrillaProductoPromocion.RowHeight = GrillaProductoPromocion.RowHeight * 1.4
    GrillaProductoPromocion.AlternatingRowStyle = True
    GrillaProductoPromocion.Styles(6).BackColor = &HF1F1F1
    
    arrCampos = Array("DES_TIPO_PROMOCION", "DES_PROMOCION", "DES_MENSAJE")
    arrCaption = Array("Tipo de Promoción", "Descripción", "Mensaje")
    arrAncho = Array(2000, 6500, 0)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft)
    GrillaProductoPromocion.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    GrillaProductoPromocion.Columns(0).AutoSize
    GrillaProductoPromocion.Columns(0).AllowSizing = False
    GrillaProductoPromocion.Columns(1).AllowSizing = False
    GrillaProductoPromocion.Columns(2).Visible = False
    
End Sub

Public Property Let codigoproducto(ByVal codProducto As String)
    strCodigoProducto = codProducto
End Property

Public Property Get GetValidaCantidadPromocion() As String
   Set oraBusqueda = objProducto.ConsultaPromocionProducto(strCodigoProducto, objUsuario.CodLocalCallCenter, mdiPrincipal.ctlCliente1.LocalAsignado)
    
    PromocionEncontradoXproducto = oraBusqueda.RecordCount
    GetValidaCantidadPromocion = PromocionEncontradoXproducto
End Property


Private Sub GrillaProductoPromocion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then

Unload Me
 End If
 
 
End Sub

