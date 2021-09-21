VERSION 5.00
Begin VB.Form frm_ADM_RptVentasProductoDet 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7290
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlGrilla grdReporteDetalle 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9128
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlToolBar TlbMenu 
      Height          =   600
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   1058
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "<- Regresar [Esc]"
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
      Index           =   11
      Left            =   120
      TabIndex        =   14
      Top             =   6840
      Width           =   1860
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Venta"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label LblTotalVenta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label LblTotalFracciones 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label LblTotalProducto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Fracciones"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Produto"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label LblCodVendedor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1040
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label LblCodProducto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1040
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Label LblDesVendedor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1920
      TabIndex        =   4
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Vendedor"
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label LblDesProducto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Producto"
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frm_ADM_RptVentasProductoDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim objVenta As New clsVenta

Public Sub LoadForm(ByVal strCodProducto As String, ByVal strModalidad As String, ByVal strFicIni As String, ByVal strFicFin As String, strVendedor)

Dim vTotalVenta As Double
Dim vTotalP As Double
Dim vTotalF As Double

    SeteaGrilla
    Dim rsDatos As oraDynaset
    
    Set rsDatos = objVenta.ReporteVentasProductoDet(objUsuario.CodigoEmpresa, _
                                                            objUsuario.CodigoLocal, _
                                                            strCodProducto, _
                                                            strModalidad, _
                                                            strFicIni, _
                                                            strFicFin, _
                                                            strVendedor)
                                                            
    Set grdReporteDetalle.DataSource = rsDatos
                                                            
                                                            
                                                            
    vTotalVenta = 0
    vTotalP = 0
    vTotalF = 0
    
    
    rsDatos.MoveFirst
    Do While Not rsDatos.EOF
    vTotalVenta = vTotalVenta + grdReporteDetalle.Columns("MTO_TOTAL").Value
    vTotalP = vTotalP + grdReporteDetalle.Columns("CANT_PRODUCTOS").Value
    vTotalF = vTotalF + grdReporteDetalle.Columns("CANT_FRACCIONES").Value
    rsDatos.MoveNext
    Loop
    
    LblTotalVenta = vTotalVenta
    LblTotalProducto = vTotalP
    LblTotalFracciones = vTotalF
    
    Set rsDatos = Nothing
                                                            
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo handle
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
handle:
If Err.Description <> "" Then MsgBox Err.Description, vbCritical, App.ProductName: Exit Sub

End Sub

Private Sub Form_Load()
setteaFormulario Me

Me.left = 0
Me.top = 0

TlbMenu.VisibleBoton Buscar, False
TlbMenu.VisibleBoton tb_Actualizar, False
TlbMenu.VisibleBoton Nuevo, False
TlbMenu.VisibleBoton Modificar, False
TlbMenu.VisibleBoton Grabar, False
TlbMenu.VisibleBoton Cancelar, False
TlbMenu.VisibleBoton Eliminar, False

End Sub



Private Sub grdReporteDetalle_DblClick()

On Error GoTo Control
MuestraDocumento
Exit Sub
Control:

      MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub



Private Sub grdReporteDetalle_KeyPress(KeyAscii As Integer)

On Error GoTo Control
MuestraDocumento
Exit Sub
Control:

      MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
                    
End Sub

Private Sub TlbMenu_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
Select Case boton
    Case tlbTipoBoton.Imprimir
        grdReporteDetalle.MostrarImprimir
    Case tb_email
        grdReporteDetalle.MostrarEmail
    Case tb_Excel
        grdReporteDetalle.MostrarExcel
    Case salir
        Unload Me
End Select
End Sub

Private Sub SeteaGrilla()

  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant

    arrCampos = Array("COD_MODALIDAD_VENTA", "DES_MODALIDAD_VENTA", "COD_PRODUCTO", "DES_PRODUCTO", "FCH_EMISION", "COD_USUARIO_DEPENDIENTE", "NOMBRE", "COD_TIPODOC", "DES_TIPODOC", "NUM_DOCUMENTO", "MTO_TOTAL", "CANT_PRODUCTOS", "CANT_FRACCIONES")
    arrCaption = Array("Codigo", "Modalidad", "Codigo", "Descripcion", "Fecha", "Usuario", "Nombre", "Tipo", "Tipo Doc", "Documento", "Monto", "Ctd Und", "Ctd Frac")
    arrAncho = Array(0, 1300, 0, 0, 1000, 0, 0, 600, 0, 1000, 1000, 900, 900)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgRight, dbgRight, dbgRight)
       
    grdReporteDetalle.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
       
    grdReporteDetalle.Columns(0).Visible = False
    grdReporteDetalle.Columns(2).Visible = False
    grdReporteDetalle.Columns(3).Visible = False
    grdReporteDetalle.Columns(5).Visible = False
    grdReporteDetalle.Columns(6).Visible = False
    grdReporteDetalle.Columns(8).Visible = False
    
    grdReporteDetalle.Columns(10).NumberFormat = "###0.00"
    grdReporteDetalle.Columns(11).NumberFormat = "###0"
    grdReporteDetalle.Columns(12).NumberFormat = "###0"
     
End Sub


Public Sub MuestraDocumento()
If grdReporteDetalle.ApproxCount = 0 Then Exit Sub

frm_ADM_PreviewDoc.Datos objUsuario.CodigoEmpresa, _
objUsuario.CodigoLocal, _
grdReporteDetalle.Columns("COD_TIPODOC").Value, _
grdReporteDetalle.Columns("NUM_DOCUMENTO").Value, _
"", _
"" & grdReporteDetalle.DataSource("COD_MODALIDAD_VENTA")
End Sub
