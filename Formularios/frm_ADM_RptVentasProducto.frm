VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_ADM_RptVentasProducto 
   BorderStyle     =   0  'None
   Caption         =   "Ventas por Producto"
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ChkFiltroModalidad 
      Caption         =   "Check1"
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   1600
      Width           =   255
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16908289
      CurrentDate     =   40277
   End
   Begin vbp_Ventas.ctlDataCombo CboModalidad 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   1580
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      MatchEntry      =   1
      Enabled         =   0   'False
   End
   Begin vbp_Ventas.ctlToolBar TlbMenu 
      Height          =   600
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   1058
   End
   Begin vbp_Ventas.ctlGrilla grdReporte 
      Height          =   4455
      Left            =   0
      TabIndex        =   6
      Top             =   2400
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7858
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpFin 
      Height          =   315
      Left            =   3600
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16908289
      CurrentDate     =   40277
   End
   Begin vbp_Ventas.ctlTextBox TxtProducto 
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   609
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Buscar [F5]"
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
      Index           =   0
      Left            =   1920
      TabIndex        =   20
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   2
      Left            =   2640
      TabIndex        =   19
      Top             =   1965
      Width           =   195
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   1
      Left            =   5160
      TabIndex        =   18
      Top             =   1965
      Width           =   195
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   0
      Left            =   4920
      TabIndex        =   17
      Top             =   1635
      Width           =   195
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   5
      Left            =   3480
      TabIndex        =   16
      Top             =   920
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Salir [Esc]"
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
      TabIndex        =   15
      Top             =   6960
      Width           =   1080
   End
   Begin VB.Label Label8 
      Caption         =   "Producto"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1280
      Width           =   855
   End
   Begin VB.Label LblCodProducto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1320
      TabIndex        =   5
      Top             =   1210
      Width           =   1215
   End
   Begin VB.Label LblDesProducto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2565
      TabIndex        =   13
      Top             =   1215
      Width           =   4575
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "Total : 0 Registro(s)"
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Hasta"
      Height          =   255
      Left            =   3090
      TabIndex        =   11
      Top             =   1980
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Rango"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1980
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   930
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Modalidad"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1630
      Width           =   855
   End
End
Attribute VB_Name = "frm_ADM_RptVentasProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Autor:Miguel Laguna
'Fecha:09/04/2010
'Proposito:Reporte de ventas por producto

Dim objVenta As New clsVenta
Dim objModalidad As New clsModalidad
Dim vProducto As String
Dim vMod As String
Dim vFecIni As String
Dim vFecFin As String


Private Sub ChkFiltroModalidad_Click()

If ChkFiltroModalidad.Value = 1 Then
CboModalidad.Enabled = True
Else
CboModalidad.Enabled = False
End If

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo handle
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
        Case vbKeyF1
            TxtProducto.SetFocus
        Case vbKeyF2
            If ChkFiltroModalidad.Value = 0 Then
            ChkFiltroModalidad.Value = 1
            ElseIf ChkFiltroModalidad.Value = 1 Then
            ChkFiltroModalidad.Value = 0
            End If
        Case vbKeyF3
            dtpInicio.SetFocus
        Case vbKeyF4
            dtpFin.SetFocus
        Case vbKeyF5
            BuscaVentaProducto
    End Select
handle:
If Err.Description <> "" Then MsgBox Err.Description, vbCritical, App.ProductName: Exit Sub

End Sub

Private Sub Form_Load()

setteaFormulario Me

Me.left = 0
Me.top = 0

LblTotal.Caption = "Total : 0 Registros"

    TlbMenu.VisibleBoton Nuevo, False
    TlbMenu.VisibleBoton Modificar, False
    TlbMenu.VisibleBoton Grabar, False
    TlbMenu.VisibleBoton Cancelar, False
    TlbMenu.VisibleBoton Eliminar, False

'Carga Modalidad

Set CboModalidad.RowSource = objModalidad.Lista
CboModalidad.ListField = "DES_MODALIDAD_VENTA"
CboModalidad.BoundColumn = "COD_MODALIDAD_VENTA"
CboModalidad.BoundText = ""

End Sub

Private Sub grdReporte_DblClick()
LLamaDetalle
End Sub

Private Sub grdReporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
LLamaDetalle
End If
End Sub

Private Sub TlbMenu_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)

On Error GoTo handle

TlbMenu.SetFocus

Select Case boton
    Case Buscar
        BuscaVentaProducto
    Case tb_Actualizar
        BuscaVentaProducto
    Case tlbTipoBoton.Imprimir
        grdReporte.MostrarImprimir
    Case tb_email
        grdReporte.MostrarEmail
    Case tb_Excel
        grdReporte.MostrarExcel
    Case salir
        Unload Me
End Select

handle:
    If Err.Description <> "" Then MsgBox Err.Description, vbCritical, App.ProductName: Exit Sub

End Sub

Public Sub BuscaVentaProducto()
       
    If LblCodProducto = "" Then MsgBox "Debe seleccionar un producto.", vbExclamation + vbOKOnly, "Validación": Exit Sub
    If ChkFiltroModalidad.Value = 1 And CboModalidad.Text = "" Then MsgBox "Debe seleccionar una modalidad de venta.", vbExclamation + vbOKOnly, "Validación": Exit Sub
    If dtpFin < dtpInicio Then MsgBox "La fecha final no puede ser menor a la fecha de inicio.", vbExclamation + vbOKOnly, "Validación": Exit Sub
    If DateDiff("d", dtpInicio.Value, dtpFin.Value) + 1 > 10 Then MsgBox "El rango para el reporte no puede ser mayor a 10 días.", vbExclamation + vbOKOnly, "Validación": Exit Sub
       
    If ChkFiltroModalidad.Value = 0 Then
    vMod = "T"
    Else
    vMod = CboModalidad.BoundText
    End If
    
    vProducto = LblCodProducto
    'vMod = cbomodalida
    vFecIni = Format(dtpInicio.Value, "dd/mm/yyyy")
    vFecFin = Format(dtpFin.Value, "dd/mm/yyyy")
       
    Dim rs As oraDynaset
    Set rs = objVenta.ReporteVentasProducto(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, vProducto, vMod, vFecIni, vFecFin)
    Set grdReporte.DataSource = rs
    
    SetGrdGrilla
    
    If rs.RecordCount = 1 Then
    LblTotal.Caption = "Total : " & rs.RecordCount & " Registro"
    Else
    LblTotal.Caption = "Total : " & rs.RecordCount & " Registros"
    End If
    
    If rs.RecordCount <> 0 Then
    grdReporte.SetFocus
    End If
    
End Sub

Private Sub TxtProducto_KeyPress(KeyAscii As Integer)

    TxtProducto.Tipo = AlfaNumerico
    If KeyAscii = 13 Then
        If Len(Trim(TxtProducto.Text)) < 3 Then MsgBox "Ingresar como minimo 3 digitos", vbInformation + vbOKOnly, App.ProductName: TxtProducto.SetFocus: Exit Sub
               Call frm_ADM_BusGenProducto.LoadForm(Trim(TxtProducto.Text), _
                                                        objUsuario.CodigoLocal)
                                                        
        frm_ADM_BusGenProducto.Show 1
        
        If frm_ADM_BusGenProducto.strCodProdGen <> "" Then
          LblCodProducto.Caption = frm_ADM_BusGenProducto.strCodProdGen
          LblDesProducto.Caption = frm_ADM_BusGenProducto.strDesProdGen
          ChkFiltroModalidad.SetFocus
        End If
    End If
    
End Sub

Private Sub SetGrdGrilla()

  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant


    '"Fec Emision",
    '"FCH_EMISION",
    
    arrCampos = Array("COD_PRODUCTO", "PRODUCTO", "VENDEDOR", "NOMBRE", "NUM_DOC", "VENTA", "PRODUCTOS", "FRACCIONES")
    arrCaption = Array("Codigo", "Descripcion", "Vendedor", "Nombre", "#Doc", "Venta", "Unidades", "Fracciones")
    arrAncho = Array(0, 0, 800, 2600, 600, 800, 1000, 1000)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgCenter, dbgRight, dbgRight, dbgRight)

    grdReporte.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    grdReporte.Columns(0).Visible = False
    grdReporte.Columns(1).Visible = False
    
    grdReporte.Columns(5).NumberFormat = "###0.00"
    grdReporte.Columns(6).NumberFormat = "###0"
    grdReporte.Columns(7).NumberFormat = "###0"
     
End Sub




Public Sub LLamaDetalle()
If grdReporte.ApproxCount <> 0 And vProducto <> "" And vMod <> "" And vFecIni <> "" And vFecFin <> "" Then
 
 frm_ADM_RptVentasProductoDet.LblCodProducto = LblCodProducto
 frm_ADM_RptVentasProductoDet.LblDesProducto = LblDesProducto
 frm_ADM_RptVentasProductoDet.LblCodVendedor = grdReporte.Columns("VENDEDOR").Value
 frm_ADM_RptVentasProductoDet.LblDesVendedor = grdReporte.Columns("NOMBRE").Value
 frm_ADM_RptVentasProductoDet.LblTotalProducto = grdReporte.Columns("PRODUCTOS").Value
 frm_ADM_RptVentasProductoDet.LblTotalFracciones = grdReporte.Columns("FRACCIONES").Value
 
 Call frm_ADM_RptVentasProductoDet.LoadForm(vProducto, vMod, vFecIni, vFecFin, grdReporte.Columns("VENDEDOR").Value)
 
End If
End Sub






