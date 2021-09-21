VERSION 5.00
Begin VB.Form frm_ADM_ProductosSeleccionados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Productos Seleccionados"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   10635
   ControlBox      =   0   'False
   Icon            =   "frm_adm_productosseleccionados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   405
      Left            =   3720
      TabIndex        =   27
      Top             =   690
      Width           =   6255
      Begin VB.Label lblTotProducto 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2160
         TabIndex        =   29
         Top             =   135
         Width           =   75
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total de Productos: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   135
         Width           =   1755
      End
   End
   Begin VB.Frame fraStock 
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   705
      Left            =   8280
      TabIndex        =   23
      Top             =   1080
      Width           =   1695
      Begin VB.CheckBox chkConStock 
         Caption         =   "Tienen Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   1
         Top             =   240
         Width           =   1545
      End
   End
   Begin VB.Frame FraTerapeutico 
      Caption         =   "Grupo y Accion Terapeutica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   0
      TabIndex        =   24
      Top             =   1800
      Width           =   3615
      Begin vbp_Ventas.ctlDataCombo ctlCboGrupo 
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo ctlCboAccion 
         Height          =   315
         Left            =   720
         TabIndex        =   5
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label lblGrupoTerapeutico 
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   435
      End
      Begin VB.Label lblAccionTerapeutico 
         AutoSize        =   -1  'True
         Caption         =   "Acción"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame FraLbLin 
      Caption         =   "Laboratorio y Linea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   0
      TabIndex        =   20
      Top             =   690
      Width           =   3500
      Begin vbp_Ventas.ctlDataCombo ctlCboLaboratorio 
         Height          =   315
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo ctlCboLinea 
         Height          =   315
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label lblLinea 
         Caption         =   "Li&n"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Lab"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame FraClase 
      Caption         =   "Clase Comercial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   3720
      TabIndex        =   15
      Top             =   1800
      Width           =   6855
      Begin vbp_Ventas.ctlDataCombo ctlCboClase 
         Height          =   315
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo ctlCboSubClase 
         Height          =   315
         Left            =   840
         TabIndex        =   7
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo ctlCboFamilia 
         Height          =   315
         Left            =   4080
         TabIndex        =   8
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo ctlCboCategoria 
         Height          =   315
         Left            =   4080
         TabIndex        =   9
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label lblSubClase 
         AutoSize        =   -1  'True
         Caption         =   "S.Clase"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   540
      End
      Begin VB.Label lblClase 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   390
      End
      Begin VB.Label lblCategoria 
         AutoSize        =   -1  'True
         Caption         =   "Categ:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3360
         TabIndex        =   17
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lblSubCategoria 
         AutoSize        =   -1  'True
         Caption         =   "SCateg:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3360
         TabIndex        =   16
         Top             =   720
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   705
      Left            =   3720
      TabIndex        =   14
      Top             =   1080
      Width           =   2175
      Begin vbp_Ventas.ctlTextBox TxtProducto 
         Height          =   345
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "&Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   705
      Left            =   6120
      TabIndex        =   13
      Top             =   1080
      Width           =   1935
      Begin vbp_Ventas.ctlDataCombo ctlCboEstado 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         MatchEntry      =   1
         Enabled         =   0   'False
      End
   End
   Begin vbp_Ventas.ctlToolBar ToolBar_ADM_Producto 
      Height          =   600
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   1058
      ModoBotones     =   7
   End
   Begin vbp_Ventas.ctlGrilla grdComisionados 
      Height          =   4815
      Left            =   0
      TabIndex        =   11
      Top             =   3000
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8493
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
End
Attribute VB_Name = "frm_ADM_ProductosSeleccionados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objProducto As New clsProducto
Dim oraBusqueda As oraDynaset
Dim objLaboratorio As New clsLaboratorio
Dim objLinea As New clsLinea

Private Sub chkConStock_Click()
    If TxtProducto.Text <> "" Then sub_Buscar
End Sub

Private Sub ctlCboGrupo_Change()
    Carga_Accion ctlCboGrupo.BoundText
End Sub

Private Sub Form_Load()

On Error GoTo Control
    Carga_Laboratorio
    Carga_Estados
    ctlCboEstado.BoundText = "ACT"
    
    ToolBar_ADM_Producto.Buttons(6).Visible = False
    Carga_Clase
    Carga_Grupo
    SetteaGrd
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub ctlCboClase_Change()
    Carga_SubClase ctlCboClase.BoundText
End Sub

Private Sub ctlCboSubClase_Change()
    Carga_Familia ctlCboClase.BoundText, ctlCboSubClase.BoundText
End Sub

Private Sub ctlCboFamilia_Change()
    Carga_Categoria ctlCboClase.BoundText, _
                    ctlCboSubClase.BoundText, _
                    ctlCboFamilia.BoundText
End Sub

Private Sub ctlCboLaboratorio_Change()
    Carga_Linea ctlCboLaboratorio.BoundText
End Sub

Private Sub ToolBar_ADM_Producto_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    Select Case boton
        Case Buscar
          sub_Buscar
        Case tb_Actualizar
          sub_Buscar
        Case Imprimir
            
        Case tb_Excel
            If grdComisionados.ApproxCount > 0 Then grdComisionados.MostrarExcel
            
        Case tb_email
            If grdComisionados.ApproxCount > 0 Then grdComisionados.MostrarEmail
        Case salir
            Unload Me
    End Select
End Sub

Private Sub TxtProducto_KeyPress(KeyAscii As Integer)
    TxtProducto.Tipo = AlfaNumerico
End Sub

Sub Carga_Estados()

    On Error GoTo handle

    Set ctlCboEstado.RowSource = objProducto.ListaEstado
    ctlCboEstado.ListField = "DES"
    ctlCboEstado.BoundColumn = "COD"
    ctlCboEstado.BoundText = "*"
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
    
End Sub

Sub Carga_Laboratorio()

    On Error GoTo handle
     Set ctlCboLaboratorio.RowSource = objLaboratorio.Lista     '''objProducto.ListaLaboratorio
     ctlCboLaboratorio.ListField = "DES"
     ctlCboLaboratorio.BoundColumn = "COD"
     ctlCboLaboratorio.BoundText = "*"
     
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
     
End Sub

Sub Carga_Linea(ByVal vstrCodLab As String)
    On Error GoTo handle
    Set ctlCboLinea.RowSource = objLinea.Lista(vstrCodLab)       '''objProducto.ListaLinea(vstrCodLab)
    ctlCboLinea.ListField = "DES"
    ctlCboLinea.BoundColumn = "COD"
    ctlCboLinea.BoundText = "*"
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Sub Carga_Clase()
    On Error GoTo handle
     Set ctlCboClase.RowSource = objProducto.ListaClase
     ctlCboClase.ListField = "DES"
     ctlCboClase.BoundColumn = "COD"
     ctlCboClase.BoundText = "*"
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Sub Carga_SubClase(ByVal vstrCodClase As String)
    On Error GoTo handle
     Set ctlCboSubClase.RowSource = objProducto.ListaSubclase(vstrCodClase)
     ctlCboSubClase.ListField = "DES"
     ctlCboSubClase.BoundColumn = "COD"
     ctlCboSubClase.BoundText = "*"
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Sub Carga_Familia(ByVal vstrCodClase As String, ByVal vstrCodSClase As String)
    On Error GoTo handle
     Set ctlCboFamilia.RowSource = objProducto.ListaFamilia(vstrCodClase, vstrCodSClase)
     ctlCboFamilia.ListField = "DES"
     ctlCboFamilia.BoundColumn = "COD"
     ctlCboFamilia.BoundText = "*"
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Sub Carga_Categoria(ByVal vstrCodClase As String, _
                    ByVal vstrCodSClase As String, _
                    ByVal vstrCodFam As String)
                    
    On Error GoTo handle
     Set ctlCboCategoria.RowSource = objProducto.ListaCategoria(vstrCodClase, vstrCodSClase, vstrCodFam)
     ctlCboCategoria.ListField = "DES"
     ctlCboCategoria.BoundColumn = "COD"
     ctlCboCategoria.BoundText = "*"
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Sub Carga_Grupo()
    
    On Error GoTo handle
     Set ctlCboGrupo.RowSource = objProducto.ListaGrupo     '''objProducto.ListaLaboratorio
     ctlCboGrupo.ListField = "DES"
     ctlCboGrupo.BoundColumn = "COD"
     ctlCboGrupo.BoundText = "*"
     
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Sub Carga_Accion(ByVal vstrCodGrpTerapeutico As String)
    On Error GoTo handle
    Set ctlCboAccion.RowSource = objProducto.ListaAccion(vstrCodGrpTerapeutico)       '''objProducto.ListaLinea(vstrCodLab)
    ctlCboAccion.ListField = "DES"
    ctlCboAccion.BoundColumn = "COD"
    ctlCboAccion.BoundText = "*"
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

Sub Nuevo()
    ctlCboLaboratorio.BoundText = "*": ctlCboLinea.BoundText = "*"
    ctlCboClase.BoundText = "*": ctlCboSubClase.BoundText = "*"
    ctlCboFamilia.BoundText = "*": ctlCboCategoria.BoundText = "*"
    ctlCboEstado.BoundText = "*"
    TxtProducto.Text = "": TxtProducto.SetFocus
    grdComisionados.Limpiar
End Sub

Sub sub_Buscar()

On Error GoTo handle

'    If ctlCboClase.BoundText = "*" Or ctlCboClase.BoundText = "05" And ctlCboGrupo.BoundText = "*" And ctlCboLaboratorio.BoundText = "*" Then
'        MsgBox "Debe selecciionar otro criterio de busqueda adicional", vbApplicationModal, App.ProductName
'        Exit Sub
'    End If
    
        Set oraBusqueda = objProducto.ConsultaProductoLocalComisionados(ctlCboLaboratorio.BoundText, _
                                                                        ctlCboLinea.BoundText, _
                                                                        ctlCboClase.BoundText, _
                                                                        ctlCboSubClase.BoundText, _
                                                                        ctlCboFamilia.BoundText, _
                                                                        ctlCboCategoria.BoundText, _
                                                                        ctlCboGrupo.BoundText, _
                                                                        ctlCboAccion.BoundText, _
                                                                        TxtProducto.Text, _
                                                                        objUsuario.CodigoLocal, _
                                                                        chkConStock.Value)
        SetteaGrd
        Set grdComisionados.DataSource = oraBusqueda
        lblTotProducto.Caption = oraBusqueda.RecordCount & " - " & "Registros Encontrados"
        Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub SetteaGrd()
Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant

    arrCampos = Array("CODIGO", "DESCRIPCION", "COMISION", "FECHA", "DES_LABORATORIO", "DES_LINEA", _
                      "PRC_UNT", "PRC_FRAC", "DES_CLASE_COM", "DES_GRUPO_TERAP", _
                      "DES_ACCION_TERAP", "FRACCIONA", "CTD_FRACCIONAMIENTO", "STOCK", "VAL_MAX")
    
    arrCaption = Array("Codigo", "Descripción", "%Selecc", "Fch.Comision", "Laboratorio", "Línea", _
                       "Prc.Und.", "Prc.Frac", "Clase", "Grp.Terap.", _
                       "Acc.Terap.", "Fracciona", "Frac.", "Stock", "Max")
    
    arrAncho = Array(800, 2500, 800, 1100, 2500, 2500, _
                     800, 800, 1300, 2500, _
                     2000, 900, 800, 800, 800)
    
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgCenter, dbgLeft, dbgLeft, dbgLeft, _
                          dbgRight, dbgRight, dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, dbgRight, dbgRight, dbgRight)

    grdComisionados.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdComisionados.Columns(1).WrapText = True

    grdComisionados.Columns(3).NumberFormat = "YYYY-MM-DD"
    grdComisionados.Columns(6).NumberFormat = "###,###0.00"
    grdComisionados.Columns(7).NumberFormat = "###,###0.00"


'GrdProductos.RowHeight = GrdProductos.RowHeight * 2.2
 
'grdComisionados.Columns(3).Visible = True
'If chkStock.Value = 0 Then grdProductos.Columns(3).Visible = False

End Sub
