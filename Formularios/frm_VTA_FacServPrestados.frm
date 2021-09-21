VERSION 5.00
Begin VB.Form frm_VTA_FacServPrestados 
   BorderStyle     =   0  'None
   Caption         =   "Facturación por Servicios Prestados"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   Icon            =   "frm_VTA_FacServPrestados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   6015
      Begin vbp_Ventas.ctlDataCombo ctlCboSConcepto 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo ctlCboConcepto 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
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
         Left            =   240
         TabIndex        =   17
         Top             =   435
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sub Concepto"
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
         Left            =   240
         TabIndex        =   16
         Top             =   810
         Width           =   1275
      End
   End
   Begin VB.Frame fra_Keito 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2415
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   6015
      Begin vbp_Ventas.ctlTextBox TxtPrecio 
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Tipo            =   4
         Alignment       =   1
         MaxLength       =   8
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
      Begin vbp_Ventas.ctlTextBox TxtCantidad 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Tipo            =   4
         Alignment       =   1
         MaxLength       =   3
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
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
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
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin VB.Label LblCodigo 
         BackColor       =   &H00E1FBFA&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblProducto 
         BackColor       =   &H00E1FBFA&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Precio:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1965
         Width           =   630
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   1485
         Width           =   855
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   4440
      Picture         =   "frm_VTA_FacServPrestados.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5745
      Picture         =   "frm_VTA_FacServPrestados.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   1095
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
      Left            =   6097
      TabIndex        =   9
      Top             =   6900
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "frm_VTA_FacServPrestados.frx":0B20
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Servicios Prestados"
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
      Left            =   480
      TabIndex        =   8
      Top             =   120
      Width           =   2145
   End
End
Attribute VB_Name = "frm_VTA_FacServPrestados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ObjServPrestado As New clsServPrestados
Dim ObjFormaPago As New clsFormaPago
Dim odynSer As oraDynaset
Dim objProducto As New clsProducto
Public pstrIDKeito As String
Public pstrCodConcep As String
Public pstrCodSConcep As String

Private Sub Form_Activate()
    ctlCboConcepto.SetFocus
End Sub

Private Sub Form_Load()
     SetteaFormulario Me
     
     '/***********************************************************************/'
     '-- CARGA EL COMBO DE CONCEPTO --'
     Set ctlCboConcepto.RowSource = ObjServPrestado.ListaConcepto(objUsuario.CodigoLocal)
     ctlCboConcepto.ListField = "DESCRIPCION"
     ctlCboConcepto.BoundColumn = "CODIGO"
     
     '-- CARGA FORMA PAGO --'
'     Set ctlCboFormaPago.RowSource = ObjFormaPago.Dev_Forma_Pago
'     ctlCboFormaPago.ListField = "DES"
'     ctlCboFormaPago.BoundColumn = "COD"
End Sub

Private Sub ctlCboConcepto_Change()
     '/***********************************************************************/'
     '-- CARGA EL COMBO DE SUBCONCEPTO --'
     'Debug.Print ctlCboConcepto.BoundText
     Set ctlCboSConcepto.RowSource = ObjServPrestado.ListaSubConcepto(objUsuario.CodigoLocal, ctlCboConcepto.BoundText)
     ctlCboSConcepto.ListField = "DESCRIPCION"
     ctlCboSConcepto.BoundColumn = "CODIGO"
     'ctlCboSConcepto.BoundText = "*"
End Sub

Private Sub ctlCboSConcepto_Change()
    pstrIDKeito = "KE"
    fra_Keito.Visible = True
    
    Set odynSer = ObjServPrestado.ListaProducto(ctlCboSConcepto.BoundText)
    Frame1.Caption = ctlCboSConcepto.BoundText
    pstrCodConcep = ctlCboConcepto.BoundText
    pstrCodSConcep = ctlCboSConcepto.BoundText
    
    LblCodigo.Caption = "" & odynSer("COD_PRODUCTO").Value
    lblProducto.Caption = "" & odynSer("DES_PRODUCTO").Value
    TxtCantidad.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error GoTo Control
        psub_KeyDownAplicacion KeyCode, Shift
        Select Case KeyCode
                            
            Case vbKeyEscape
                Unload Me
        End Select
    
    Exit Sub
Control:
        MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub cmdAceptar_Click()
Dim Indicador As String
Dim PctComi  As Double


    On Error GoTo CtrlErr
    '//////agregado por PHERRERA 24/08/07
    If ctlCboConcepto.BoundText = "" Then MsgBox "Debe seleccionar el concepto.", vbOKOnly + vbExclamation, "Error": ctlCboConcepto.SetFocus: Exit Sub
    If ctlCboSConcepto.BoundText = "" Then MsgBox "Debe seleccionar el sub-concepto.", vbOKOnly + vbExclamation, "Error": ctlCboSConcepto.SetFocus: Exit Sub
    '//////////////////////////////////
    If TxtCantidad.Visible = True Then If Val(TxtCantidad.Text) = 0 Then MsgBox "Ingrese la Cantidad", vbCritical, Caption: TxtCantidad.SetFocus: Exit Sub
    'End If
      
    If TxtPrecio.Visible = True Then If Val(TxtPrecio.Text) = 0 Then MsgBox "Ingrese el Precio", vbCritical, Caption: TxtPrecio.SetFocus: Exit Sub
    'End If
    
    frmPedido.grdPedido.Limpiar

    If Trim(LblCodigo.Caption) = "" Then Exit Sub
    Indicador = objProducto.CodIndicadorReceta(LblCodigo.Caption)
    PctComi = objProducto.pctComision(LblCodigo.Caption, objUsuario.CodigoLocal, Format(ptmTipoPrecio, "000"))

    frmPedido.grdPedido.Array1 = objVenta.AgregaProducto(LblCodigo.Caption, _
                            lblProducto.Caption, _
                            TxtCantidad.Text, _
                            "0", _
                            TxtPrecio.Text, _
                            FactServicios, _
                            Producto_Normal, , , , , , Indicador, PctComi)
                            
                            
    frmPedido.Cal_Montos
    frmPedido.grdPedido.Rebind
    
    
    Unload Me
    
    Exit Sub
    
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
    
    
End Sub



Private Sub cmdCancelar_Click()
    Unload Me
    'objVenta.CancelarVenta
End Sub


