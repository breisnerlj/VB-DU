VERSION 5.00
Begin VB.Form frm_ADM_PVM_Ingreso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PVM"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdF5 
      Caption         =   "[F5] Quitar Producto"
      Height          =   375
      Left            =   2040
      TabIndex        =   36
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "[Esc] Cerrar"
      Height          =   375
      Left            =   8640
      TabIndex        =   35
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Frame fraFiltro 
      Height          =   1575
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   9975
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "Li&mpiar"
         Height          =   615
         Left            =   8640
         Picture         =   "frm_ADM_PVM_Ingreso.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Elimina los No Solicitados"
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdAñadir 
         Caption         =   "&Añadir"
         Height          =   615
         Left            =   8640
         Picture         =   "frm_ADM_PVM_Ingreso.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Añade Productos a Detalle"
         Top             =   200
         Width           =   1095
      End
      Begin vbp_Ventas.ctlDataCombo cboLaboratorio 
         Height          =   315
         Left            =   1080
         TabIndex        =   25
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboLinea 
         Height          =   315
         Left            =   5280
         TabIndex        =   26
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlTextBox TxtProducto 
         Height          =   375
         Left            =   1080
         TabIndex        =   27
         Top             =   600
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   661
         Tipo            =   8
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
      Begin VB.Label lblProducto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1080
         TabIndex        =   33
         Top             =   1080
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label lblCod_Producto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6360
         TabIndex        =   31
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         Caption         =   "&Producto"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Lí&nea"
         Height          =   255
         Left            =   4680
         TabIndex        =   29
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "&Laboratorio"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Su Local ha vendido en los ultimos meses"
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   6120
      Width           =   9975
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "Meses"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   250
         Width           =   2535
      End
      Begin VB.Label lblMes4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "30 d."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   5550
         TabIndex        =   20
         Top             =   250
         Width           =   1000
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Venta Mensual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   510
         Width           =   2415
      End
      Begin VB.Label lblRot30 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5550
         TabIndex        =   18
         Top             =   510
         Width           =   1000
      End
      Begin VB.Label lblMes3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "60 d."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   4550
         TabIndex        =   17
         Top             =   250
         Width           =   1000
      End
      Begin VB.Label lblRot60 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4550
         TabIndex        =   16
         Top             =   510
         Width           =   1000
      End
      Begin VB.Label lblMes2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "90 d."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   3540
         TabIndex        =   15
         Top             =   250
         Width           =   1005
      End
      Begin VB.Label lblRot90 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3540
         TabIndex        =   14
         Top             =   510
         Width           =   1005
      End
      Begin VB.Label lblMes1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "120 d."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   2540
         TabIndex        =   13
         Top             =   250
         Width           =   1000
      End
      Begin VB.Label lblRot120 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2540
         TabIndex        =   12
         Top             =   510
         Width           =   1005
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   9975
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "Laboratorio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "Min. Exhib."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "Transito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   4680
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "Cant. Ult. Ped."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   6240
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblDesLaboratorio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lblFecUltPedido 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblStkLocal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4680
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblStkAlm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6240
         TabIndex        =   3
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "[F3] Ver Historico"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   9975
      Begin vbp_Ventas.ctlGrillaArray grdDetalle 
         Height          =   3375
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   5953
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
End
Attribute VB_Name = "frm_ADM_PVM_Ingreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objPvm As New clsSPVM
Dim solicitud As New clsSPVM

Private Sub cboLaboratorio_Change()
    On Error GoTo CtrlErr
    
    CargaLinea cboLaboratorio.BoundText
    LimpiaProducto
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub CargaLinea(ByVal vstrCodLab As String)
    Dim objLinea As New clsLinea
    
    On Error GoTo CtrlErr
    Set cboLinea.RowSource = objLinea.Lista(vstrCodLab, "1", "[SELECCIONAR]")
    Set objLinea = Nothing
    
    cboLinea.ListField = "DES"
    cboLinea.BoundColumn = "COD"
    cboLinea.BoundText = "*"
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub
    
Private Sub LimpiaProducto()
    lblCod_Producto.Caption = ""
    lblProducto.Caption = ""
    lblEstado.Caption = ""
End Sub

Private Sub cmdAñadir_Click()
    Dim odynTemp As oraDynaset
    
    If lblCod_Producto.Caption = "" And cboLaboratorio.BoundText = "*" Then
        MsgBox "Seleccione criterio de búsqueda, ya sea línea o Producto", vbExclamation, "Aviso"
        Exit Sub
    End If
    If cboLaboratorio.BoundText <> "*" And cboLinea.BoundText = "*" Then
        MsgBox "Seleccione línea", vbExclamation, "Aviso"
        cboLinea.SetFocus
        Exit Sub
    End If
        
    Set odynTemp = solicitud.ListaProductosPVM(objUsuario.CodigoLocal, _
                                               Me.cboLaboratorio.BoundText, _
                                               Me.cboLinea.BoundText, _
                                               Me.lblCod_Producto.Caption)
    If odynTemp.RecordCount = 0 Then
        MsgBox "No se ubicaron productos activos con criterio de búsqueda", vbExclamation, "Aviso"
        If lblCod_Producto.Caption <> "" Then
            TxtProducto.SetFocus
        Else
            cboLinea.SetFocus
        End If
    Else
        AdicionaDetalle odynTemp
        Me.lblCod_Producto.Caption = ""
        Me.lblProducto.Caption = ""
        Me.TxtProducto.Text = ""
    End If

    Set odynTemp = Nothing
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdF5_Click()
    If Me.grdDetalle.ApproxCount < 1 Then Exit Sub
    If MsgBox("¿Seguro(a) de Eliminar el producto " & Me.grdDetalle.Columns(2).Value & " de la lista?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
       Exit Sub
    End If
    Me.grdDetalle.Delete
    Me.grdDetalle.Refresh
End Sub

Private Sub cmdLimpiar_Click()
    On Error GoTo CtrlErr
'    solicitud.Detalle.Clear
'    solicitud.Detalle.ReDim 0, -1, 0, 7
'    Me.grdDetalle.Limpiar
    If grdDetalle.ApproxCount < 1 Then Exit Sub
        
    If MsgBox("¿Seguro(a) de Eliminar Items que no tienen cantidades de solicitud?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
        
    grdDetalle.MoveFirst

    Screen.MousePointer = vbHourglass
    While Not grdDetalle.EOF
        If Val("" & grdDetalle.Columns(6).Value) = 0 Then
            grdDetalle.Delete
        Else
            grdDetalle.MoveNext
        End If
    Wend
    Screen.MousePointer = vbDefault
            
    Me.grdDetalle.MoveFirst
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub Command2_Click()
    If Me.grdDetalle.ApproxCount > 0 Then
        frm_ADM_PVM_Historico.codProducto = Me.grdDetalle.Columns(0).Value
        frm_ADM_PVM_Historico.lblDesLaboratorio = Me.grdDetalle.Columns(2).Value
        frm_ADM_PVM_Historico.lblDesProducto.Caption = Me.grdDetalle.Columns(0).Value & " - " & Me.grdDetalle.Columns(1).Value 'Me.ctlGrilla1.Columns(0).Value & " - " & Me.ctlGrilla1.Columns(2).Value
        frm_ADM_PVM_Historico.Show vbModal
    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Command4_Click
    End If
    If KeyCode = vbKeyF3 Then
        Command2_Click
    End If
    If KeyCode = vbKeyF5 Then
        cmdF5_Click
    End If
End Sub

Private Sub Form_Load()
    solicitud.Detalle.Clear
    solicitud.Detalle.ReDim 0, -1, 0, 7
    Me.grdDetalle.Limpiar
    CargaLaboratorio
    SeteaGrilla2
End Sub

Private Sub CargaLaboratorio()
    Dim objLaboratorio As New clsLaboratorio
    Set cboLaboratorio.RowSource = objLaboratorio.Lista("1", "[SELECCIONAR]")
    Set objLaboratorio = Nothing
    
    cboLaboratorio.ListField = "DES"
    cboLaboratorio.BoundColumn = "COD"
    cboLaboratorio.BoundText = "*"
End Sub

Public Sub AdicionaDetalle(ByRef rodynTemp As oraDynaset)
    Dim intRow As Integer
    Dim strMensaje As String
    Dim blnEncontrado As Boolean
    Dim btnActualizar As Boolean
    Dim inTf As Integer

    rodynTemp.MoveFirst
    strMensaje = ""
    btnActualizar = False
    intRow = 0
    
    Screen.MousePointer = vbHourglass
    While Not rodynTemp.EOF
        'Permite Ubicar en producto en la grilla o de lo contrario encontralo lo pinta'
        blnEncontrado = False
        If solicitud.Detalle.Count(1) > 0 Then '''' .UpperBound(1) > solicitud.Detalle.LowerBound(1) - 1 Then
            inTf = solicitud.Detalle.Find(solicitud.Detalle.LowerBound(1), 0, CStr("" & rodynTemp("COD_PRODUCTO").Value))
            If Not inTf = solicitud.Detalle.LowerBound(1) - 1 Then
                blnEncontrado = True
            End If
        End If
               
        If Not blnEncontrado Then
            btnActualizar = True
            solicitud.Detalle.InsertRows intRow
            solicitud.Detalle(intRow, 0) = "" & rodynTemp("COD_PRODUCTO").Value
            solicitud.Detalle(intRow, 1) = "" & rodynTemp("COD_ESTADO").Value
            solicitud.Detalle(intRow, 2) = "" & rodynTemp("DES_PRODUCTO").Value
            solicitud.Detalle(intRow, 3) = "" & rodynTemp("DES_LABORATORIO").Value
            solicitud.Detalle(intRow, 4) = "" & rodynTemp("STK").Value
            solicitud.Detalle(intRow, 5) = "" & rodynTemp("PVM").Value
            solicitud.Detalle(intRow, 6) = "" & rodynTemp("SOLICITADO").Value
            solicitud.Detalle(intRow, 7) = "" & rodynTemp("APROBADO").Value
            intRow = intRow + 1
        Else
            strMensaje = strMensaje & CStr("" & rodynTemp("COD_PRODUCTO").Value) & " - " & CStr("" & rodynTemp("DES_PRODUCTO").Value) & Chr(13)
        End If
        
        rodynTemp.MoveNext
    Wend
    Screen.MousePointer = vbDefault
    
    If strMensaje <> "" Then
        MsgBox "Los siguientes productos ya se encontraban en la lista: " & Chr(13) & strMensaje, vbInformation, "Aviso"
    End If
    grdDetalle.Rebind
    grdDetalle.Refresh

End Sub

Sub modificarDetalle(ByVal codigo As String)
    Dim max, intRow As Integer
    Dim rodynTemp As oraDynaset
    
    intRow = solicitud.Detalle.Find(0, 0, codigo, XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
            
    Set rodynTemp = solicitud.ListaProductosPVM(objUsuario.CodigoLocal, "", "", codigo)
    solicitud.Detalle(intRow, 0) = "" & rodynTemp("COD_PRODUCTO").Value
    solicitud.Detalle(intRow, 1) = "" & rodynTemp("COD_ESTADO").Value
    solicitud.Detalle(intRow, 2) = "" & rodynTemp("DES_PRODUCTO").Value
    solicitud.Detalle(intRow, 3) = "" & rodynTemp("DES_LABORATORIO").Value
    solicitud.Detalle(intRow, 4) = "" & rodynTemp("STK").Value
    solicitud.Detalle(intRow, 5) = "" & rodynTemp("PVM").Value
    solicitud.Detalle(intRow, 6) = "" & rodynTemp("SOLICITADO").Value
    solicitud.Detalle(intRow, 7) = "" & rodynTemp("APROBADO").Value
    Me.grdDetalle.Rebind
End Sub

Sub CargaSubDetalle()
    Dim objDatos As New clsEstadistica
    Dim odynDatos As oraDynaset
        
    On Error GoTo Handle
    If Me.grdDetalle.ApproxCount <= 0 Or IsNull(Me.grdDetalle.Columns(0).Value) Then
        Me.lblRot30.Caption = "0.00"
        Me.lblRot60.Caption = "0.00"
        Me.lblRot90.Caption = "0.00"
        Me.lblRot120.Caption = "0.00"
        Me.lblStkLocal.Caption = "0.00"
        Me.lblStkAlm.Caption = "0.00"
        Me.lblDesLaboratorio.Caption = ""
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    Set odynDatos = objDatos.Lista_Cantidades(objUsuario.CodigoLocal, Me.grdDetalle.Columns(0).Value)
    odynDatos.MoveFirst
    While Not odynDatos.EOF
        Me.lblRot30.Caption = Format(odynDatos("MES_0").Value, "0.00")
        Me.lblRot60.Caption = Format(odynDatos("MES_1").Value, "0.00")
        Me.lblRot90.Caption = Format(odynDatos("MES_2").Value, "0.00")
        Me.lblRot120.Caption = Format(odynDatos("MES_3").Value, "0.00")
        Me.lblStkLocal.Caption = Format(odynDatos("STK_TRANSITO").Value, "0.00")
        Me.lblStkAlm.Caption = Format(odynDatos("CTD_ULT_PED").Value, "0.00")
        odynDatos.MoveNext
    Wend
    Me.lblDesLaboratorio.Caption = Me.grdDetalle.Columns(3).Value
    Set odynDatos = Nothing
    Set objDatos = Nothing
    Me.MousePointer = vbDefault
    
    Exit Sub
Handle:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdDetalle_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo Handle
    If ColIndex = 6 Then
        If Not IsNumeric(Trim(grdDetalle.Columns(ColIndex).Value)) And _
                Trim(grdDetalle.Columns(ColIndex).Value) <> "" Then
            MsgBox "El valor no es valido", vbCritical, "Error"
            Cancel = True
            Exit Sub
        End If
    
        If MsgBox("¿Seguro que desea Registrar la cantidad?", vbYesNo) = vbYes Then
            objPvm.AsignarCantidadPVM objUsuario.CodigoLocal, _
                                      Me.grdDetalle.Columns(0).Value, _
                                      objUsuario.codigo, _
                                      Val("" & Me.grdDetalle.Columns(6).Value)
        Else
            Cancel = True
            Exit Sub
        End If
    End If
    Exit Sub
Handle:
    Cancel = True
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdDetalle_AfterColUpdate(ByVal ColIndex As Integer)
    '''If MsgBox("¿Seguro que desea Registrar la cantidad?", vbYesNo) = vbYes Then
        modificarDetalle Me.grdDetalle.Columns(0).Value
    '''Else
    '''    Exit Sub
    '''End If
End Sub

Private Sub grdDetalle_RegistroSeleccionado(ByVal DatoColumna0 As String)
    CargaSubDetalle
End Sub

Private Sub TxtProducto_KeyPress(KeyAscii As Integer)
    On Error GoTo Handle
    
    If KeyAscii = vbKeyReturn Then
        
        If Len(Trim(TxtProducto.Text)) < 3 Then
            MsgBox "use por lo menos 3 caracteres", vbExclamation, "Aviso"
            Exit Sub
        End If
        
        Dim frm As New frm_ADM_ProductoDatos
        frm.Dato = Trim(TxtProducto.Text)
        frm.Show vbModal
        
        If frm.Salida(1) <> "" Then
            cboLaboratorio.BoundText = "*"
        End If
                
        lblCod_Producto.Caption = frm.Salida(1)
        lblProducto.Caption = frm.Salida(2)
        lblEstado.Caption = frm.Salida(3)
        
        Set frm = Nothing
        
        cmdAñadir_Click
        
    End If

    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Public Sub SeteaGrilla2()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim arrFoco As Variant
  Dim i As Integer
      
    '---------------------------------------------------------------
    '-- Detalle
    '---------------------------------------------------------------

            
    arrCampos = Array("cod_producto", "cod_estado", "des_producto", "des_laboratorio", _
                       "stk", "pvm", "solicitado", "aprobado")
    
    arrCaption = Array("Codigo", "Estado", "Descripcion", "Laboratorio", _
                       "Stock", "PVM Act.", "PVM Sol.", "PVM Apr.")
    
    arrAncho = Array(650, 1500, 4500, 500, _
                     900, 900, 900, 800)
    
    arrFoco = Array(False, False, False, False, _
                    False, False, True, False)
                    
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, dbgLeft, _
                          dbgCenter, dbgCenter, dbgCenter, dbgRight)
                              
    With grdDetalle
        .FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion, arrFoco
    
        For i = 0 To .Columns.Count - 1
            .Columns(i).AllowSizing = False
            .Columns(i).WrapText = False
        Next i
       
        .AllowUpdate = True
    
        .EditorStyle.BackColor = vbWhite
        .EditorStyle.ForeColor = RGB(180, 0, 180)
        .EditorStyle.Font.Bold = True
    
        .Columns(6).BackColor = vbInfoBackground
        .Columns(6).DataWidth = 4
        .Columns(2).FetchStyle = True
        .Columns(3).Visible = False
        .Columns(7).Visible = False
    
        .Array1 = solicitud.Detalle
    End With
    
    lblMes1.Caption = Format$(DateAdd("m", -3, Now), "mmm-yy")
    lblMes2.Caption = Format$(DateAdd("m", -2, Now), "mmm-yy")
    lblMes3.Caption = Format$(DateAdd("m", -1, Now), "mmm-yy")
    lblMes4.Caption = Format$(Now, "mmm-yy")
    
End Sub
