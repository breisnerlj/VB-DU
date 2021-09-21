VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frm_ADM_SMMNuevo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Modificación de Máximos - Nuevo"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAñadir 
      Caption         =   "&Añadir"
      Height          =   615
      Left            =   9240
      Picture         =   "frm_ADM_SMMNuevo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Añade Productos a Detalle"
      Top             =   840
      Width           =   1095
   End
   Begin ORADCLibCtl.ORADC oradcMotivo 
      Height          =   255
      Left            =   1440
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   450
      _StockProps     =   207
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   ""
      Connect         =   ""
      RecordSource    =   ""
   End
   Begin TrueDBGrid70.TDBDropDown drdMotivo 
      Bindings        =   "frm_ADM_SMMNuevo.frx":058A
      Height          =   2535
      Left            =   5160
      TabIndex        =   14
      Top             =   3600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4471
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Cod."
      Columns(0).DataField=   "COD_MOTIVO_SMM"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripción"
      Columns(1).DataField=   "DES_MOTIVO_SMM"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=820"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=741"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=20"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).WrapText=1"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=20"
      Splits(0)._ColumnProps(14)=   "Column(1).WrapText=1"
      Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   1
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   0
      RowDividerStyle =   2
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   0   'False
      ListField       =   ""
      DataField       =   ""
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   13160660
      ValueTranslate  =   0   'False
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.valignment=2,.bgcolor=&H80000018&"
      _StyleDefs(37)  =   ":id=28,.wraptext=-1"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.valignment=2,.wraptext=-1"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(45)  =   "Named:id=33:Normal"
      _StyleDefs(46)  =   ":id=33,.parent=0"
      _StyleDefs(47)  =   "Named:id=34:Heading"
      _StyleDefs(48)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(49)  =   ":id=34,.wraptext=-1"
      _StyleDefs(50)  =   "Named:id=35:Footing"
      _StyleDefs(51)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   "Named:id=36:Selected"
      _StyleDefs(53)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(54)  =   "Named:id=37:Caption"
      _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(56)  =   "Named:id=38:HighlightRow"
      _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(58)  =   "Named:id=39:EvenRow"
      _StyleDefs(59)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(60)  =   "Named:id=40:OddRow"
      _StyleDefs(61)  =   ":id=40,.parent=33"
      _StyleDefs(62)  =   "Named:id=41:RecordSelector"
      _StyleDefs(63)  =   ":id=41,.parent=34"
      _StyleDefs(64)  =   "Named:id=42:FilterBar"
      _StyleDefs(65)  =   ":id=42,.parent=33"
   End
   Begin vbp_Ventas.ctlGrillaArray grdDetalle 
      Height          =   4335
      Left            =   0
      TabIndex        =   13
      Top             =   2400
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   7646
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Li&mpiar"
      Height          =   615
      Left            =   9240
      Picture         =   "frm_ADM_SMMNuevo.frx":05A4
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Elimina los No Solicitados"
      Top             =   1560
      Width           =   1095
   End
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   1058
      ModoBotones     =   10
      EnabledEfecto   =   0   'False
   End
   Begin VB.Frame fraFiltro 
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   10815
      Begin vbp_Ventas.ctlDataCombo cboLaboratorio 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo cboLinea 
         Height          =   315
         Left            =   5400
         TabIndex        =   3
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlTextBox TxtProducto 
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
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
         Caption         =   "BROCODUAL"
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
         TabIndex        =   12
         Top             =   1200
         Width           =   6615
      End
      Begin VB.Label lblCod_Producto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "77700"
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
         TabIndex        =   11
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ACT"
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
         Left            =   7800
         TabIndex        =   10
         Top             =   1200
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "&Producto"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Lí&nea"
         Height          =   255
         Left            =   4800
         TabIndex        =   2
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "&Laboratorio"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm_ADM_SMMNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objSMM As New clsSMM

Public Property Get Solicitud() As clsSMM
    Set Solicitud = objSMM
End Property

Public Property Set Solicitud(ByVal vstrNewValue As clsSMM)
    Set objSMM = vstrNewValue
End Property

Private Sub cboLaboratorio_Change()
    On Error GoTo CtrlErr
    
    CargaLinea cboLaboratorio.BoundText
    LimpiaProducto
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdAñadir_Click()
    On Error GoTo CtrlErr
    
    If lblCod_Producto.Caption = "" And cboLaboratorio.BoundText = "*" Then
        MsgBox "Seleccione criterio de búsqueda, ya sea línea o Producto", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If cboLaboratorio.BoundText <> "*" And cboLinea.BoundText = "*" Then
        MsgBox "Seleccione línea", vbExclamation, "Aviso"
        cboLinea.SetFocus
        Exit Sub
    End If
    
    
    Dim odynTemp As oraDynaset
    
    Set odynTemp = Solicitud.Filtro(Solicitud.CodLocal, _
                                    lblCod_Producto.Caption, _
                                    cboLaboratorio.BoundText, _
                                    cboLinea.BoundText)
    If odynTemp.RecordCount = 0 Then
        MsgBox "No se ubicaron productos activos con criterio de búsqueda", vbExclamation, "Aviso"
        
        If lblCod_Producto.Caption <> "" Then
            TxtProducto.SetFocus
        Else
            cboLinea.SetFocus
        End If
        
    Else
        AdicionaDetalle odynTemp
    End If
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdLimpiar_Click()
    On Error GoTo CtrlErr
    
    If grdDetalle.ApproxCount < 1 Then Exit Sub
    
    If MsgBox("¿Seguro(a) de Eliminar Items que no tienen cantidades de solicitud?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    
    grdDetalle.MoveFirst
''COMENTADO POR ARTURO ESCATE 26/05/2008
    While Not grdDetalle.EOF
        Screen.MousePointer = vbHourglass
        If "" & grdDetalle.Columns(10).Value = "" Then
            grdDetalle.Delete
        Else
            grdDetalle.MoveNext
        End If
        Screen.MousePointer = vbDefault
    Wend
        
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    On Error GoTo CtrlErr
    Select Case boton
        Case Nuevo
        Case Modificar
        Case Buscar
        Case tb_Actualizar
        Case Imprimir
        Case tb_Excel
        Case tb_email
        Case Grabar
            Graba
        Case Cancelar
            Cancela
        Case Eliminar
        Case salir
            
    End Select
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub Form_Load()
    On Error GoTo CtrlErr
    
    SeteaGrilla
    CargaLaboratorio
    LimpiaProducto
    CargaMotivos
    
    Dim objEstVta As New clsEstadistica
    
    objEstVta.Limpia
    
    Set objEstVta = Nothing
    
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub grdDetalle_AfterColUpdate(ByVal ColIndex As Integer)
    grdDetalle.MoveNext
    grdDetalle.MovePrevious
End Sub

Private Sub grdDetalle_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    Select Case ColIndex
        Case 10, 11, 12
            Cancel = False
        Case Else
            Cancel = True
    End Select
End Sub

Private Sub grdDetalle_ButtonClick(ByVal ColIndex As Integer)
    
    On Error GoTo CtrlErr
    
    Select Case ColIndex
        Case 7 'ventas
            
        Dim frm As New frm_ADM_Graf_Ventas
        Dim objDatos As New clsEstadistica
        Dim odynDatos As oraDynaset
        Dim varDaTos(0 To 1, 0 To 9) As Variant
    
            
        Set odynDatos = objDatos.Lista(Solicitud.CodLocal, grdDetalle.Columns(0).Value)
        
        odynDatos.MoveFirst
        
        While Not odynDatos.EOF
                    
            varDaTos(1, odynDatos("ORDEN").Value) = odynDatos("CTD_VENTA").Value
            varDaTos(0, odynDatos("ORDEN").Value) = odynDatos("PERIODO").Value
            odynDatos.MoveNext
            
        Wend
        
        varDaTos(1, 0) = "Cantidad de Venta"
        frm.Titulo = grdDetalle.Columns(0).Value & ": " & grdDetalle.Columns(1).Value
        frm.Datos = varDaTos
        frm.Mostrar
            
        Set frm = Nothing
            
            
    End Select
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub grdDetalle_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
    On Error GoTo CtrlErr
        
    Select Case Col
        Case 1
            If grdDetalle.Columns("Sel.").CellValue(Bookmark) = "1" Then
                CellStyle.ForeColor = vbRed
                CellStyle.Font.Bold = True
            End If
            
        Case 10, 11, 12 'Campos a ingresar
            Select Case Condition
                Case CellStyleConstants.dbgMarqueeRow
                    CellStyle.Font.Bold = True
                Case CellStyleConstants.dbgMarqueeRow + CellStyleConstants.dbgCurrentCell
                    CellStyle.BackColor = vbCyan  'RGB(60, 255, 255)  'vbInfoBackground
                    CellStyle.Font.Bold = True
            End Select
        End Select
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbExclamation, "Error"

End Sub

Private Sub grdDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo CtrlErr
    
    If grdDetalle.ApproxCount < 1 Then Exit Sub
    
    If Not grdDetalle.EditActive Then
    
        If KeyCode = vbKeyDelete Then
            If MsgBox("¿Desea eliminar Item de la lista?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Item") = vbYes Then
                grdDetalle.Delete
                grdDetalle.Refresh
            End If
        End If
    End If
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub grdDetalle_LostFocus()
    grdDetalle.Refresh
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
        
    End If

    Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub Graba()
Dim i As Integer
    
    If grdDetalle.EditActive Then
        grdDetalle.MoveNext
        grdDetalle.MovePrevious
    End If
    
    If grdDetalle.ApproxCount < 1 Then Exit Sub
    
    If MsgBox("¿Seguro(a) de registrar la Solicitud?", vbQuestion + vbYesNo + vbDefaultButton2, "Grabar") = vbNo Then Exit Sub
    
    
    For i = Solicitud.Detalle.LowerBound(1) To Solicitud.Detalle.UpperBound(1)
        If "" & Solicitud.Detalle(i, 10) = "" Then
            MsgBox "Se han encontrado items que no cuentan con cantidad a Solicitar" & Chr(13) & _
                   "Limpiar primero estos items con la opción ""Limpiar""", vbExclamation, "Grabar"
            Exit Sub
        End If
        
        If "" & Solicitud.Detalle(i, 10) <> "" And "" & Solicitud.Detalle(i, 11) = "" Then
            MsgBox "El Motivo no es Opcional, tiene que seleccionar uno", vbExclamation, "Grabar"
            grdDetalle.Bookmark = i
            Exit Sub
        End If
        
    Next i
    
               
    Solicitud.CodUsuario = objUsuario.Codigo
    Solicitud.Grabar
    
    MsgBox "Se registró la Solicitud Número " & Solicitud.Numero, vbInformation, "Grabar"
    
    Unload Me
    'Nuevo
End Sub

Private Sub Cancela()
Dim blnPosible As Boolean
Dim intEncontrado As Integer
    
    blnPosible = False
    
    If grdDetalle.ApproxCount > 1 Then
        blnPosible = True
    End If
    
'    If blnPosible Then
'        intEncontrado = Solicitud.Detalle.Find(0, 10, 1, XORDER_ASCEND, XCOMP_GT, XTYPE_INTEGER)
'
'        If intEncontrado = Solicitud.Detalle.LowerBound(1) - 1 Then
'            blnPosible = False
'        End If
'
'    End If
    
    If blnPosible Then
        If MsgBox("¿Seguro(a) de Salir sin grabar?", vbQuestion + vbYesNo + vbDefaultButton2, "Cancelar") = vbNo Then
            Exit Sub
        End If
    End If
    Unload Me
    
End Sub

Private Sub CargaLinea(ByVal vstrCodLab As String)
    Dim objLinea As New clsLinea
    Set cboLinea.RowSource = objLinea.Lista(vstrCodLab, "1", "[SELECCIONAR]")       '''objProducto.ListaLinea(vstrCodLab)
    Set objLinea = Nothing
    
    cboLinea.ListField = "DES"
    cboLinea.BoundColumn = "COD"
    cboLinea.BoundText = "*"
    
End Sub

Private Sub CargaLaboratorio()
    Dim objLaboratorio As New clsLaboratorio
    Set cboLaboratorio.RowSource = objLaboratorio.Lista("1", "[SELECCIONAR]")
    Set objLaboratorio = Nothing
    
    cboLaboratorio.ListField = "DES"
    cboLaboratorio.BoundColumn = "COD"
    cboLaboratorio.BoundText = "*"
         
End Sub

Private Sub LimpiaProducto()
    lblCod_Producto.Caption = ""
    lblProducto.Caption = ""
    lblEstado.Caption = ""
End Sub

Private Sub CargaMotivos()
 
    drdMotivo.RowHeight = 0
    drdMotivo.RowHeight = drdMotivo.RowHeight * 1.7
    
    
    
    grdDetalle.Columns("Motivo").DropDown = drdMotivo
    drdMotivo.DataField = "COD_MOTIVO_SMM"
    drdMotivo.ListField = "DES_MOTIVO_SMM"
    drdMotivo.AllowRowSizing = False
    drdMotivo.AllowColMove = False
    drdMotivo.EmptyRows = True
        
    Dim objMotivo As New clsMotivoSMM
    
    Set oradcMotivo.Recordset = objMotivo.ListaActivos
    Dim odynClone As oraDynaset
    
    Set odynClone = oradcMotivo.Recordset.Clone
    
    odynClone.MoveFirst
    
    While Not odynClone.EOF
        psub_Grilla_Traslate grdDetalle, 11, odynClone("COD_MOTIVO_SMM").Value, odynClone("DES_MOTIVO_SMM").Value
        odynClone.MoveNext
    Wend
        
    drdMotivo.Height = 0
    'drdMotivo.Height = drdMotivo.RowHeight
    drdMotivo.Height = 3 * drdMotivo.Height + drdMotivo.RowHeight * (IIf(oradcMotivo.Recordset.RecordCount > 8, 8, oradcMotivo.Recordset.RecordCount))
    
    drdMotivo.Appearance = dbgFlat
    drdMotivo.Columns(0).BackColor = RGB(240, 240, 240)
    
    
    Set objMotivo = Nothing
            
    grdDetalle.Columns("Motivo").DropDownList = True

End Sub

Private Sub AdicionaDetalle(ByRef rodynTemp As oraDynaset)
Dim IntRow As Integer
Dim strMensaje As String
Dim blnAdd  As Boolean
Dim btnActualizar As Boolean
Dim inTf As Integer

    rodynTemp.MoveFirst
    strMensaje = ""
    btnActualizar = False
    IntRow = 0
        
    
    While Not rodynTemp.EOF
        Screen.MousePointer = vbHourglass
        'Permite Ubicar en producto en la grilla o de lo contrario encontralo lo pinta'
        blnAdd = False
        If Solicitud.Detalle.UpperBound(1) > Solicitud.Detalle.LowerBound(1) - 1 Then
             inTf = Solicitud.Detalle.Find(0, 0, CStr("" & rodynTemp("COD_PRODUCTO").Value), XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
             If inTf = -1 Then
                   blnAdd = True
             End If
        Else
            blnAdd = True
        End If
        
        If blnAdd Then
            btnActualizar = True
            Solicitud.Detalle.InsertRows IntRow
            Solicitud.Detalle(IntRow, 0) = "" & rodynTemp("COD_PRODUCTO").Value
            Solicitud.Detalle(IntRow, 1) = "" & rodynTemp("DES_PRODUCTO").Value
            Solicitud.Detalle(IntRow, 2) = "" & rodynTemp("LABORATORIO").Value
            Solicitud.Detalle(IntRow, 3) = "" & rodynTemp("LINEA").Value
            Solicitud.Detalle(IntRow, 4) = "" & rodynTemp("FLG_SELECCIONADO").Value
            Solicitud.Detalle(IntRow, 5) = "" & rodynTemp("COD_EST_ABAST").Value
            Solicitud.Detalle(IntRow, 6) = "" & rodynTemp("CTD_DIAS_AGOTADO").Value
            Solicitud.Detalle(IntRow, 7) = "" & rodynTemp("VENTAS").Value
            Solicitud.Detalle(IntRow, 8) = "" & rodynTemp("CTD_QUIEBRES").Value
            Solicitud.Detalle(IntRow, 9) = "" & rodynTemp("MAX_ACTUAL").Value
            Solicitud.Detalle(IntRow, 10) = "" & rodynTemp("MAX_SOLICITADO").Value
            Solicitud.Detalle(IntRow, 11) = "" & rodynTemp("COD_MOTIVO").Value
            Solicitud.Detalle(IntRow, 12) = "" & rodynTemp("OBSERVACION").Value
            Solicitud.Detalle(IntRow, 13) = "" & rodynTemp("DIAS_INVENTARIO").Value
            Solicitud.Detalle(IntRow, 14) = "" & rodynTemp("STOCK").Value
            Solicitud.Detalle(IntRow, 15) = "" & rodynTemp("CATEGORIA").Value
            IntRow = IntRow + 1
        Else
             strMensaje = strMensaje & CStr("" & rodynTemp("COD_PRODUCTO").Value) & " - " & CStr("" & rodynTemp("DES_PRODUCTO").Value) & Chr(13)
        End If
        
        rodynTemp.MoveNext
        Screen.MousePointer = vbDefault
    Wend
        
    
    If btnActualizar = True Then
        grdDetalle.Rebind
        grdDetalle.MoveFirst
        grdDetalle.Col = 10
        grdDetalle.SetFocus
    End If
    
    If strMensaje <> "" Then
        Me.Refresh
        MsgBox "Los siguientes productos ya se encontraban en la lista: " & Chr(13) & strMensaje, vbInformation, "Aviso"
        grdDetalle.Bookmark = inTf
    End If
    
    
    
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim arrFoco As Variant
      
    '---------------------------------------------------------------
    '-- Detalle
    '---------------------------------------------------------------
    arrCampos = Array("Código", "Descripción", "Laboratorio", _
                       "Línea", "Sel.", "Est.Abast.", _
                       "Días Agotado", "Ventas", "Quiebres de Stock", _
                       "Max. Actual", "Max. Solic.", "Motivo", _
                       "Observación", "dias inventario", "Stock", _
                       "Categoria")
    
    arrCaption = Array("Código", "Descripción", "Laboratorio", _
                       "Línea", "Sel.", "Est Abast", _
                       "Días Agot.", "Vnt", "Nro Quieb", _
                       "Máx Actual", "Máx. Solic", "Motivo", _
                       "Observación", "Días Inv", "Stock", _
                       "Categoria Comercial")
    
    arrAncho = Array(650, 2800, 1200, _
                     1200, 400, 400, _
                     500, 400, 500, _
                     600, 600, 1400, _
                     2000, 500, 500, _
                     1200)
    
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, _
                          dbgLeft, dbgCenter, dbgCenter, _
                          dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgCenter, dbgLeft, _
                          dbgLeft, dbgCenter, dbgCenter, _
                          dbgLeft)
                              
    grdDetalle.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    grdDetalle.Columns(13).Order = 8
    
    grdDetalle.HeadLines = 2
    grdDetalle.RowHeight = 0
    grdDetalle.RowHeight = grdDetalle.RowHeight * 2
    grdDetalle.Columns("Vnt").ButtonText = True
    grdDetalle.Columns("Vnt").ButtonAlways = True
    grdDetalle.AllowUpdate = True
    
    grdDetalle.EditorStyle.BackColor = vbWhite 'RGB(242, 242, 252)
    grdDetalle.EditorStyle.ForeColor = RGB(180, 0, 180)
    grdDetalle.EditorStyle.Font.Bold = True
    
    grdDetalle.Columns(10).BackColor = vbInfoBackground
    'grdDetalle.Columns(10).ForeColor = vbBlue
    grdDetalle.Columns(11).BackColor = vbInfoBackground
    'grdDetalle.Columns(11).ForeColor = vbBlue
    grdDetalle.Columns(12).BackColor = vbInfoBackground
    'grdDetalle.Columns(12).ForeColor = vbBlue
    'grdDetalle.Columns("Max. Solic.").Font.Bold = True
    
    
    grdDetalle.Columns(12).EditMask = String(100, "&")
    grdDetalle.Columns(10).EditMask = "####"
    
    
    grdDetalle.Array1 = Solicitud.Detalle
    Solicitud.Detalle.ReDim 0, -1, 0, 15
    
    Dim s As TrueDBGrid70.Split
    Set s = grdDetalle.Splits.Add(1)
    
    With grdDetalle.Splits(0)
        .RecordSelectors = False
        .SizeMode = dbgExact
        .Size = grdDetalle.Columns(0).Width + grdDetalle.Columns(1).Width + grdDetalle.Columns(2).Width
        .AllowSizing = False
    End With
    
    With grdDetalle.Splits(1)
        .RecordSelectors = False
        .AllowSizing = False
        .SizeMode = dbgScalable
    End With
    
    Dim c As Column
    For Each c In grdDetalle.Columns
        If c.ColIndex > 2 Then
            grdDetalle.Splits(0).Columns(c.ColIndex).Visible = False
        Else
            grdDetalle.Splits(1).Columns(c.ColIndex).Visible = False
        End If
    Next c
    
    grdDetalle.Columns("Días Agot.").Visible = False
    
    
    
    grdDetalle.Columns("Descripción").FetchStyle = True
    grdDetalle.Columns("Sel.").Visible = False
    grdDetalle.Columns("Stock").Order = grdDetalle.Columns("Vnt").Order + 1
    
    
    grdDetalle.Columns(10).FetchStyle = True
    grdDetalle.Columns(11).FetchStyle = True
    grdDetalle.Columns(12).FetchStyle = True
    
    
End Sub


