VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_ADM_EntregaDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmrRecontar 
      Caption         =   "Recontar"
      Height          =   375
      Left            =   7920
      TabIndex        =   15
      Top             =   6000
      Width           =   1695
   End
   Begin vbp_Ventas.ctlTextBox txtGlosa 
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      Tipo            =   8
      MaxLength       =   100
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
   Begin vbp_Ventas.ctlTextBox txtChofer 
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      Tipo            =   2
      MaxLength       =   100
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
   Begin vbp_Ventas.ctlDataCombo cboTransportista 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin TrueOleDBGrid70.TDBGrid grdDetalle 
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   3625
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "N° Guia"
      Columns(0).DataField=   "NUM_GUIA"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "N° Pedido"
      Columns(1).DataField=   "NUM_PEDIDO"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "F. Recepción"
      Columns(2).DataField=   "FCH_RECEPCION"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "F. Emisión"
      Columns(3).DataField=   "FCH_EMISION"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   4
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "X"
      Columns(4).DataField=   "FCH_SELECCIONADO"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   2
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=38"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(50)  =   "Named:id=33:Normal"
      _StyleDefs(51)  =   ":id=33,.parent=0"
      _StyleDefs(52)  =   "Named:id=34:Heading"
      _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(54)  =   ":id=34,.wraptext=-1"
      _StyleDefs(55)  =   "Named:id=35:Footing"
      _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(57)  =   "Named:id=36:Selected"
      _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=37:Caption"
      _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(61)  =   "Named:id=38:HighlightRow"
      _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(63)  =   "Named:id=39:EvenRow"
      _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(65)  =   "Named:id=40:OddRow"
      _StyleDefs(66)  =   ":id=40,.parent=33"
      _StyleDefs(67)  =   "Named:id=41:RecordSelector"
      _StyleDefs(68)  =   ":id=41,.parent=34,.bgcolor=&H8000000F&"
      _StyleDefs(69)  =   "Named:id=42:FilterBar"
      _StyleDefs(70)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdIniciar 
      Caption         =   "Iniciar Recepción"
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   6000
      Width           =   1815
   End
   Begin vbp_Ventas.ctlTextBox txtPlaca 
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      Tipo            =   8
      MaxLength       =   6
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
   Begin vbp_Ventas.ctlTextBox txtBultos 
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   1170
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Tipo            =   3
      Alignment       =   2
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
   Begin vbp_Ventas.ctlTextBox txtPrecintos 
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   1200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Tipo            =   3
      Alignment       =   2
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
   Begin vbp_Ventas.ctlGrillaArray ctlgrdguias 
      Height          =   3735
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   6588
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox txtentrega 
      Height          =   495
      Left            =   8040
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Tipo            =   3
      Alignment       =   2
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "F3-Ver Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   120
      TabIndex        =   18
      Top             =   6000
      Width           =   1515
   End
   Begin VB.Label Label6 
      Caption         =   "Glosa"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   405
   End
   Begin VB.Label Label5 
      Caption         =   "Cant. Precintos"
      Height          =   195
      Left            =   3240
      TabIndex        =   13
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label Label4 
      Caption         =   "Cant. Bultos"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Placa Unidad"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Transportista/Chofer"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Empresa Transporte"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1425
   End
End
Attribute VB_Name = "frm_ADM_EntregaDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objEntrega As New clsEntrega
Dim xDetalle As New XArrayDB
Dim strIdEntrega As String

Public Sub carga(identrega As String)
    strIdEntrega = identrega
    xDetalle.ReDim 0, -1, 0, 5
    CargaCombos
    cargaDetalle identrega
    If identrega = "" Then
        Me.Caption = "Guias Pendiente de Recepcionar"
    Else
        Me.Caption = "Entrega N°:" & identrega
        txtentrega.Text = identrega
    End If
    Me.Show vbModal
End Sub

Sub CargaCombos()
'Set cboTransportista.RowSource = objEntrega.ListaTransportista("", "1", "Seleccionar")
'    cboTransportista.BoundColumn = "ID_TRANSPORTISTA"
'    cboTransportista.ListField = "DES_TRANSPORTISTA"
'    cboTransportista.Text = "Seleccionar"
End Sub

Sub cargaDetalle(identrega As String)
On Error GoTo Handle
    Dim rs As oraDynaset
    Dim rsCabecera As oraDynaset
    Dim i As Integer
    Set rs = objEntrega.ListaPendiente(objUsuario.CodigoLocal, identrega, "")
    
    Set rsCabecera = objEntrega.Lista(objUsuario.CodigoLocal, "", "", "", identrega)
    
    If rsCabecera.RecordCount > 0 Then
        If "" & rsCabecera("COD_ESTADO").Value = "CER" Then
            cmdIniciar.Visible = False
            cmrRecontar.Visible = True
        ElseIf "" & rsCabecera("COD_ESTADO").Value = "EMI" Then
            cmdIniciar.Visible = True
            cmrRecontar.Visible = False
        Else
            cmdIniciar.Visible = False
            cmrRecontar.Visible = False
        End If
        txtBultos.Text = "" & rsCabecera("CTD_BULTOS")
        txtChofer.Text = "" & rsCabecera("DES_CHOFER")
        txtPlaca.Text = "" & rsCabecera("DES_PLACA")
        txtGlosa.Text = "" & rsCabecera("DES_GLOSA")
        txtPrecintos.Text = "" & rsCabecera("CTD_PRECINTOS")
        cboTransportista.BoundText = "" & rsCabecera("COD_EMPRESA")
    Else
        cmdIniciar.Visible = True
        cmrRecontar.Visible = False
    End If
        
    i = 0
    While Not rs.EOF
    xDetalle.AppendRows
        xDetalle(i, 0) = rs("FLG_SELECCIONADO").Value * (-1)
        xDetalle(i, 1) = rs("NUM_GUIA").Value
        xDetalle(i, 2) = rs("NUM_PEDIDO").Value
        xDetalle(i, 3) = rs("FCH_RECEPCION").Value
        xDetalle(i, 4) = rs("FCH_EMISION").Value
        
        i = i + 1
        rs.MoveNext
    Wend
    
    'seteagrilla2
    SeteaGrilla
    Me.ctlgrdguias.Array1 = xDetalle
'    grdDetalle.Array = xDetalle
'    grdDetalle.Rebind
'
'    grdDetalle.Columns(0).AllowFocus = False
'    grdDetalle.Columns(1).AllowFocus = False
'    grdDetalle.Columns(2).AllowFocus = False
'    grdDetalle.Columns(3).AllowFocus = False
Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName
   
End Sub

Private Sub cmdDetGuia_Click()
VerDetGuia
End Sub

Private Sub cmdIniciar_Click()
    If BuscarDatosSeleccionados = False Then
    MsgBox "No se ha seleccionado Guía.", vbCritical, "Error"
    Exit Sub
    End If
    If cboTransportista.Text = "Seleccionar" Then
        MsgBox "Falta Seleccionar el Transportista", vbCritical, "Error"
        cboTransportista.SetFocus
        GoTo salir
    End If
    If Trim(txtChofer.Text) = "" Then
        MsgBox "Falta Ingresar el Transportista/Chofer", vbCritical, "Error"
        txtChofer.SetFocus
        GoTo salir
    End If
    If Trim(txtPlaca.Text) = "" Then
        MsgBox "Falta Ingresar la Placa", vbCritical, "Error"
        txtPlaca.SetFocus
        GoTo salir
    End If
    If Trim(txtBultos.Text) = "" Then
        MsgBox "Falta Ingresar el Numero de Bultos", vbCritical, "Error"
        txtBultos.SetFocus
        GoTo salir
    End If
    If Trim(txtPrecintos.Text) = "" Then
        MsgBox "Falta Ingresar el Numero de Precintos", vbCritical, "Error"
        txtPrecintos.SetFocus
        GoTo salir
    End If
    If txtentrega.Text = "" Then
        txtentrega.Text = Graba
    End If
    frm_ADM_EntregaProd.carga Trim(txtentrega.Text), "0"
    Unload Me
salir:
End Sub

Private Function BuscarDatosSeleccionados()

Dim j As Integer
'    grdDetalle.MoveNext
'    grdDetalle.MovePrevious
Me.ctlgrdguias.Update
Me.ctlgrdguias.MoveNext
Me.ctlgrdguias.MovePrevious
    BuscarDatosSeleccionados = False
    For j = xDetalle.LowerBound(1) To xDetalle.UpperBound(1)
        If xDetalle(j, 0) = -1 Then
           j = xDetalle.UpperBound(1) + 1
           BuscarDatosSeleccionados = True
        End If
    Next
End Function

Function Graba() As String
On Error GoTo Handle
Dim i As Integer
Dim Entrega As String
Dim arrGuias As String
arrGuias = ""

While i < xDetalle.Count(1)
    If Val(xDetalle(i, 0)) <> 0 Then
        arrGuias = arrGuias & xDetalle(i, 1) & "|"
    End If
    i = i + 1
Wend

objEntrega.GrabaEntrega Entrega, arrGuias, objUsuario.Codigo, objUsuario.CodigoLocal, cboTransportista.BoundText, Trim(txtChofer.Text), Val(txtPlaca.Text), Val(txtBultos.Text), Trim(txtPrecintos.Text), Trim(txtGlosa.Text)
Graba = Entrega
Exit Function
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Function

Private Sub cmrRecontar_Click()
    If txtentrega.Text = "" Then
        txtentrega.Text = Graba
    End If
    
    frm_ADM_EntregaProd2.carga Trim(txtentrega.Text), "1"
    Unload Me
'    If MsgBox("Desea Cerrar la recepción", vbYesNo, App.ProductName) = vbYes Then
'        objEntrega.CierraRecepcion strIdEntrega, objUsuario.NombrePC, objUsuario.Codigo
'    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim arrFoco As Variant

    arrCampos = Array("FLG_SELECCIONADO", "NUM_GUIA", "NUM_PEDIDO", "FCH_RECEPCION", "FCH_EMISION")
    arrCaption = Array("X", "Nº Guía", "Nº Pedido", "Fec. Recepcion", "Fec. Emision")
    arrAncho = Array(800, 1200, 1300, 2000, 2000, 500)
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgLeft, dbgLeft, dbgCenter)
    arrFoco = Array(True, False, False, False, False)
    Me.ctlgrdguias.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion, arrFoco
    Me.ctlgrdguias.AllowUpdate = True
    
    Me.ctlgrdguias.Columns(0).ValueItems.Presentation = dbgCheckBox
    Me.ctlgrdguias.Columns(4).Merge = False
    Me.ctlgrdguias.Columns(1).Merge = False
    Me.ctlgrdguias.Columns(2).Merge = False
    Me.ctlgrdguias.Columns(3).Merge = False
    
    ctlgrdguias.EditorStyle.BackColor = vbWhite
    ctlgrdguias.EditorStyle.ForeColor = RGB(180, 0, 180)
    ctlgrdguias.EditorStyle.Font.Bold = True
    
    ctlgrdguias.Columns(1).BackColor = vbInfoBackground
    ctlgrdguias.Columns(2).BackColor = vbInfoBackground

End Sub

Private Sub ctlgrdguias_DblClick()
VerDetGuia
End Sub

Sub VerDetGuia()
Dim mensaje As String
mensaje = Me.ctlgrdguias.Columns(1).Value
frm_ADM_DetGuias.numGuia = mensaje
frm_ADM_DetGuias.Show vbModal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
    VerDetGuia
End If
End Sub
