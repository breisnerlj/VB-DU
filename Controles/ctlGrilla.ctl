VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.UserControl ctlGrilla 
   ClientHeight    =   3510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   DrawWidth       =   51
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3510
   ScaleWidth      =   5250
   Begin ORADCLibCtl.ORADC ORADC1 
      Height          =   375
      Left            =   1380
      Top             =   3060
      Visible         =   0   'False
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   661
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
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   4080
      Top             =   2460
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   4680
      Top             =   2460
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSComDlg.CommonDialog dlgArchivo 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Abrir archivo de beneficiarios"
      Filter          =   "Archivos de Excel (*.xls)|*.xls"
   End
   Begin TrueDBGrid70.TDBGrid grdDatos 
      Bindings        =   "ctlGrilla.ctx":0000
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6165
      _LayoutType     =   0
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   0
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
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=-1,.fontsize=825,.italic=0"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0"
      _StyleDefs(46)  =   ":id=33,.charset=0"
      _StyleDefs(47)  =   ":id=33,.fontname=Arial"
      _StyleDefs(48)  =   "Named:id=34:Heading"
      _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H80000015&,.fgcolor=&H80000014&"
      _StyleDefs(50)  =   ":id=34,.wraptext=-1"
      _StyleDefs(51)  =   "Named:id=35:Footing"
      _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   "Named:id=36:Selected"
      _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=37:Caption"
      _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(57)  =   "Named:id=38:HighlightRow"
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H800000&,.fgcolor=&H80000014&,.bold=0,.fontsize=825"
      _StyleDefs(59)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(60)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(61)  =   "Named:id=39:EvenRow"
      _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(63)  =   "Named:id=40:OddRow"
      _StyleDefs(64)  =   ":id=40,.parent=33"
      _StyleDefs(65)  =   "Named:id=41:RecordSelector"
      _StyleDefs(66)  =   ":id=41,.parent=34"
      _StyleDefs(67)  =   "Named:id=42:FilterBar"
      _StyleDefs(68)  =   ":id=42,.parent=33"
   End
   Begin TrueDBGrid70.TDBGrid GrdArray 
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6165
      _LayoutType     =   0
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   0
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
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0"
      _StyleDefs(46)  =   ":id=33,.charset=0"
      _StyleDefs(47)  =   ":id=33,.fontname=Arial"
      _StyleDefs(48)  =   "Named:id=34:Heading"
      _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   ":id=34,.wraptext=-1"
      _StyleDefs(51)  =   "Named:id=35:Footing"
      _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   "Named:id=36:Selected"
      _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=37:Caption"
      _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(57)  =   "Named:id=38:HighlightRow"
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&HFEEBDE&,.fgcolor=&H80000012&,.bold=0,.fontsize=825"
      _StyleDefs(59)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(60)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(61)  =   "Named:id=39:EvenRow"
      _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(63)  =   "Named:id=40:OddRow"
      _StyleDefs(64)  =   ":id=40,.parent=33"
      _StyleDefs(65)  =   "Named:id=41:RecordSelector"
      _StyleDefs(66)  =   ":id=41,.parent=34"
      _StyleDefs(67)  =   "Named:id=42:FilterBar"
      _StyleDefs(68)  =   ":id=42,.parent=33"
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuExcel 
         Caption         =   "Exportar a Excel"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "Enviar por E-mail"
      End
   End
End
Attribute VB_Name = "ctlGrilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Click()
Public Event DblClick()
Public Event RegistroSeleccionado(ByVal DatoColumna0 As String)
Public Event DblClickRegistro(ByVal DatoColumna0 As String)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
'Public Event CambiaSeleccionado(ByVal Color As Variant)
Public Event KeyPress(KeyAscii As Integer)

Public Event FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)
Public Event FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)

Public Event ButtonClick(ByVal ColIndex As Integer)
 
Private strClave As String, lngRegistros As Long
Private blnModoSort As Boolean, blnMenuPopUp As Boolean, blnResalte As Boolean
Private TipoGrilla As Boolean

Public Property Get Styles() As TrueDBGrid70.Styles
   Set Styles = grdDatos.Styles
End Property
Public Property Let Styles(ByVal vNewValue As TrueDBGrid70.Styles)
    grdDatos.Styles.vNewValue
End Property

Public Property Set DataSource(ByVal rs As OracleInProcServer.oraDynaset)
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Set ORADC1.Recordset = rs
    'lngRegistros = rs.RecordCount
    Set rs = Nothing
    Screen.MousePointer = vbNormal
End Property

Public Property Get DataSource() As OracleInProcServer.oraDynaset
    Set DataSource = ORADC1.Recordset
End Property

Private Sub GrdArray_DblClick()
    RaiseEvent DblClick
    RaiseEvent DblClickRegistro(strClave)
End Sub

Private Sub grdDatos_ButtonClick(ByVal ColIndex As Integer)
    RaiseEvent ButtonClick(ColIndex)
End Sub

Public Function CambiaSeleccionado(Fondo As Variant)
    grdDatos.Styles(5).Font.Bold = True
    grdDatos.Styles(5).Font.Size = 12
    grdDatos.Styles(5).BackColor = Fondo
    'RaiseEvent CambiaSeleccionado(Fondo)
End Function

Private Sub grdDatos_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
    RaiseEvent FetchCellStyle(Condition, Split, Bookmark, col, CellStyle)
End Sub

Private Sub grdDatos_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)
    RaiseEvent FetchRowStyle(Split, Bookmark, RowStyle)
End Sub

Private Sub grdDatos_GotFocus()
    If blnResalte = False Then
        'grdDatos.Styles(5).BackColor = &HFFD7D7 ' &HFFEAEA
        'grdDatos.Styles(5).Font.Bold = True
    End If
End Sub

Private Sub grdDatos_LostFocus()
    If blnResalte = False Then
        'grdDatos.Styles(5).BackColor = &HFEE9E9
        'grdDatos.Styles(5).Font.Bold = False
    End If
End Sub

'En esta seccion se va disparar un evento cada vez que se seleccione una
'nueva fila y retorna el dato de la columna0
Private Sub grdDatos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    strClave = grdDatos.Columns(0).Text
    RaiseEvent RegistroSeleccionado(strClave)
End Sub

Public Sub ExportToFile(ByVal Path As String, Append As Boolean)
    grdDatos.ExportToFile Path, Append
End Sub

Public Function IsSelected(ByVal Bookmark As Variant) As Integer
    IsSelected = grdDatos.IsSelected(Bookmark)
End Function

Public Sub FormatoGrilla(ByVal arrayDataField As Variant, ByVal arrayCaption As Variant, _
                         ByVal arrarWidth As Variant, ByVal arrayAlignment As Variant, _
                         Optional ByVal arrayFoco As Variant)
    Dim i As Byte
    Dim columna As TrueDBGrid70.Column

    grdDatos.Columns.Clear 'Quita todas las columnas
    GrdArray.Columns.Clear 'Quita todas las columnas
    For i = 0 To UBound(arrayDataField)
        If TipoGrilla = True Then
            Set columna = GrdArray.Columns.Add(i)
        Else
            Set columna = grdDatos.Columns.Add(i)
        End If
        
        If Not IsMissing(arrarWidth) Then columna.Width = arrarWidth(i)
        If Not IsMissing(arrayAlignment) Then columna.Alignment = arrayAlignment(i)
        If TipoArray = False Then
            If Not IsMissing(arrayDataField) Then columna.DataField = arrayDataField(i)
        End If
        columna.Caption = arrayCaption(i)
        columna.AllowSizing = True
        columna.WrapText = True
        columna.Visible = True
    Next i
    grdDatos.AllowAddNew = False
    GrdArray.AllowAddNew = False
    grdDatos.AllowUpdate = False
    GrdArray.AllowUpdate = False
    
    grdDatos.HoldFields
    GrdArray.HoldFields
    grdDatos.Splits(0).AllowColSelect = False
    GrdArray.Splits(0).AllowColSelect = False
    grdDatos.Splits(0).AllowRowSelect = True
    GrdArray.Splits(0).AllowRowSelect = True
    grdDatos.Splits(0).AllowRowSizing = False
    GrdArray.Splits(0).AllowRowSizing = False
    grdDatos.Splits(0).AllowSizing = False
    GrdArray.Splits(0).AllowSizing = False
    grdDatos.Splits(0).Style.VerticalAlignment = dbgVertCenter
    GrdArray.Splits(0).Style.VerticalAlignment = dbgVertCenter
End Sub

Public Function Limpiar()
    grdDatos.Close True
End Function

Public Property Get Columns() As TrueDBGrid70.Columns
   Set Columns = grdDatos.Columns
End Property

Public Property Let Columns(ByVal vNewValue As TrueDBGrid70.Columns)
    grdDatos.Columns.vNewValue
End Property

Public Property Let HeadLines(vHeadLines As String)
    grdDatos.HeadLines = vHeadLines
End Property

'------------ Eventos clasicos -------------------
Private Sub grdDatos_DblClick()
    RaiseEvent DblClick
    RaiseEvent DblClickRegistro(strClave)
End Sub

Private Sub grdDatos_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub grdDatos_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub grdDatos_Click()
    RaiseEvent Click
End Sub

'-------------------------------------------------
Public Sub Refresh()
    grdDatos.Refresh
End Sub

Public Sub Rebind()
    grdDatos.Rebind
End Sub

Public Sub MoveFirst()
    grdDatos.MoveFirst
End Sub

Public Sub MoveLast()
    grdDatos.MoveLast
End Sub

Public Sub MoveNext()
    grdDatos.MoveNext
End Sub

Public Sub MovePrevious()
    grdDatos.MovePrevious
End Sub
'----------- Propiedades -------------------------------------
Public Property Get Clave() As String
Attribute Clave.VB_Description = "Retorna el valor de la columna cero para un registro. Solo Lectura"
    Clave = strClave
End Property

Public Property Get Bookmark() As Variant
    Bookmark = grdDatos.Bookmark
End Property
Public Property Let Bookmark(ByVal vNewValue As Variant)
    grdDatos.Bookmark = vNewValue
End Property

Public Property Get row() As Integer
    row = grdDatos.row
End Property
Public Property Let row(ByVal vNewValue As Integer)
    grdDatos.row = vNewValue
End Property

Public Property Get col() As Integer
    col = grdDatos.col
End Property
Public Property Let col(ByVal vNewValue As Integer)
    grdDatos.col = vNewValue
End Property

Public Property Get Caption() As String
    Caption = grdDatos.Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
    grdDatos.Caption = vNewValue
End Property

Public Property Get FColorCaption() As Variant
    FColorCaption = grdDatos.ForeColor
End Property

Public Property Let FColorCaption(ByVal vNewValue As Variant)
    grdDatos.ForeColor = vNewValue
End Property

Public Property Get BColorCaption() As Variant
    BColorCaption = grdDatos.BackColor
End Property

Public Property Let BColorCaption(ByVal vNewValue As Variant)
    grdDatos.BackColor = vNewValue
End Property
'--------------------------------------------------------------'
'--- cambios por crueda 11/10/2006 ---'
Public Property Get RowHeight()
    RowHeight = grdDatos.RowHeight
    If TipoGrilla = True Then RowHeight = GrdArray.RowHeight
End Property

Public Property Let RowHeight(ByVal vNewValue As Variant)
    grdDatos.RowHeight = vNewValue
    GrdArray.RowHeight = vNewValue
End Property

'--------------------------------------------------------------'
'Public Property Get StylesFondo()
'     StylesFondo = grdDatos.Styles(5).Font.Bold
'     If TipoGrilla = True Then StylesFondo = GrdArray.Styles(5).Font.Bold
'End Property
'
'Public Property Let StylesFondo(ByVal vNewValue As Variant)
'    grdDatos.Styles(5).Font.Bold = vNewValue
'    GrdArray.Styles(5).Font.Bold = vNewValue
'End Property

Public Property Get MultiSelect() As TrueDBGrid70.MultiSelectConstants
    MultiSelect = grdDatos.MultiSelect
End Property

Public Property Let MultiSelect(ByVal vNewValue As TrueDBGrid70.MultiSelectConstants)
    grdDatos.MultiSelect = vNewValue
End Property

Public Property Get ApproxCount() As Long
    ApproxCount = grdDatos.ApproxCount
End Property

Public Property Let ApproxCount(ByVal vNewValue As Long)
    grdDatos.ApproxCount = vNewValue
End Property

Public Property Get Enabled() As Boolean
    Enabled = grdDatos.Enabled
End Property
Public Property Let Enabled(ByVal vNewValue As Boolean)
    grdDatos.Enabled = vNewValue
End Property

Public Property Get MenuPopUp() As Boolean
Attribute MenuPopUp.VB_Description = "Activa o Desactiva el Menu contextual que muestra la opciones de: Excel, Imprimir y e-mail"
    MenuPopUp = blnMenuPopUp
End Property
Public Property Let MenuPopUp(ByVal vNewValue As Boolean)
    blnMenuPopUp = vNewValue
End Property

'Public Property Get Registros() As Long
'    Registros = grdDatos.ApproxCount
'End Property

Public Property Get Resalte() As Boolean
Attribute Resalte.VB_Description = "Si es False cuando el control pierde el foco el texto de la selección se muestra sin resalte. True es el dafault."
    Resalte = blnResalte
End Property
Public Property Let Resalte(ByVal vNewValue As Boolean)
    blnResalte = vNewValue
End Property

Public Property Get ColumnHeaders() As Boolean
    ColumnHeaders = grdDatos.ColumnHeaders
    If TipoGrilla = True Then ColumnHeaders = GrdArray.ColumnHeaders
End Property
Public Property Let ColumnHeaders(ByVal vNewValue As Boolean)
    grdDatos.ColumnHeaders = vNewValue
    GrdArray.ColumnHeaders = vNewValue
End Property

Public Property Get ColumnFooter() As Boolean
    ColumnFooter = grdDatos.ColumnFooters
End Property

Public Property Let ColumnFooter(ByVal vNewValue As Boolean)
    grdDatos.ColumnFooters = vNewValue
End Property

'---------------------------------------------------------------
Private Sub grdDatos_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton And blnMenuPopUp Then
        PopupMenu mnuMenu
    End If
End Sub

Private Sub mnuExcel_Click()
    On Error GoTo Control
    Dim ret As Variant
    dlgArchivo.CancelError = True
    dlgArchivo.DialogTitle = "Ingrese el nombre del archivo."
    dlgArchivo.ShowSave
    ret = dlgArchivo.FileName
    If Trim(ret) = "" Then MsgBox "No se encontro el archivo.", vbOKOnly + vbExclamation, App.ProductName: Exit Sub
    If Dir(ret) <> "" Then
        If MsgBox("Ya existe un archivo con el nombre especificado." & Chr(13) & "Desea continuar ?", vbYesNo + vbQuestion, App.ProductName) = vbNo Then Exit Sub
    End If
    If TipoGrilla = True Then
        GrdArray.ExportToFile ret, False
    Else
        grdDatos.ExportToFile ret, False
    End If
    MsgBox "Se exportó el archivo correctamente.", vbOKOnly + vbInformation, App.ProductName
    Exit Sub
Control:
    Select Case Err.Number
    Case 32755
    Case Else
        MsgBox Err.Description, vbOKOnly + vbExclamation, App.ProductName
    End Select
End Sub

Private Sub mnuImprimir_Click()
If TipoGrilla = True Then
    GrdArray.PrintInfo.PrintPreview
Else
    grdDatos.PrintInfo.PrintPreview
End If
End Sub

Public Property Let PrintInfo(ByVal mvar As TrueDBGrid70.PrintInfo)
    grdDatos.PrintInfo.mvar
End Property

Public Property Get PrintInfo() As TrueDBGrid70.PrintInfo
    Set PrintInfo = grdDatos.PrintInfo
End Property
Public Property Let TipoArray(ByVal IsArray As Boolean)
    TipoGrilla = IsArray
    If TipoGrilla = True Then
        grdDatos.Visible = False
        GrdArray.Visible = True
    Else
        grdDatos.Visible = True
        GrdArray.Visible = False
    End If
End Property

Public Property Get TipoArray() As Boolean
    TipoArray = TipoGrilla
End Property

Public Property Let Array1(ByVal Xarray As XArrayDB)
    GrdArray.Array = Xarray
    GrdArray.Rebind
End Property

Public Property Get Array1() As XArrayDB
    Set Array1 = GrdArray.Array
End Property

Public Property Get SelBookmarks() As SelBookmarks
    Set SelBookmarks = grdDatos.SelBookmarks
End Property

'Public Property Let Columns(ByVal vNewValue As TrueDBGrid70.Columns)
'    grdDatos.Columns.vNewValue
'End Property

Private Sub mnuEmail_Click()
    On Error Resume Next
    With MAPISession1
       .SignOff
       .LogonUI = True
       .DownLoadMail = False
       .SignOn
       .NewSession = True
        MAPIMessages1.SessionID = .SessionID
     MAPIMessages1.SessionID = .SessionID
    End With
    
    With MAPIMessages1
        grdDatos.ExportToFile "c:\temp.xls", False
        .Compose
        '.RecipAddress = ""
        .AddressResolveUI = True
        .AttachmentPathName = "c:\temp.xls"
        .ResolveName
        '.MsgSubject = ""
        .MsgNoteText = " "
        .Send True
        Kill "c:\temp.xls" 'Borra archivo temporal
    End With
End Sub

Public Sub MostrarExcel()
    mnuExcel_Click
End Sub

Public Sub MostrarEmail()
    mnuEmail_Click
End Sub

Public Sub MostrarImprimir()
    mnuImprimir_Click
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", grdDatos.Enabled, True
    PropBag.WriteProperty "MenuPopUp", blnMenuPopUp, True
    PropBag.WriteProperty "Resalte", blnResalte, True
    PropBag.WriteProperty "ColumnHeaders", grdDatos.ColumnHeaders, True
    PropBag.WriteProperty "MultiSelect", grdDatos.MultiSelect, 1
 '   PropBag.WriteProperty "MarqueeStyle", grdDatos.MarqueeStyle, 5
 '   PropBag.WriteProperty "DataMode", grdDatos.DataMode, dbgUnboundSt
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    grdDatos.Enabled = PropBag.ReadProperty("Enabled", True)
    blnMenuPopUp = PropBag.ReadProperty("MenuPopUp", True)
    blnResalte = PropBag.ReadProperty("Resalte", True)
    grdDatos.ColumnHeaders = PropBag.ReadProperty("ColumnHeaders", True)
    grdDatos.MultiSelect = PropBag.ReadProperty("MultiSelect", 1)
'    grdDatos.MarqueeStyle = PropBag.ReadProperty("MarqueeStyle", 5)
'    grdDatos.DataMode = PropBag.ReadProperty("DataMode", dbgUnboundSt)
End Sub

Private Sub UserControl_Resize()
    grdDatos.Width = UserControl.Width
    grdDatos.Height = UserControl.Height
    GrdArray.Width = UserControl.Width
    GrdArray.Height = UserControl.Height
End Sub

'codigo para proporcionar la capacidad de ordenar las columnas
'El performance es malo
'ADO  en este sentido es mucho mejor !!!!!!!

''Private Sub grdDatos_HeadClick(ByVal ColIndex As Integer)
''    If blnModoSort Then
''        pxdbDatos.QuickSort pxdbDatos.LowerBound(1), pxdbDatos.UpperBound(1), ColIndex, XORDER_ASCEND, XTYPE.XTYPE_DEFAULT
''    Else
''        pxdbDatos.QuickSort pxdbDatos.LowerBound(1), pxdbDatos.UpperBound(1), ColIndex, XORDER_DESCEND, XTYPE.XTYPE_DEFAULT
''    End If
''    blnModoSort = Not blnModoSort
''    grdDatos.Refresh
''End Sub
''
''
''Public Property Set DataSource(ByVal rs As OracleInProcServer.OraDynaset)
''    On Error Resume Next
''    Dim cols As Long, i As Long, j As Long
''    pxdbDatos.ReDim 0, rs.RecordCount - 1, 0, rs.Fields.Count
''    cols = rs.Fields.Count
''    Do While Not rs.EOF
''        For i = 0 To cols
''            pxdbDatos(j, i) = rs.Fields(i).Value
''        Next
''        j = j + 1
''        rs.MoveNext
''    Loop
''    Set grdDatos.Array = pxdbDatos
''    grdDatos.Rebind
''    grdDatos.Refresh
''    Registros = rs.RecordCount
''    Set rs = Nothing
''End Property

Public Property Get FetchRowStyle() As Boolean
    FetchRowStyle = grdDatos.FetchRowStyle
End Property

Public Property Let FetchRowStyle(ByVal vNewValue As Boolean)
    grdDatos.FetchRowStyle = vNewValue
End Property

Public Property Get MarqueeStyle() As MarqueeStyleConstants
    MarqueeStyle = grdDatos.MarqueeStyle
End Property

Public Property Let MarqueeStyle(ByVal vNewValue As MarqueeStyleConstants)
    grdDatos.MarqueeStyle = vNewValue
End Property

Public Property Get RowDividerStyle() As DividerStyleConstants
    RowDividerStyle = grdDatos.RowDividerStyle
End Property

Public Property Let RowDividerStyle(ByVal vNewValue As DividerStyleConstants)
    grdDatos.RowDividerStyle = vNewValue
End Property

Public Property Get PrimerRegistro() As Boolean
    PrimerRegistro = grdDatos.BOF
End Property

Public Property Let AlternatingRowStyle(ByVal vNewValue As Boolean)
    grdDatos.AlternatingRowStyle = vNewValue
End Property

'---- Creado por CCIEZA 25/09/2009
'----------------------------------------------------------------
Public Property Get CheckedAll(ByVal varCol As Variant) As CheckBoxConstants
Dim varChecked  As CheckBoxConstants
Dim xdb As XArrayDB
Dim intIndex As Integer
Dim i As Integer
        
    If GrdArray.ApproxCount = 0 Then
        varChecked = vbUnchecked
    Else
            
        intIndex = GrdArray.Columns(varCol).ColIndex
        
        varChecked = vbGrayed
        
        Set xdb = GrdArray.Array
        For i = xdb.LowerBound(1) To xdb.UpperBound(1)
            If Abs(xdb(i, intIndex)) = 1 And varChecked = vbGrayed Then
                varChecked = vbChecked
            ElseIf Not (Abs(xdb(i, intIndex)) = 1) And varChecked = vbGrayed Then
                varChecked = vbUnchecked
            ElseIf Abs(xdb(i, intIndex)) = 1 And varChecked = vbChecked Then
                varChecked = vbChecked
            ElseIf Not (Abs(xdb(i, intIndex)) = 1) And varChecked = vbUnchecked Then
                varChecked = vbUnchecked
            Else
                varChecked = vbGrayed
                Exit For
            End If
        Next i
        
    End If
    CheckedAll = varChecked
End Property

Public Property Let CheckedAll(ByVal varCol As Variant, ByVal vChecked As CheckBoxConstants)
Dim xdb As XArrayDB
Dim intIndex As Integer
Dim i As Integer
    
    If vChecked = vbChecked Or vChecked = vbUnchecked Then
        intIndex = GrdArray.Columns(varCol).ColIndex
        
        'GrdArray.MoveNext
        'GrdArray.MovePrevious
        
        Set xdb = GrdArray.Array
        
        For i = xdb.LowerBound(1) To xdb.UpperBound(1)
            xdb(i, intIndex) = IIf(vChecked = vbChecked, "1", "0")
        Next i
        
        'GrdArray.Array = xdb
        
        GrdArray.Columns(varCol).Refetch
    End If
End Property
