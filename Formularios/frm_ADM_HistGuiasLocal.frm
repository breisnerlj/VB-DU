VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_ADM_HistGuiasLocal 
   BorderStyle     =   0  'None
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7005
   Icon            =   "frm_ADM_HistGuiasLocal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin ORADCLibCtl.ORADC oradcDetalle 
      Height          =   255
      Left            =   6360
      Top             =   6480
      Visible         =   0   'False
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
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
   Begin ORADCLibCtl.ORADC oradcCabecera 
      Height          =   255
      Left            =   5760
      Top             =   3000
      Visible         =   0   'False
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
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
   Begin TrueDBGrid70.TDBGrid grdDetalle 
      Bindings        =   "frm_ADM_HistGuiasLocal.frx":0442
      Height          =   3075
      Left            =   0
      TabIndex        =   1
      Top             =   3480
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   5424
      _LayoutType     =   0
      _RowHeight      =   -2147483647
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
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin TrueDBGrid70.TDBGrid grdCabecera 
      Bindings        =   "frm_ADM_HistGuiasLocal.frx":045D
      Height          =   2835
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7000
      _ExtentX        =   12356
      _ExtentY        =   5001
      _LayoutType     =   0
      _RowHeight      =   -2147483647
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
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=101,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Named:id=33:Normal"
      _StyleDefs(39)  =   ":id=33,.parent=0"
      _StyleDefs(40)  =   "Named:id=34:Heading"
      _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(42)  =   ":id=34,.wraptext=-1"
      _StyleDefs(43)  =   "Named:id=35:Footing"
      _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(45)  =   "Named:id=36:Selected"
      _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(47)  =   "Named:id=37:Caption"
      _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(49)  =   "Named:id=38:HighlightRow"
      _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=39:EvenRow"
      _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(53)  =   "Named:id=40:OddRow"
      _StyleDefs(54)  =   ":id=40,.parent=33"
      _StyleDefs(55)  =   "Named:id=41:RecordSelector"
      _StyleDefs(56)  =   ":id=41,.parent=34"
      _StyleDefs(57)  =   "Named:id=42:FilterBar"
      _StyleDefs(58)  =   ":id=42,.parent=33"
   End
   Begin MSComctlLib.ImageList IlsImagen 
      Left            =   7320
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistGuiasLocal.frx":0479
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistGuiasLocal.frx":0A13
            Key             =   "buscar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistGuiasLocal.frx":0FAD
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistGuiasLocal.frx":1547
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistGuiasLocal.frx":1AE1
            Key             =   "Chek"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistGuiasLocal.frx":207B
            Key             =   "Bien"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistGuiasLocal.frx":2615
            Key             =   "Agregar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistGuiasLocal.frx":2BAF
            Key             =   "Atras"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistGuiasLocal.frx":3149
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistGuiasLocal.frx":36E3
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistGuiasLocal.frx":3C7D
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistGuiasLocal.frx":4217
            Key             =   "Hora"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   2
      Top             =   6570
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   1111
      ButtonWidth     =   1349
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "IlsImagen"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Mostrar"
            Key             =   "Mostrar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Key             =   "Salir"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   0
      Picture         =   "frm_ADM_HistGuiasLocal.frx":47B1
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Servicios Generales-Historial de Guias"
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
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   4140
   End
End
Attribute VB_Name = "frm_ADM_HistGuiasLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objSSGG As New clsSSGG




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
        sub_sge_llena_grilla_cabecera
    Case vbKeyEscape
        mdiPrincipal.picComandos.Enabled = True
        Unload Me
End Select
End Sub

Private Sub Form_Load()
    On Error GoTo Control
    Me.Top = 0
    Me.Left = 0
    redimensionaForm
    'Me.Appearance = 0
    Screen.MousePointer = vbHourglass
    'Call sub_sge_Centra_Ventana(Me)
    spSetGrdCabecera grdCabecera
    spSetGrdDetalle grdDetalle
    Screen.MousePointer = vbNormal
    Exit Sub
Control:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub

Private Sub sub_sge_llena_grilla_cabecera()
    Dim StrSql As String
    Dim oraDynaset As oraDynaset
    On Error GoTo Control
    Screen.MousePointer = vbHourglass
    'StrSql = "SELECT DECODE(CA.EST_GUIA,'EMI','----------',CA.NUM_GUIA) NUM_GUIA,CA.FCH_EMISION,CA.EST_GUIA,CA.FLG_ENVIO_CD, " & _
             "U.DES_NOMBRE || ' ' || U.APE_PAT_USUARIO || ' ' || U.APE_MAT_USUARIO USU_EMISOR,CA.FCH_RECEPCION, " & _
             "U1.DES_NOMBRE || ' ' || U1.APE_PAT_USUARIO || ' ' || U1.APE_MAT_USUARIO USU_RECEPTOR,CA.OBS_EMISION " & _
             "FROM SSGG.CAB_GUIA_SGE CA,NUEVO.MAE_USUARIO_BTL" & _
             " U, NUEVO.MAE_USUARIO_BTL U1 " & _
             "WHERE U.COD_USUARIO=CA.USU_EMISION AND U1.COD_USUARIO (+)= CA.USU_RECEPCION AND " & _
             "CA.COD_DESTINO='" & gstrCodAreaUsuario & "' ORDER BY 2 DESC"
    'Set oraDynaset = godbVentas.CreateDynaset(StrSql, ORADYN_READONLY)
    Set oraDynaset = objSSGG.ListaHistorial(gstrCodAreaUsuario)
    '--------------------------------------
    Set oradcCabecera.Recordset = oraDynaset
    Screen.MousePointer = vbNormal
    Exit Sub
Control:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub

Private Sub sub_sge_llena_grilla_detalle(ByVal codGuia As String)
    Dim StrSql As String
    Dim oraDynaset As oraDynaset
    On Error GoTo Control
    Screen.MousePointer = vbHourglass
    'StrSql = "SELECT D.COD_PRODUCTO,P.DES_PRODUCTO,P.COD_UNID_CONSUMO,C.SIG_UNID_CONSUMO," & _
             "P.FLG_FRACCIONAMIENTO , P.CTD_FRACCIONAMIENTO, D.CTD_PRODUCTO, D.CTD_PRODUCTO_FRAC " & _
             "FROM SSGG.DET_GUIA_SGE D,SSGG.MAE_PRODUCTO_SGE" & _
             " P,SSGG.MAE_UNIDAD_CONSUMO  C " & _
             "WHERE D.COD_PRODUCTO = P.COD_PRODUCTO And P.COD_UNID_CONSUMO = C.COD_UNID_CONSUMO " & _
             "AND D.NUM_GUIA ='" & codGuia & "'"
    'Set oraDynaset = godbVentas.CreateDynaset(StrSql, ORADYN_READONLY)
    Set oraDynaset = objSSGG.ListaDetalleHistorial(codGuia)
    Set oradcDetalle.Recordset = oraDynaset
    Screen.MousePointer = vbNormal
    Exit Sub
Control:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub

Private Sub spSetGrdCabecera(ByRef rgrd As TDBGrid)
    Dim pvarAncho As Variant
    Dim pvarTitulo As Variant
    Dim pvarAlinea As Variant
    Dim pvarCampoDato As Variant
 
   pvarAncho = Array(1200, 1200, 1000, 400, 2000, 1400, 2000, 4000)
   pvarTitulo = Array("# Guía", "Fecha Emisión", "Estado", "Vía CD", "Emisor", "Fecha Recepción", "Receptor", "Observación")
   pvarAlinea = Array(2, 2, 2, 2, 0, 2, 0, 0)
   pvarCampoDato = Array("NUM_GUIA", "FCH_EMISION", "EST_GUIA", "FLG_ENVIO_CD", _
                         "USU_EMISOR", "FCH_RECEPCION", "USU_RECEPTOR", "OBS_EMISION")
    
   objSSGG.spGrilla_Carga rgrd, pvarTitulo, pvarAncho, pvarAlinea, pvarCampoDato
   objSSGG.spGrilla_Traslate rgrd, "EST_GUIA", "EMI", "EMITIDA"
   objSSGG.spGrilla_Traslate rgrd, "EST_GUIA", "ANU", "ANULADA"
   objSSGG.spGrilla_Traslate rgrd, "EST_GUIA", "REC", "RECIBIDA"
   objSSGG.spGrilla_Traslate rgrd, "FLG_ENVIO_CD", "1", "SI"
   objSSGG.spGrilla_Traslate rgrd, "FLG_ENVIO_CD", "0", "NO"
        
   rgrd.MarqueeStyle = dbgHighlightRow
   rgrd.HeadBackColor = &H8000000F '&H80000007
   rgrd.HeadForeColor = &H80000012 '&HFFFF&
'   rgrd.HeadBackColor = &H80000007
'   rgrd.HeadForeColor = &HFFFF&
   rgrd.RowHeight = 0
   rgrd.RowHeight = 420
   rgrd.HeadLines = 2
   rgrd.Font.Size = 8
   rgrd.Styles(5).Font.Size = 8

    rgrd.Columns("NUM_GUIA").Style.BackColor = RGB(254, 244, 207)
    rgrd.Columns("NUM_GUIA").Style.ForeColor = RGB(202, 24, 4)
   
    rgrd.Columns("FCH_EMISION").Style.BackColor = RGB(254, 244, 207)
    'rgrd.Columns("FCH_EMISION").Style.Font.Bold = True
   
    rgrd.Columns("EST_GUIA").Style.BackColor = RGB(254, 244, 207)
    rgrd.Columns("EST_GUIA").Style.ForeColor = RGB(202, 24, 4)
    'rgrd.Columns("EST_ORDEN_COMPRA").Style.Font.Bold = True

End Sub

Private Sub spSetGrdDetalle(ByRef rgrd As TDBGrid)

    Dim pvarAncho As Variant
    Dim pvarTitulo As Variant
    Dim pvarAlinea As Variant
    Dim pvarCampoDato As Variant
 
   pvarAncho = Array(300, 1000, 5700, 1000, 700, 1000, 800, 800, 800)
   pvarTitulo = Array("I", "Código Producto", "Producto", "Código Unidad", "Unidad Consumo", "Fracciona?", "Ctd. Frac.", "Unid. Enviadas", "Frac. Enviadas")
   pvarAlinea = Array(2, 2, 0, 2, 0, 2, 1, 1, 1)
   pvarCampoDato = Array("ITEM", "COD_PRODUCTO", "DES_PRODUCTO", "COD_UNID_CONSUMO", "SIG_UNID_CONSUMO", "FLG_FRACCIONAMIENTO", "CTD_FRACCIONAMIENTO", "CTD_PRODUCTO", "CTD_PRODUCTO_FRAC")
    
   objSSGG.spGrilla_Carga rgrd, pvarTitulo, pvarAncho, pvarAlinea, pvarCampoDato
   objSSGG.spGrilla_Traslate rgrd, "FLG_FRACCIONAMIENTO", "1", "SI"
   objSSGG.spGrilla_Traslate rgrd, "FLG_FRACCIONAMIENTO", "0", "NO"
        
   rgrd.MarqueeStyle = dbgHighlightRow
   rgrd.HeadBackColor = &H8000000F '&H80000007
   rgrd.HeadForeColor = &H80000012 '&HFFFF&
'   rgrd.HeadBackColor = &H80000007
'   rgrd.HeadForeColor = &HFFFF&
   rgrd.RowHeight = 0
   rgrd.RowHeight = 420
   rgrd.HeadLines = 2
   rgrd.Font.Size = 8
   rgrd.Styles(5).Font.Size = 8
   rgrd.Columns(0).Visible = False
   rgrd.Columns(3).Visible = False
   
    rgrd.Columns("CTD_PRODUCTO").Style.BackColor = RGB(254, 244, 207)
    rgrd.Columns("CTD_PRODUCTO").Style.ForeColor = RGB(202, 24, 4)
    rgrd.Columns("CTD_PRODUCTO").Style.Font.Bold = True
    
    rgrd.Columns("CTD_PRODUCTO_FRAC").Style.BackColor = RGB(254, 244, 207)
    rgrd.Columns("CTD_PRODUCTO_FRAC").Style.ForeColor = RGB(202, 24, 4)
    rgrd.Columns("CTD_PRODUCTO_FRAC").Style.Font.Bold = True
End Sub

Private Sub grdCabecera_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim temp As String
    temp = grdCabecera.Columns(0).Text
    sub_sge_llena_grilla_detalle temp
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        sub_sge_llena_grilla_cabecera
    Case 3
        mdiPrincipal.picComandos.Enabled = True
        Unload Me
End Select
End Sub
'Procedimiento para redimensionar los controles y objetos del form
Sub redimensionaForm()
    If objSSGG.verifica_resolucion < 1792 Then
        Me.Width = 7250
        grdCabecera.Width = 7250
        grdDetalle.Width = 7200
    Else
        Me.Width = 10515
        grdCabecera.Width = 10400
        grdDetalle.Width = 10400
        
    End If
End Sub
