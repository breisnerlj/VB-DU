VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frm_VTA_RecetarioM 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAñadir 
      Caption         =   "&Añadir"
      Height          =   615
      Left            =   0
      Picture         =   "frm_VTA_RecetarioM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4935
      Width           =   1095
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modifica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin vbp_Ventas.ctlGrillaArray GrdInsumos 
      Height          =   1695
      Left            =   0
      TabIndex        =   41
      Top             =   3240
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2990
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlDataCombo ctlCboProveedor 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin TrueDBGrid70.TDBGrid GrdInsumos1 
      Height          =   855
      Left            =   0
      TabIndex        =   40
      Top             =   3240
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1508
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "#"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Cod Ins"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Tipo Insumo"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Codigo"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Descripción"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Medida"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Ctd Base"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Cantidad"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "% Margen"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Precio"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Sub Total"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Cod Unico"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Cod Und Med"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   13
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=13"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=132"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=53"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=2"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1191"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1111"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=1958"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1879"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=1296"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1217"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=4789"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=4710"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=0"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=1191"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=1111"
      Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=0"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=1455"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=1376"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(37)=   "Column(7).Width=1455"
      Splits(0)._ColumnProps(38)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(7)._WidthInPix=1376"
      Splits(0)._ColumnProps(40)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(42)=   "Column(8).Width=1323"
      Splits(0)._ColumnProps(43)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(8)._WidthInPix=1244"
      Splits(0)._ColumnProps(45)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(46)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(47)=   "Column(9).Width=1482"
      Splits(0)._ColumnProps(48)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(9)._WidthInPix=1402"
      Splits(0)._ColumnProps(50)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(51)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(52)=   "Column(10).Width=1561"
      Splits(0)._ColumnProps(53)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(10)._WidthInPix=1482"
      Splits(0)._ColumnProps(55)=   "Column(10)._ColStyle=2"
      Splits(0)._ColumnProps(56)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(57)=   "Column(11).Width=1588"
      Splits(0)._ColumnProps(58)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(11)._WidthInPix=1508"
      Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(61)=   "Column(12).Width=1931"
      Splits(0)._ColumnProps(62)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(12)._WidthInPix=1852"
      Splits(0)._ColumnProps(64)=   "Column(12).Order=13"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   0
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      RowDividerStyle =   0
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=12,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13,.alignment=1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=74,.parent=13,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=78,.parent=13,.alignment=0"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=75,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=76,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=77,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=0"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=46,.parent=13,.alignment=1"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=47,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=48,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=49,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=70,.parent=13,.alignment=1,.bgcolor=&HE5E5E5&"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=54,.parent=13,.alignment=1,.bgcolor=&HE5E5E5&"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=51,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=52,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=53,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
      _StyleDefs(88)  =   "Named:id=33:Normal"
      _StyleDefs(89)  =   ":id=33,.parent=0"
      _StyleDefs(90)  =   "Named:id=34:Heading"
      _StyleDefs(91)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(92)  =   ":id=34,.wraptext=-1"
      _StyleDefs(93)  =   "Named:id=35:Footing"
      _StyleDefs(94)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(95)  =   "Named:id=36:Selected"
      _StyleDefs(96)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(97)  =   "Named:id=37:Caption"
      _StyleDefs(98)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(99)  =   "Named:id=38:HighlightRow"
      _StyleDefs(100) =   ":id=38,.parent=33,.bgcolor=&HFAF0E7&,.fgcolor=&H0&"
      _StyleDefs(101) =   "Named:id=39:EvenRow"
      _StyleDefs(102) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(103) =   "Named:id=40:OddRow"
      _StyleDefs(104) =   ":id=40,.parent=33"
      _StyleDefs(105) =   "Named:id=41:RecordSelector"
      _StyleDefs(106) =   ":id=41,.parent=34"
      _StyleDefs(107) =   "Named:id=42:FilterBar"
      _StyleDefs(108) =   ":id=42,.parent=33"
   End
   Begin vbp_Ventas.ctlDataCombo ctlCboTipCliente 
      Height          =   315
      Left            =   5640
      TabIndex        =   5
      Top             =   2145
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlTextBox TxtObservacion 
      Height          =   855
      Left            =   2880
      TabIndex        =   12
      Top             =   5280
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1508
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
   Begin VB.CommandButton cmdBuscarCliente 
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   5280
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscarMedico 
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   5640
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscarProv 
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   5640
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   615
      Left            =   5745
      Picture         =   "frm_VTA_RecetarioM.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   4440
      Picture         =   "frm_VTA_RecetarioM.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6240
      Width           =   1095
   End
   Begin vbp_Ventas.ctlTextBox TxtProveedor 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      TipoSQL         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbp_Ventas.ctlTextBox TxtMedico 
      Height          =   315
      Left            =   480
      TabIndex        =   3
      Top             =   1500
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      ColorDefault    =   -2147483639
      ColorDefault    =   -2147483639
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbp_Ventas.ctlTextBox TxtCliente 
      Height          =   315
      Left            =   480
      TabIndex        =   6
      Top             =   2280
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      ColorDefault    =   -2147483639
      ColorDefault    =   -2147483639
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      Caption         =   "Tipo Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   10
      Left            =   4080
      TabIndex        =   39
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LblCliente 
      BackColor       =   &H00DBFBFA&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   38
      Top             =   2595
      Width           =   4935
   End
   Begin VB.Label LblMedico 
      BackColor       =   &H00DBFBFA&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   37
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Label LblNomprov 
      BackColor       =   &H00DBFBFA&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   36
      Top             =   1050
      Visible         =   0   'False
      Width           =   4935
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
      Index           =   11
      Left            =   6097
      TabIndex        =   35
      Top             =   6960
      Width           =   390
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
      Index           =   12
      Left            =   4380
      TabIndex        =   34
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   9
      Left            =   2880
      TabIndex        =   33
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones"
      Height          =   195
      Index           =   7
      Left            =   3120
      TabIndex        =   32
      Top             =   5040
      Width           =   1065
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Insumos"
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   31
      Top             =   3000
      Width           =   585
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F4"
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
      Index           =   8
      Left            =   120
      TabIndex        =   30
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   29
      Top             =   2085
      Width           =   525
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F3"
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
      Left            =   120
      TabIndex        =   28
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Médico"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   27
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F2"
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
      Left            =   120
      TabIndex        =   26
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor:"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   25
      Top             =   480
      Width           =   780
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F1"
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
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Imp. sin redondear :"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   6360
      Width           =   1395
   End
   Begin VB.Label lblSubTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0.00"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2220
      TabIndex        =   22
      Top             =   6360
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Redondeo :"
      Height          =   195
      Left            =   795
      TabIndex        =   21
      Top             =   6600
      Width           =   840
   End
   Begin VB.Label lblRedondeo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0.00"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2220
      TabIndex        =   20
      Top             =   6600
      Width           =   315
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descuento :"
      Height          =   195
      Left            =   645
      TabIndex        =   19
      Top             =   6000
      Width           =   870
   End
   Begin VB.Label lblDecuento 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0.00"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   2220
      TabIndex        =   18
      Top             =   6000
      Width           =   315
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Imp. a cobrar S/. :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   6960
      Width           =   1590
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2145
      TabIndex        =   16
      Top             =   6960
      Width           =   390
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   2640
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   2640
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label lblValor 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   2145
      TabIndex        =   15
      Top             =   5760
      Width           =   390
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Sub Total :"
      Height          =   195
      Left            =   735
      TabIndex        =   14
      Top             =   5760
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cotización Recetario Magistral"
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
      TabIndex        =   13
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "frm_VTA_RecetarioM.frx":109E
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "frm_VTA_RecetarioM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objCliente As New clsCliente
Dim objProducto As New clsProducto
Dim objProveedor As New clsProveedor

Dim objMedico As New clsMedico
Public pxdbInsumos As New XArrayDB

Public pstrDatoProv As String
Public pstrDatoMedico As String
Public pstrFlgCli As String
Public pstrDatoCliente As String

Public pstrIdProductoBtl As String
Public pstrIdProducto As String
Public pstrIdCantidad As String
Public pstrIdflgFracc As String
Public pstrIdPctDsc As String
Public pstrIdPreVta As String
Public pstrIdTipoVta As String

Public pstrFlgRM As String

Dim IntRow As Integer
Dim dblValor#, dblDcto#, dblSTot#, dblRedondeo#, dblTotal#, dblCant#
Dim bolCancela As Boolean
Dim rsMedico As oraDynaset
Dim rsCliente As oraDynaset

Public pstrRucProv As String
Private dblPrecioF As Double
Public CantBase As String


Private Sub Command1_Click()

End Sub

Private Sub cmdAñadir_Click()
    
    frm_VTA_RecetarioMagistral.strCodInsumo = ""
    frm_VTA_RecetarioMagistral.strConcentracion = ""
    frm_VTA_RecetarioMagistral.strCantidad = ""
    
    frm_VTA_RecetarioMagistral.Show

    
End Sub

Private Sub cmdModificar_Click()
    frm_VTA_RecetarioMagistral.strCodInsumo = GrdInsumos.Columns(1).Value
    frm_VTA_RecetarioMagistral.strConcentracion = GrdInsumos.Columns(8).Value
    frm_VTA_RecetarioMagistral.strCantidad = GrdInsumos.Columns(7).Value
    frm_VTA_RecetarioMagistral.txtBuscar.Text = GrdInsumos.Columns(3).Value
    frm_VTA_RecetarioMagistral.Buscar
    frm_VTA_RecetarioMagistral.CargarValores
    frm_VTA_RecetarioMagistral.Show
End Sub

Private Sub ctlCboProveedor_Click(Area As Integer)
'    pstrDatoProv = Trim(txtProveedor.Text)
'    pstrRucProv = Trim(txtProveedor.Text)
    
    pstrDatoProv = ctlCboProveedor.BoundText
    pstrRucProv = ctlCboProveedor.BoundText
    frm_VTA_RecetarioMagistral.pstrRucProv = ctlCboProveedor.BoundText
    
    
End Sub

Private Sub Form_Activate()
On Error GoTo handle
    'TxtProveedor.SetFocus
    'gstrFlgRM = "1"
    'pstrFlgRM = "1"
    'penumVentCli = Recetario_Magistral
    'objVenta.CodigoTipoVenta = Recetario

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo handle
    psub_KeyDownAplicacion KeyCode, Shift
    
    Select Case KeyCode
        Case vbKeyF1
            TxtProveedor.SetFocus
        Case vbKeyF2
            TxtMedico.SetFocus
        Case vbKeyF3
            TxtCliente.SetFocus
        Case vbKeyF4
            GrdInsumos.SetFocus
        Case vbKeyF9
            TxtObservacion.SetFocus
        Case vbKeyInsert
            cmdAñadir_Click
        Case vbKeyEscape
                cmdCancelar_Click
        Case vbKeyReturn
                If Shift = 1 Then cmdGrabar_Click
            
            
    End Select

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName


End Sub

Private Sub Form_Load()

    'SetteaFormulario Me
    Call sub_BegArray(pxdbInsumos)
    setteaFormulario Me
    'Me.top = 0
    'Me.left = 0
    
'    GrdInsumos.RecordSelectors = False
    On Error GoTo CtrlErr
    
'    Set ctlCboTipCliente.RowSource = ObjCliente.Lista_Tipo_Doc_Indet
'    ctlCboTipCliente.ListField = "DES"
'    ctlCboTipCliente.BoundColumn = "COD"
'    ctlCboTipCliente.BoundText = "*"
    
    Set ctlCboTipCliente.RowSource = objCliente.ListaTipo
    ctlCboTipCliente.ListField = "DES"
    ctlCboTipCliente.BoundColumn = "COD"
    ctlCboTipCliente.BoundText = "*"
    
    peTransacc = GbProforma
    
    '*** Seteo de la Grilla ***'
    
    SetGrd
    Dim i As Integer
    For i = 0 To GrdInsumos.Columns.Count - 1
        GrdInsumos.Columns(i).AllowFocus = False
    Next i
''    GrdInsumos.MarqueeStyle = dbgHighlightCell
''    GrdInsumos.Columns(6).AllowFocus = True
''    GrdInsumos.Columns(7).AllowFocus = True
''    GrdInsumos.Columns(8).AllowFocus = True
    
    GrdInsumos.Columns(0).Visible = False
    GrdInsumos.Columns(1).Visible = False
    GrdInsumos.Columns(3).Visible = False
    GrdInsumos.Columns(11).Visible = False
    GrdInsumos.Columns(12).Visible = False
    
    
    
''    GrdInsumos.Columns(6).NumberFormat = "#0"
''    GrdInsumos.Columns(7).NumberFormat = "#0." & String(gintDecTot, "0")
''    GrdInsumos.Columns(8).NumberFormat = "#0." & String(gintDec, "0")
''    GrdInsumos.Columns(9).NumberFormat = "#0." & String(gintDecTot, "0")
''    GrdInsumos.Columns(10).NumberFormat = "#0." & String(gintDecTot, "0")
    
    
    
    '''''Set ctlCboProveedor.RowSource = objProveedor.ListaRegMagistral(pstrDatoProv, pstrFlgRM, objUsuario.CodigoLocal)
    Set ctlCboProveedor.RowSource = objProveedor.ListaRegMagistral(pstrDatoProv, "1", objUsuario.CodigoLocal)
    ctlCboProveedor.ListField = "NOM_PROVEEDOR"
    ctlCboProveedor.BoundColumn = "RUC_PROVEEDOR"
    ctlCboProveedor.BoundText = objVenta.ProvPreDeterminadoRM
    
    CantBase = ""
    
    dblPrecioF = objUsuario.PrecMinRM
    
    Exit Sub
CtrlErr:

    MsgBox Err.Description, vbOKOnly + vbInformation, Err.Number
End Sub

Private Sub sub_BegArray(ByVal vxdb As XArrayDB)
On Error GoTo handle
    vxdb.ReDim 0, -1, 0, 13
    GrdInsumos.Array1 = vxdb
    GrdInsumos.Rebind
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Sub psub_Items(ByRef rxdb As XArrayDB, ByVal intCol%)
On Error GoTo handle
    Dim i%
    For i = "0" & GrdInsumos.Columns(0).Value To rxdb.UpperBound(1)
        rxdb(i, intCol) = i + 1
    Next i

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Public Sub psub_Agrega_Insumo(Optional vstrCodIns As String, Optional vstrDesIns As String, _
                              Optional vstrCod As String, Optional vstrDes As String, _
                              Optional vstrUnd As String, Optional vstrPctMargen As String, _
                              Optional vstrCtdBase As String, Optional vstrCant As String, _
                              Optional vstrPrecio As String, Optional vstrSubTot As String, _
                              Optional vstrCodProdBtl As String, Optional vstrCodUnd As String)
On Error GoTo handle
    Dim k As Integer
    On Error GoTo CntError
    
    '*** Valida que solo el compuesto tenga una base ******************************************************'
    
    
    IntRow = pxdbInsumos.Find(0, 3, vstrCod, XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
    
    If vstrCodIns = "001" Then
        Dim w%
        w = pxdbInsumos.Find(0, 1, vstrCodIns)
            If IntRow <> 0 Then
                If w = 0 Then
                    MsgBox "Ya existe una Base en los compuesto del recetario", vbExclamation, Caption: Exit Sub
                End If
            End If
    End If
    '*****************************************************************************************************'
    
    
    

    On Error GoTo 0
    GoTo Ok
CntError:
    IntRow = -1
    On Error GoTo 0
    GoTo Ok
Ok:

    If IntRow = -1 Then
            pxdbInsumos.AppendRows 1
            Call psub_Items(pxdbInsumos, 0)
            pxdbInsumos(pxdbInsumos.UpperBound(1), 1) = vstrCodIns
            pxdbInsumos(pxdbInsumos.UpperBound(1), 2) = vstrDesIns
            pxdbInsumos(pxdbInsumos.UpperBound(1), 3) = vstrCod
            pxdbInsumos(pxdbInsumos.UpperBound(1), 4) = vstrDes
            pxdbInsumos(pxdbInsumos.UpperBound(1), 5) = vstrUnd
            pxdbInsumos(pxdbInsumos.UpperBound(1), 6) = vstrCtdBase
            pxdbInsumos(pxdbInsumos.UpperBound(1), 7) = vstrCant
            pxdbInsumos(pxdbInsumos.UpperBound(1), 8) = vstrPctMargen
            pxdbInsumos(pxdbInsumos.UpperBound(1), 9) = vstrPrecio
            pxdbInsumos(pxdbInsumos.UpperBound(1), 10) = (vstrCant * vstrPrecio) 'vstrSubTot
            pxdbInsumos(pxdbInsumos.UpperBound(1), 11) = vstrCodProdBtl 'Codigo Unico Btl'
            pxdbInsumos(pxdbInsumos.UpperBound(1), 12) = vstrCodUnd 'Codigo Unidad capacidad'
    
      Else
            pxdbInsumos(IntRow, 1) = vstrCodIns
            pxdbInsumos(IntRow, 2) = vstrDesIns
            pxdbInsumos(IntRow, 3) = vstrCod
            pxdbInsumos(IntRow, 4) = vstrDes
            pxdbInsumos(IntRow, 5) = vstrUnd
            pxdbInsumos(IntRow, 6) = vstrCtdBase
            pxdbInsumos(IntRow, 7) = vstrCant
            pxdbInsumos(IntRow, 8) = vstrPctMargen
            pxdbInsumos(IntRow, 9) = vstrPrecio
            pxdbInsumos(IntRow, 10) = (vstrCant * vstrPrecio)
            pxdbInsumos(IntRow, 11) = vstrCodProdBtl 'Codigo Unico Btl'
            pxdbInsumos(IntRow, 12) = vstrCodUnd 'Codigo Unidad capacidad'
      
      
      
    End If
    'GrdInsumos.Rebind

'    dblPreTot = 0
'    For k = 0 To pxdbInsumos.UpperBound(1)
'        dblPreTot = dblPreTot + Val(pxdbInsumos(k, 10))
'    Next k

    '** calcular los montos **'
    Call psub_Calcula_Precio

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub cmdBuscarProv_Click()
On Error GoTo handle
    pstrDatoProv = Trim(TxtProveedor.Text)
    frm_VTA_ProveedorDatos.Show vbModal
    pstrRucProv = Trim(TxtProveedor.Text)

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub cmdBuscarMedico_Click()
On Error GoTo handle
    pstrDatoMedico = Trim(TxtMedico.Text)
    frm_VTA_MedicoDatos.Show vbModal

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub cmdBuscarCliente_Click()
On Error GoTo handle
    If ctlCboTipCliente.BoundText = "*" Then MsgBox "Seleccione un tipo de Cliente", vbCritical, Caption: Exit Sub
    pstrFlgCli = ctlCboTipCliente.BoundText
    pstrDatoCliente = Trim(TxtCliente.Text)
    Set frm_VTA_ClienteDatos.GrdBusCliente.DataSource = objCliente.ListaClientesGen(pstrDatoCliente, pstrFlgCli)
    frm_VTA_ClienteDatos.Pantalla = 2
    frm_VTA_ClienteDatos.Show

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    If bolCancela Then
'        If MsgBox("Salir del formulario", vbYesNo + vbQuestion, App.ProductName) = vbYes Then
'            Cancel = 0
'        Else
'            Cancel = 1
'        End If
'    End If
End Sub

'/*Private Sub GrdInsumos_AfterColUpdate(ByVal ColIndex As Integer)
'    pxdbInsumos(GrdInsumos.Bookmark, ColIndex) = GrdInsumos.Columns(ColIndex).Value
'    Select Case ColIndex
'        Case 6
'            pxdbInsumos(GrdInsumos.Bookmark, 6) = (pxdbInsumos(GrdInsumos.Bookmark, 6))
'        Case 7
'            On Error GoTo CtrlErr1
'            ' Formula para sacar el "Pct"  =>  (Cant / Base )*100 '
'            pxdbInsumos(GrdInsumos.Bookmark, 8) = ((pxdbInsumos(GrdInsumos.Bookmark, 7)) / (pxdbInsumos(GrdInsumos.Bookmark, 6))) * 100
'            pxdbInsumos(GrdInsumos.Bookmark, 10) = (pxdbInsumos(GrdInsumos.Bookmark, 7)) * (pxdbInsumos(GrdInsumos.Bookmark, 9))
'            psub_Calcula_Precio
'            GrdInsumos.Col = 7
'CtrlErr1:
'            On Error GoTo 0
'
'        Case 8
'            On Error GoTo CtrlErr2
'            ' Formula para sacar la "Cant"  =>  (Pct / 100 )* Ctd Base '
'            pxdbInsumos(GrdInsumos.Bookmark, 7) = (pxdbInsumos(GrdInsumos.Bookmark, 8) / 100) * (pxdbInsumos(GrdInsumos.Bookmark, 6))
'            pxdbInsumos(GrdInsumos.Bookmark, 10) = (pxdbInsumos(GrdInsumos.Bookmark, 7)) * (pxdbInsumos(GrdInsumos.Bookmark, 9))
'            psub_Calcula_Precio
'            GrdInsumos.Col = 8
'CtrlErr2:
'            On Error GoTo 0
'    End Select
'    GrdInsumos.Rebind
'End Sub*/

'/*Private Sub GrdInsumos_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'    pxdbInsumos(GrdInsumos.Bookmark, ColIndex) = GrdInsumos.Columns(ColIndex).Value
'    Select Case ColIndex
'        Case 6, 7, 8
'            If Not IsNumeric(GrdInsumos.Text) And GrdInsumos.Text <> "" Then
'                    MsgBox "El valor ingreso no es válido", vbExclamation, "Error"
'                    pxdbInsumos(GrdInsumos.Bookmark, ColIndex) = OldValue
'                    Cancel = True
'              ElseIf Trim(GrdInsumos.Text) <> "" Then
'                    'GrdInsumos.Columns(ColIndex).Value = Format(Trim(GrdInsumos.Columns(ColIndex).Value), "0." & String(gintDecTot, "0"))
'                    GrdInsumos.Columns(ColIndex).Value = Trim(GrdInsumos.Columns(ColIndex).Value)
'            End If
'    End Select
'
'End Sub*/

Private Sub GrdInsumos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If pxdbInsumos.UpperBound(1) <> "-1" Then
        If pxdbInsumos(GrdInsumos.Bookmark, 1) = "001" Then
            GrdInsumos.Columns(6).AllowFocus = False
            GrdInsumos.Columns(7).AllowFocus = False
            GrdInsumos.Columns(8).AllowFocus = False
          Else
            GrdInsumos.Columns(6).AllowFocus = True
            GrdInsumos.Columns(7).AllowFocus = True
            GrdInsumos.Columns(8).AllowFocus = True
        End If
    End If
End Sub



Private Sub GrdInsumos_Click()
    GrdInsumos.SetFocus
End Sub

Private Sub GrdInsumos_DblClick()
    cmdModificar_Click
End Sub



Private Sub GrdInsumos_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo handle
    Select Case KeyCode
        Case vbKeyReturn
            cmdModificar_Click
        Case vbKeyDelete
             If GrdInsumos.ApproxCount > 0 Then
               If MsgBox("Esta seguro de Eliminar el Insumo .. ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                    GrdInsumos.SetFocus
                    GrdInsumos.Delete
                    GrdInsumos.Rebind
                    '''frm_VTA_CantidadProducto.strCantAnt = ""
                    psub_Calcula_Precio
                End If
             End If
             
             If GrdInsumos.ApproxCount = 0 Then CantBase = ""
             
    End Select

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub TxtCliente_Change()
    On Error GoTo handle
    If Len(TxtCliente.Text) <= 0 Then LblCliente.Caption = ""

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub TxtCliente_KeyPress(KeyAscii As Integer)
On Error GoTo handle
    If KeyAscii = 13 Then
            Set rsCliente = objCliente.ListaBuscador(TxtCliente.Text)
            If Not rsCliente.EOF Then
                LblCliente.Caption = rsCliente("NOMBRE").Value
            Else
                If Trim(TxtCliente.Text) <> "" Then
                    If MsgBox("El Cliente NO EXISTE, desea agregarlo ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                        frm_VTA_Cliente.ctlCliente1.Cargar
                        frm_VTA_Cliente.ctlCliente1.Codigo = ""
                        frm_VTA_Cliente.CargarValores
                        frm_VTA_Cliente.Show vbModal
                    End If
                Else
                    LblCliente.Caption = ""
                    'SendKeys "{TAB}"
                End If
            End If
    End If
Exit Sub
handle:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
End Sub

Private Sub TxtMedico_Change()
On Error GoTo handle
    If Len(TxtMedico.Text) <= 0 Then LblMedico.Caption = ""

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub TxtMedico_KeyPress(KeyAscii As Integer)
On Error GoTo handle
    If KeyAscii = 13 Then
            Set rsMedico = objMedico.ListaCMP(TxtMedico.Text)
            If Not rsMedico.EOF Then
                LblMedico.Caption = rsMedico("NOM_MEDICO").Value
            Else
                If Trim(TxtMedico.Text) <> "" Then
                    If MsgBox("El Médico NO EXISTE, desea agregarlo ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                        Call frmGrabaMedico.Datos("", "", "", "", "", "", "0", "", "-(Nuevo)")
                    End If
                Else
                    LblMedico.Caption = ""
                    'SendKeys "{TAB}"
                End If
            End If

    End If
Exit Sub

handle:
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
    
End Sub

Private Sub TxtProveedor_Change()
On Error GoTo handle
    If Len(TxtProveedor.Text) <= 0 Then LblNomprov.Caption = ""

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

'*********************************** *******************************'
'**** Carga el detalle de Insumos en variables de Memoria****'
Public Sub psub_Load_DetInsumos()
On Error GoTo handle
    Dim i As Integer
    pstrIdProducto = "": pstrIdPctDsc = ""
    pstrIdCantidad = "": pstrIdPreVta = ""
    pstrIdProductoBtl = ""
    
'    For i = 0 To pxdbInsumos.UpperBound(1)
'        pstrIdProducto = pstrIdProducto & IIf(IsNull(Trim(pxdbInsumos(i, 3))) Or Trim(pxdbInsumos(i, 3)) = "", "", Trim(pxdbInsumos(i, 3))) & "|"
'        pstrIdPctDsc = pstrIdPctDsc & IIf(IsNull(Trim(pxdbInsumos(i, 8))) Or Trim(pxdbInsumos(i, 8)) = "", "0", Trim(pxdbInsumos(i, 8))) & "|"
'        pstrIdCantidad = pstrIdCantidad & IIf(IsNull(Trim(pxdbInsumos(i, 7))) Or Trim(pxdbInsumos(i, 7)) = "", "0", Trim(pxdbInsumos(i, 7))) & "|"
'        pstrIdflgFracc = pstrIdflgFracc & "0" & "|"
'        pstrIdPreVta = pstrIdPreVta & IIf(IsNull(Trim(pxdbInsumos(i, 10))) Or Trim(pxdbInsumos(i, 10)) = "", "0", Trim(pxdbInsumos(i, 10))) & "|"
'        pstrIdProductoBtl = pstrIdProductoBtl & IIf(IsNull(Trim(pxdbInsumos(i, 11))) Or Trim(pxdbInsumos(i, 11)) = "", "0", Trim(pxdbInsumos(i, 11))) & "|"  'Codigo Unico RM'
'        pstrIdTipoVta = pstrIdTipoVta & objVenta.CodigoTipoVenta & "|"
'    Next i
     
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub
'*********************************** *******************************'

Private Sub psub_Calcula_Precio()
On Error GoTo handle
    Dim dblPreTot As Double
    Dim k%
    
    dblPreTot = 0
    For k = 0 To pxdbInsumos.UpperBound(1)
        dblPreTot = dblPreTot + Val(pxdbInsumos(k, 10))
    Next k

   dblValor = dblPreTot: lblValor.Caption = Format(dblValor, "#0.00")
   dblDcto = 0: lblDecuento.Caption = Format(dblDcto, "#0.00")
   dblSTot = dblValor + dblDcto: lblSubTotal.Caption = Format(dblSTot, "#0.00")
   dblRedondeo = 0: lblRedondeo.Caption = Format(dblRedondeo, "#0.00")
   dblTotal = dblSTot + dblRedondeo: lblTotal.Caption = Format(dblTotal, "#0.00")
   
    If dblTotal < dblPrecioF Then
        '*** Cuando el monto es menor a S/ 7.00
        '*** q' es el monto minimo => Sera S/ 7.00
        dblTotal = dblPrecioF
        lblTotal.Caption = Format(dblTotal, "#0.00")
    End If
    
   Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
   
End Sub

Sub Calculando_Precion_rm()
    
   
    
End Sub


Private Sub cmdGrabar_Click()
On Error GoTo handle
    'If TxtProveedor.Text = "" Then MsgBox "Ingresar el Proveedor", vbCritical, Caption: Exit Sub
    
    '** Variables cargadas en memoria para pasar al SP PROFORMAS **'
    pstrDatoProv = Trim(TxtProveedor.Text)
    pstrDatoCliente = Trim(TxtCliente.Text)
    pstrDatoMedico = Trim(TxtMedico.Text)
    
    '-- Carga Arreglo de Insumos --'
    psub_Load_DetInsumos
    
    '-- Pasando el Codigo Unico RM --'
    Dim strCod As String
    Dim strDes As String
    
    ' Valida antes que el arreglo producto no tenga registro ya que en el magistral
    ' solo se puede hacer una modalidad y no poder tener mas de 1 --- 23/08/2007 Por Crueda
    
    'If objVenta.Producto.UpperBound(1) <> -1 Then Exit Sub
    '************************************************************'
    
    If pxdbInsumos.UpperBound(1) = -1 Then Exit Sub
    strCod = objProducto.ListaDevRM(objUsuario.CodigoLocal, pstrRucProv)
    If Val(strCod) = 0 Then MsgBox "Revise el Nº Ruc del Proveedor", vbCritical, "Aviso": TxtProveedor.SetFocus: Exit Sub
    strDes = objProducto.ListaDescripcion(strCod)
    
    '****************************************************************************************************'
    '****************************************************************************************************'
    '**  Permite hacer el calculo del redondeo a favor o en contra segun los decimales del recetario ****'
    '**                                 CAMBIOS HECHOS EL 15/03/2007                                   **'
    '****************************************************************************************************'
    
    
    
    If dblTotal < Format(dblPrecioF, "#0.00") Then
        '*** Cuando el monto es menor a S/ 7.00
        '*** q' es el monto minimo => Sera S/ 7.00
        dblTotal = Format(dblPrecioF, "#0.00")
    ElseIf dblTotal >= Format(dblPrecioF, "#0.00") Then
        Dim dblDec As Double
        dblDec = right(Format(dblTotal, "#0.00"), 2)
        '*** Evaluando los decimales     ***'
        '*** para los calculos del monto ***'
        If dblDec < 10 Then
            '*** Cuando el monto es de S/ 8.01 => Sera S/ 8.00 ***'
            Dim p%
            Dim Cadena1$, cadena2$
            For p = 1 To Len(dblTotal)
                Cadena1 = Mid(dblTotal, p, 1)
                If Cadena1 <> "." Then
                    cadena2 = cadena2 & Cadena1
                  Else
                    GoTo Union1
                End If
            Next p
Union1:
            dblTotal = cadena2
            
        Else
           '*** Cuando el monto es de S/ 8.1 => Sera S/ 9.00 ***'
           Dim w%
           Dim Cadena3$, Cadena4$
           For w = 1 To Len(dblTotal)
                Cadena3 = Mid(dblTotal, w, 1)
                If Cadena3 <> "." Then
                    Cadena4 = Cadena4 & Cadena3
                  Else
                    GoTo Union2
                End If
            Next w
Union2:
            dblTotal = Val(Cadena4) + 1
            
        End If
    End If
   
   '****************************************************************************************************'
   '****************************************************************************************************'
   '****************************************************************************************************'
    Dim strCant%
    Dim Indicador As String
    Dim PctComi  As Double
    '** Cambio para pasarle como cantidad lo mismo que le salio la cotizacion **'
    Indicador = objProducto.CodIndicadorReceta(strCod)
    PctComi = objProducto.pctComision(strCod, objUsuario.CodigoLocal, Format(ptmTipoPrecio, "000"))
    
    strCant = CStr(dblTotal)
    objVenta.AgregaProducto strCod, strDes, strCant, "0", dblTotal, objVenta.CodigoTipoVenta, Producto_Normal, , , , , , Indicador, PctComi
                                  
    
    Dim k%
    k = 0
    For k = 0 To pxdbInsumos.UpperBound(1)
        objVenta.AgregaRecetarioM strCod, _
                                  pxdbInsumos(k, 3), _
                                  pxdbInsumos(k, 7), _
                                  "0", _
                                  pxdbInsumos(k, 9), _
                                  objVenta.CodigoTipoVenta, _
                                  pxdbInsumos(k, 6), _
                                  pxdbInsumos(k, 8), _
                                  pxdbInsumos(k, 12)
    Next k
    'Dim k%
    'k = 1
    'For k = 1 To GrdInsumos.ApproxCount
'        objVenta.AgregaRecetarioM pstrIdProductoBtl, _
'                                  pstrIdProducto, _
'                                  pstrIdCantidad, _
'                                  pstrIdflgFracc, _
'                                  pstrIdPreVta, _
'                                  pstrIdTipoVta
    'Next k
                              
    
    
    frmPedido.Cal_Promo
    frmPedido.Cal_Montos
    frmPedido.grdPedido.Rebind
    
    'Unload Me
    
    Me.Hide

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub cmdCancelar_Click()
'    pstrDatoProv = "": pstrDatoCliente = ""
'    pstrDatoMedico = ""
'    pstrIdProducto = "": pstrIdProducto = ""
'    pstrIdCantidad = "": pstrIdPctDsc = "":
'    pstrIdPreVta = ""
On Error GoTo handle
    bolCancela = True
''''    Unload Me
''''
''''
''''    frm_VTA_Busqueda.grdProductos.Limpiar
''''    frm_VTA_Busqueda.grdAlternativos.Limpiar
''''    frm_VTA_Busqueda.grdComplementarios.Limpiar
''''    frm_VTA_Busqueda.txtBuscar.selection
    
    

''   If MsgBox("¿ Desea borrar todos los Datos del Documento.. ?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
''        frmPedido.psub_BeginArry
''        frm_VTA_Documento.blnTipoDoc = False
''        mdiPrincipal.subNuevo
''        frm_VTA_Modalidad.Show vbModal
        
        
        Me.Hide
        
        
''   End If
Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
   
    
    
    'objVenta.CancelarVenta
End Sub


Private Sub SetGrd()
Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant

    arrCampos = Array("", "", "", "", "", "", "", "", "", "", "", "", "")
    arrCaption = Array("#", "Cod Ins", "Tipo Insumo", "Código", "Descripción", "Medida", "Base", "Cant.", "% Margen", "Precio", "Sub Total", "Cod Unico", "Cod Und Med")
    arrAncho = Array(0, 0, 1000, 0, 2100, 500, 500, 500, 500, 550, 950, 0, 0)
    arrAlineacion = Array(dbgGeneral, dbgGeneral, dbgLeft, dbgGeneral, dbgLeft, dbgLeft, dbgRight, dbgRight, dbgRight, dbgRight, dbgRight, dbgGeneral, dbgGeneral)
    GrdInsumos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion

End Sub



