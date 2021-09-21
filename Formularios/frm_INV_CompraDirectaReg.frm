VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_INV_CompraDirectaReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva Recepción / Registro"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   11730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   1200
      Picture         =   "frm_INV_CompraDirectaReg.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   615
      Left            =   0
      Picture         =   "frm_INV_CompraDirectaReg.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   4995
      Left            =   0
      TabIndex        =   10
      Top             =   1800
      Width           =   11745
      Begin VB.CommandButton cmdBuscarPrd 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   6960
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
      Begin vbp_Ventas.ctlTextBox txtProducto 
         Height          =   375
         Left            =   1320
         TabIndex        =   33
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   661
         MaxLength       =   30
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
      Begin TrueDBGrid70.TDBGrid grdDetalle 
         Height          =   3495
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   6165
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
         Columns(1).Caption=   "Código"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descripción"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Cant. Factura"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Cant. Física"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Precio"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).DataField=   ""
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).DataField=   ""
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
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
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2170"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2090"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=1"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=9446"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=9366"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=2752"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2672"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=1"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=2778"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2699"
         Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=0"
         Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(27)=   "Column(5).Width=2593"
         Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2514"
         Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=0"
         Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(32)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(39)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(40)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(41)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(43)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(44)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(45)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(46)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(47)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(48)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(49)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(50)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(51)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(52)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(53)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(54)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(55)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(56)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(57)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(58)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(59)=   "Column(12).Order=13"
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
         DirectionAfterEnter=   0
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
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=50,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=54,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=51,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=52,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=53,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=82,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=79,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=80,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=81,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=70,.parent=13"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=67,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=68,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=69,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=86,.parent=13"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=83,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=84,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=85,.parent=17"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=66,.parent=13"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=63,.parent=14"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=64,.parent=15"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=65,.parent=17"
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
      Begin vbp_Ventas.ctlTextBox txtRedondeo 
         Height          =   375
         Left            =   4920
         TabIndex        =   40
         Top             =   4320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Tipo            =   4
         Alignment       =   1
         MaxLength       =   11
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
      Begin vbp_Ventas.ctlTextBox txtSubtotal 
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   4320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Tipo            =   4
         Alignment       =   1
         MaxLength       =   11
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
      Begin vbp_Ventas.ctlTextBox txtImpuesto 
         Height          =   375
         Left            =   1320
         TabIndex        =   37
         Top             =   4320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Tipo            =   4
         Alignment       =   1
         MaxLength       =   11
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
      Begin vbp_Ventas.ctlTextBox txtTotal 
         Height          =   375
         Left            =   3720
         TabIndex        =   39
         Top             =   4320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Tipo            =   4
         Alignment       =   1
         MaxLength       =   11
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
      Begin vbp_Ventas.ctlTextBox txtInafecto 
         Height          =   375
         Left            =   2520
         TabIndex        =   38
         Top             =   4320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Tipo            =   4
         Alignment       =   1
         MaxLength       =   11
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
      Begin VB.Label lblMRedondeo 
         Alignment       =   2  'Center
         Caption         =   "Redondeo"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5040
         TabIndex        =   25
         Top             =   4680
         Width           =   1050
      End
      Begin VB.Label lblMInafecto 
         Alignment       =   2  'Center
         Caption         =   "Inafecto"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2520
         TabIndex        =   23
         Top             =   4680
         Width           =   1170
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Inafecto"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   9240
         TabIndex        =   22
         Top             =   4680
         Width           =   1200
      End
      Begin VB.Label lblInafecto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   9240
         TabIndex        =   21
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Producto :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   330
         Width           =   735
      End
      Begin VB.Label lblMSubtotal 
         Alignment       =   2  'Center
         Caption         =   "Afecto"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   4680
         Width           =   1170
      End
      Begin VB.Label lblMImpuesto 
         Alignment       =   2  'Center
         Caption         =   "Impuesto"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1320
         TabIndex        =   18
         Top             =   4680
         Width           =   1170
      End
      Begin VB.Label lblMTotal 
         Alignment       =   2  'Center
         Caption         =   "Total"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3720
         TabIndex        =   17
         Top             =   4680
         Width           =   1170
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Afecto"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6840
         TabIndex        =   16
         Top             =   4680
         Width           =   1170
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Impuesto"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   8160
         TabIndex        =   15
         Top             =   4680
         Width           =   1020
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   10440
         TabIndex        =   14
         Top             =   4680
         Width           =   1170
      End
      Begin VB.Label lblSubtotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   6840
         TabIndex        =   13
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label lblImpuesto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   8040
         TabIndex        =   12
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   10440
         TabIndex        =   11
         Top             =   4320
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11745
      Begin VB.Frame Frame6 
         Caption         =   "&Recepción"
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
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   4935
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "&Buscar"
            Height          =   615
            Left            =   3600
            Picture         =   "frm_INV_CompraDirectaReg.frx":0B14
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   160
            Width           =   1095
         End
         Begin vbp_Ventas.ctlTextBox txtOCompra 
            Height          =   375
            Left            =   1680
            TabIndex        =   26
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            MaxLength       =   11
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
            Caption         =   "Orden de Compra"
            ForeColor       =   &H00004080&
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   330
            Width           =   1245
         End
      End
      Begin vbp_Ventas.ctlTextBox txtProveedor 
         Height          =   315
         Left            =   5400
         TabIndex        =   29
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         MaxLength       =   11
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
      Begin vbp_Ventas.ctlTextBox txtDocumento 
         Height          =   315
         Left            =   3720
         TabIndex        =   31
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Tipo            =   3
         MaxLength       =   11
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
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   6480
         TabIndex        =   32
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16449539
         CurrentDate     =   38874
      End
      Begin vbp_Ventas.ctlDataCombo ctlcboMotivo 
         Height          =   315
         Left            =   1320
         TabIndex        =   28
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlDataCombo ctlcboDocumento 
         Height          =   315
         Left            =   1320
         TabIndex        =   30
         Top             =   1440
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label lblObservaciones 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<Sin Observaciones> "
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   5400
         TabIndex        =   43
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Motivo :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1140
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4320
         TabIndex        =   8
         Top             =   1140
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Documento :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1500
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5760
         TabIndex        =   6
         Top             =   1500
         Width           =   540
      End
      Begin VB.Label lblProveedor 
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
         Left            =   6960
         TabIndex        =   5
         Top             =   1125
         Width           =   75
      End
      Begin VB.Label lblForma 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   8160
         TabIndex        =   2
         Top             =   1440
         Width           =   1170
      End
      Begin VB.Label lblFormaPago 
         AutoSize        =   -1  'True
         Caption         =   "CONTADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   9360
         TabIndex        =   1
         Top             =   1440
         Width           =   915
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Eliminar Item : CTRL+DEL"
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   9720
      TabIndex        =   24
      Top             =   7005
      Width           =   1845
   End
End
Attribute VB_Name = "frm_INV_CompraDirectaReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Definicion de propiedades
Private objProducto As New clsProducto
Private objCompra As New clsCompra
Private objProveedor As New clsProveedor
Private istrCaso As String
Private ixdbDetalle As New XArrayDB
Private iintColumnas As Integer
Private iintColDesProducto As Integer
Private iintColCodProducto As Integer
Private iintColQFactura As Integer
Private iintColQFisica As Integer
Private iintColPrecio As Integer
Private iintColQOFactura As Integer
Private iintColQOFisica As Integer
Private iintColOPrecio As Integer
Private iintColProductoProv As Integer
Private iintColLote As Integer
Private iintColVcmto As Integer
Private iintColObs As Integer
Private prstProducto As oraDynaset
Private istrMensajePostGraba As String
Private istrGuiaGenerada As String
Private istrRegCompra As String
Private ibooConsulta As Boolean
Private istrProveedor As String
Private istrTipDocumento As String
Private istrNumDocumento As String

Private Sub cmdBuscar_Click()

On Error GoTo Control

    Dim lstrOCompra As String
    Dim lrsOCompra As oraDynaset
    
    'Validamos que haya ingresado una OC
    lstrOCompra = Trim(txtOCompra.Text)
    
    'Poblar los textos y grilla con los datos obtenidos
    ctlcboMotivo.BoundText = "*"
    lblFormaPago.Caption = ""
    txtProveedor.Text = ""
    lblProveedor.Caption = ""
    ctlcboDocumento.BoundText = "*"
    txtDocumento.Text = ""
    dtpFecha.Value = CDate(Format(Now, "dd/mm/yyyy"))
    SetArray
    
    'Invocar package que trae los datos de la OC
    If lstrOCompra <> "" Then
        CargaDatos lstrOCompra
    End If
    
    'Si no encuentra la OC entonces mostramos ayuda
    If grdDetalle.ApproxCount < 1 Then
        
        Dim objCompra As New clsCompra
        Set lrsOCompra = objCompra.OrdenCompraLista(objUsuario.CodigoLocal)
        Set objCompra = Nothing
        
        If Not lrsOCompra.EOF Then
    
            Dim lfrmAyuda As New frm_INV_ProductoDatos
            Dim larrCabecera As Variant
    
            ReDim larrCabecera(0 To 9)
            larrCabecera(0) = "O.Compra"
            larrCabecera(1) = "Cód.Proveedor"
            larrCabecera(2) = "Proveedor"
            larrCabecera(3) = "Estado"
            larrCabecera(4) = "Observaciones"
            larrCabecera(5) = "Local"
            larrCabecera(6) = "Cód.Tipo"
            larrCabecera(7) = "Tipo"
            larrCabecera(8) = "Cód.Pago"
            larrCabecera(9) = "Forma de Pago"
            
            Screen.MousePointer = vbHourglass
            Set lfrmAyuda.irsProducto = lrsOCompra
            lfrmAyuda.iarrCabecera = larrCabecera
            lfrmAyuda.Caption = "Ordenes de compra pendientes"
            lfrmAyuda.Show vbModal
            lstrOCompra = Trim(CStr(lfrmAyuda.istrCodProducto))
            txtOCompra.Text = lstrOCompra
            Screen.MousePointer = vbDefault
            Set lfrmAyuda = Nothing
            
            If lstrOCompra <> "" Then
                CargaDatos lstrOCompra
            End If
            
        End If
        
    End If
    
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
    
End Sub

Private Sub cmdBuscarPrd_Click()

On Error GoTo Control

    Dim lstrProducto As String
    Dim lintRow As Integer
    Dim lbooEncontrado As Boolean
    
    'Ubicar en la grilla
    lstrProducto = Trim(txtProducto.Text)
    lbooEncontrado = False
    
    If lstrProducto = "" Then
        MsgBox "Debe indicar el Producto", vbExclamation, "Aviso"
        txtProducto.SetFocus
        Exit Sub
    End If
    
    'Buscamos por Codigo
    If IsNumeric(lstrProducto) Then
        'Caso codigo de barra
        If Val(lstrProducto) > 99999 Then
            Dim lobjProducto As New clsProducto
            lstrProducto = lobjProducto.CodBarraProducto(lstrProducto)
            Set lobjProducto = Nothing
        End If
        
        If Not lbooEncontrado Then
            If ixdbDetalle.UpperBound(1) > 0 Then
                lintRow = ixdbDetalle.Find(ixdbDetalle.LowerBound(1), grdDetalle.Columns(iintColCodProducto).ColIndex, lstrProducto)
                lbooEncontrado = (lintRow > -1)
            End If
        End If
        If Not lbooEncontrado Then
            If ixdbDetalle.UpperBound(1) > 0 Then
                lintRow = ixdbDetalle.Find(ixdbDetalle.LowerBound(1), grdDetalle.Columns(iintColCodProducto).ColIndex, lstrProducto, XORDER_ASCEND, XCOMP_LT)
                lbooEncontrado = (lintRow > -1)
            End If
        End If
        If Not lbooEncontrado Then
            If ixdbDetalle.UpperBound(1) > 0 Then
                lintRow = ixdbDetalle.Find(ixdbDetalle.LowerBound(1), grdDetalle.Columns(iintColCodProducto).ColIndex, lstrProducto, XORDER_ASCEND, XCOMP_GT)
                lbooEncontrado = (lintRow > -1)
            End If
        End If
    End If
    'Buscamos por Descripcion
    If Not IsNumeric(lstrProducto) Then
        If Not lbooEncontrado Then
            If ixdbDetalle.UpperBound(1) > 0 Then
                lintRow = ixdbDetalle.Find(ixdbDetalle.LowerBound(1), grdDetalle.Columns(iintColDesProducto).ColIndex, lstrProducto)
                lbooEncontrado = (lintRow > -1)
            End If
        End If
        If Not lbooEncontrado Then
            If ixdbDetalle.UpperBound(1) > 0 Then
                lintRow = ixdbDetalle.Find(ixdbDetalle.LowerBound(1), grdDetalle.Columns(iintColDesProducto).ColIndex, lstrProducto, XORDER_ASCEND, XCOMP_LT)
                lbooEncontrado = (lintRow > -1)
            End If
        End If
        If Not lbooEncontrado Then
            If ixdbDetalle.UpperBound(1) > 0 Then
                lintRow = ixdbDetalle.Find(ixdbDetalle.LowerBound(1), grdDetalle.Columns(iintColDesProducto).ColIndex, lstrProducto, XORDER_ASCEND, XCOMP_GT)
                lbooEncontrado = (lintRow > -1)
            End If
        End If
    End If
    
    If lbooEncontrado Then
        grdDetalle.Bookmark = lintRow
        grdDetalle.SetFocus
    Else
        MsgBox "Producto no encontrado", vbInformation, "Aviso"
        txtProducto.SetFocus
    End If
   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub cmdCancelar_Click()
    
    Unload Me
    
End Sub

Private Sub cmdGrabar_Click()

On Error GoTo Control

    istrMensajePostGraba = ""
    istrGuiaGenerada = ""
    istrRegCompra = ""
    
    grdDetalle.EditActive = False

    If Not AntesDeGrabar Then
        Exit Sub
    End If
    
    'Invocamos el package para grabar
    grabar
    
    If istrMensajePostGraba <> "" Then
        MsgBox istrMensajePostGraba, vbInformation, "Aviso"
    End If
    If istrMensajePostGraba = "" Then
        MsgBox "REGISTRO GRABADO", vbInformation, "Aviso"
    End If
    
    'Impresion de Guia generada
    If istrCaso = "RECEPCION" And istrGuiaGenerada <> "" Then
        If MsgBox("¿Desea imprimir la Guia generada?" & vbCr & istrGuiaGenerada, vbYesNo Or vbDefaultButton1 Or vbQuestion, "Confirmación") = vbYes Then
            Dim objGuia As New clsGuia
            objGuia.spImprime_Guia_Dev "", "", istrGuiaGenerada
            Set objGuia = Nothing
        End If
    End If
    If istrRegCompra <> "" Then
        MsgBox "Registro de Compra Generado : " & istrRegCompra, vbExclamation, "Aviso"
    End If
    
    NuevoRegistro
    
    'Enfoque inicial
    If istrCaso = "RECEPCION" Then
        txtOCompra.Focus
    End If
    If istrCaso = "COMPRA" Then
        ctlcboMotivo.SetFocus
    End If
   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub ctlcboDocumento_LostFocus()
    MostrarCampos
End Sub

Private Sub Form_Load()
    
On Error GoTo Control
    
    Dim lstrTipo As String
    
    Move (Screen.Width - Width) / 2, ((Screen.Height - Height) / 2) - 800
    
    'Columnas en el array
    iintColumnas = 12
    
    'Seteamos array
    SetArray
    
    'Seteamos la Grilla
    SetGrid
    
    '--------------------------------------------------------------------------
    'Poblamos los combobox
    '--------------------------------------------------------------------------
    If istrCaso = "RECEPCION" Then
        lstrTipo = "1"
    End If
    If istrCaso = "COMPRA" Then
        lstrTipo = "0"
    End If
    
    Set ctlcboMotivo.RowSource = objCompra.ListaMotivo(lstrTipo)
    ctlcboMotivo.ListField = "DES"
    ctlcboMotivo.BoundColumn = "COD"
    ctlcboMotivo.BoundText = "*"

    Set ctlcboDocumento.RowSource = objCompra.ListaDocumento
    ctlcboDocumento.ListField = "DES"
    ctlcboDocumento.BoundColumn = "COD"
    ctlcboDocumento.BoundText = "*"
    '--------------------------------------------------------------------------
    
    'Seteamos campos segun sea el caso
    txtOCompra.Enabled = (istrCaso = "RECEPCION")
    cmdBuscar.Enabled = (istrCaso = "RECEPCION")
    ctlcboMotivo.Enabled = (istrCaso = "COMPRA")
    txtProveedor.Enabled = (istrCaso = "COMPRA")
    
    'Seteamos a solo lectura
    If ibooConsulta Then
        SoloLectura
    End If
    
    'Datos por defecto
    NuevoRegistro
    If ibooConsulta Then
        ConsultaRegistro
    End If
    
    MostrarCampos
   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
    
End Sub

Public Sub Mostrar(ByVal pstrCaso As String, _
                   Optional ByVal pbooConsulta As Boolean = False, _
                   Optional ByVal pstrProveedor As String = "", _
                   Optional ByVal pstrTipDocumento As String = "", _
                   Optional ByVal pstrNumDocumento As String = "")
    
    istrCaso = UCase(Trim(pstrCaso))
    ibooConsulta = pbooConsulta
    istrProveedor = pstrProveedor
    istrTipDocumento = pstrTipDocumento
    istrNumDocumento = pstrNumDocumento
    
    Me.Caption = istrCaso
    Me.Show vbModal

End Sub

Private Sub SetGrid()

    Dim i As Integer
    
    'Fijamos la ubicacion de las columnas dinamicas
    iintColCodProducto = 1
    iintColDesProducto = 2
    'Columnas confirmadas por el usuario
    iintColQFactura = 3
    iintColQFisica = 4
    iintColPrecio = 5
    'Columnas originales segun documento
    iintColQOFactura = 6
    iintColQOFisica = 7
    iintColOPrecio = 8
    'Columnas de apoyo
    iintColProductoProv = 9
    iintColObs = 12
    'Columnas adicionales ingresadas por el usuario
    iintColLote = 10
    iintColVcmto = 11
    
    'Deshabilitamos columnas
    For i = 0 To grdDetalle.Columns.Count - 1
        grdDetalle.Columns(i).AllowFocus = False
    Next i
    grdDetalle.MarqueeStyle = dbgFloatingEditor 'dbgHighlightCell
    
    'Columnas editables
    If istrCaso = "COMPRA" Then
        grdDetalle.Columns(iintColDesProducto).AllowFocus = True
        grdDetalle.Columns(iintColQFactura).AllowFocus = True
        grdDetalle.Columns(iintColQFisica).AllowFocus = True
        grdDetalle.Columns(iintColPrecio).AllowFocus = True
    End If
    If istrCaso = "RECEPCION" Then
        grdDetalle.Columns(iintColQFactura).AllowFocus = True
        grdDetalle.Columns(iintColQFisica).AllowFocus = True
        grdDetalle.Columns(iintColPrecio).AllowFocus = True
    End If
    grdDetalle.Columns(iintColLote).AllowFocus = True
    grdDetalle.Columns(iintColVcmto).AllowFocus = True
    
    'Formato de columnas tipo numero
    grdDetalle.Columns(iintColQFactura).NumberFormat = "#0"
    grdDetalle.Columns(iintColQFisica).NumberFormat = "#0"
    grdDetalle.Columns(iintColPrecio).NumberFormat = "#0." & String(2, "0")
    
    'Encabezado segun sea el caso
    If istrCaso = "COMPRA" Then
        grdDetalle.Columns(iintColQFactura).Caption = "Cant. Fracción"
        grdDetalle.Columns(iintColQFisica).Caption = "Cant. Unidad"
        grdDetalle.Columns(iintColPrecio).Caption = "SubTotal"
    End If
    If istrCaso = "RECEPCION" Then
        grdDetalle.Columns(iintColQFactura).Caption = "Cant. Factura"
        grdDetalle.Columns(iintColQFisica).Caption = "Cant. Física"
    End If
    grdDetalle.Columns(iintColCodProducto).Caption = "Código"
    grdDetalle.Columns(iintColDesProducto).Caption = "Descripción"
    grdDetalle.Columns(iintColLote).Caption = "Lote"
    grdDetalle.Columns(iintColVcmto).Caption = "Vencimiento"
    grdDetalle.Columns(iintColObs).Caption = "Observaciones"
    
    'Tamaño de las columnas
    grdDetalle.Columns(iintColCodProducto).Width = 6 * 120
    grdDetalle.Columns(iintColDesProducto).Width = 40 * 120
    grdDetalle.Columns(iintColQFactura).Width = 10 * 120
    grdDetalle.Columns(iintColQFisica).Width = 10 * 120
    grdDetalle.Columns(iintColPrecio).Width = 10 * 120
    grdDetalle.Columns(iintColLote).Width = 10 * 120
    grdDetalle.Columns(iintColVcmto).Width = 8 * 120
    grdDetalle.Columns(iintColObs).Width = 40 * 120
    
    'Alineamiento
    grdDetalle.Columns(iintColQFactura).Alignment = dbgRight
    grdDetalle.Columns(iintColQFisica).Alignment = dbgRight
    grdDetalle.Columns(iintColPrecio).Alignment = dbgRight
    
    'Celdas que pueden cambiar de apariencia
    grdDetalle.Columns(iintColQFactura).FetchStyle = True
    grdDetalle.Columns(iintColQFisica).FetchStyle = True
    grdDetalle.Columns(iintColPrecio).FetchStyle = True
    
    grdDetalle.Array = ixdbDetalle

End Sub

Private Sub SetArray()

On Error GoTo handle
    ixdbDetalle.ReDim 0, -1, 0, iintColumnas
    grdDetalle.Array = ixdbDetalle
    grdDetalle.Rebind
    
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub CargaDatos(ByVal pstrOCompra As String)
    
    Dim rsDatos As oraDynaset
    Dim rsDetalle As oraDynaset
    Dim i As Integer
    
    Set rsDatos = objCompra.OrdenCompra(pstrOCompra, objUsuario.CodigoLocal)
    Set rsDetalle = objCompra.OrdenCompraDetalle(pstrOCompra, "D", objUsuario.CodigoLocal)
    
    If rsDatos Is Nothing Or rsDetalle Is Nothing Then
        Exit Sub
    End If
    
    ctlcboMotivo.BoundText = rsDatos("COD_TIPO").Value
    lblFormaPago.Caption = rsDatos("COD_FPAGO").Value & "-" & rsDatos("DES_FPAGO").Value
    lblObservaciones.Caption = rsDatos("OBS").Value
    txtProveedor.Text = rsDatos("RUC").Value
    lblProveedor.Caption = rsDatos("PROVEEDOR").Value
    
    'Caso contado en la recepcion
    MostrarCampos
    
    While Not rsDetalle.EOF
        ixdbDetalle.AppendRows
        ixdbDetalle(i, 0) = i
        
        If istrCaso = "RECEPCION" Then
            ixdbDetalle(i, iintColCodProducto) = rsDetalle("COD_PRODUCTO").Value
            ixdbDetalle(i, iintColDesProducto) = rsDetalle("DES_PRODUCTO").Value
            ixdbDetalle(i, iintColQFactura) = 0
            ixdbDetalle(i, iintColQFisica) = 0
            ixdbDetalle(i, iintColPrecio) = rsDetalle("PRC_FINAL_UNIT").Value
            ixdbDetalle(i, iintColQOFactura) = rsDetalle("CTD_PRODUCTO").Value
            ixdbDetalle(i, iintColQOFisica) = rsDetalle("CTD_PRODUCTO").Value
            ixdbDetalle(i, iintColOPrecio) = rsDetalle("PRC_FINAL_UNIT").Value
            ixdbDetalle(i, iintColProductoProv) = rsDetalle("COD_PRODUCTO_PROV").Value
            ixdbDetalle(i, iintColLote) = ""
            ixdbDetalle(i, iintColVcmto) = ""
            ixdbDetalle(i, iintColObs) = "" & rsDetalle("OBS_COMPRA").Value
        End If

        i = i + 1
        rsDetalle.MoveNext
    Wend
    grdDetalle.Rebind
    totaliza
    
End Sub

Private Sub grdDetalle_AfterColUpdate(ByVal ColIndex As Integer)
    
On Error GoTo Control

    Dim intFila%
    Dim lrsProducto As oraDynaset
    Dim lstrCodProducto As String
    Dim lstrAno As String
    Dim lstrMes As String

    intFila = grdDetalle.Bookmark
    
    Select Case ColIndex
        Case iintColDesProducto
            Dim strBusca$
            strBusca = UCase(Trim(grdDetalle.Columns(iintColDesProducto).Text))
            If strBusca = "" Then
                ixdbDetalle(intFila, iintColDesProducto) = ""
                ixdbDetalle(intFila, iintColCodProducto) = ""
                grdDetalle.Rebind
                Exit Sub
            End If
            
            Set lrsProducto = objProducto.ListaBusqueda(strBusca)
            
            If lrsProducto(0).Value = -1 Then
                MsgBox lrsProducto(1).Value, vbExclamation, "Alerta"
                ixdbDetalle(intFila, iintColCodProducto) = ""
                ixdbDetalle(intFila, iintColDesProducto) = ""
                Exit Sub
            End If
            
            If lrsProducto.RecordCount <= 0 Then
                ixdbDetalle(intFila, iintColCodProducto) = ""
                ixdbDetalle(intFila, iintColDesProducto) = ""
            End If
            
            If lrsProducto.RecordCount = 1 Then
                ixdbDetalle(intFila, iintColCodProducto) = CStr(lrsProducto("COD").Value)
                ixdbDetalle(intFila, iintColDesProducto) = CStr(lrsProducto("DES").Value)
            End If
            
            If lrsProducto.RecordCount > 1 Then
                Dim lfrmProducto As New frm_INV_ProductoDatos
                Dim larrCabecera As Variant

                ReDim larrCabecera(0 To 7)
                larrCabecera(0) = "Código"
                larrCabecera(1) = "Descripción"
                larrCabecera(2) = "Estado"
                larrCabecera(3) = "Cód.Lab."
                larrCabecera(4) = "Laboratorio"
                larrCabecera(5) = "Abrev."
                larrCabecera(6) = "Cód.Lín"
                larrCabecera(7) = "Línea"
                
                Screen.MousePointer = vbHourglass
                Set lfrmProducto.irsProducto = lrsProducto
                lfrmProducto.iarrCabecera = larrCabecera
                lfrmProducto.Show vbModal
                ixdbDetalle(intFila, iintColCodProducto) = CStr(lfrmProducto.istrCodProducto)
                ixdbDetalle(intFila, iintColDesProducto) = CStr(lfrmProducto.istrDesProducto)
                ixdbDetalle(intFila, iintColProductoProv) = ""
                Screen.MousePointer = vbDefault
            End If
            
        Case iintColQFactura
            If IsNumeric(grdDetalle.Columns(iintColQFactura).Text) Then
                ixdbDetalle(intFila, iintColQFactura) = Val(grdDetalle.Columns(iintColQFactura).Text)
            Else
                ixdbDetalle(intFila, iintColQFactura) = 0
            End If
            
        Case iintColQFisica
            If IsNumeric(grdDetalle.Columns(iintColQFisica).Text) Then
                ixdbDetalle(intFila, iintColQFisica) = Val(grdDetalle.Columns(iintColQFisica).Text)
            Else
                ixdbDetalle(intFila, iintColQFisica) = 0
            End If
            
        Case iintColPrecio
            If IsNumeric(grdDetalle.Columns(iintColPrecio).Text) Then
                ixdbDetalle(intFila, iintColPrecio) = Val(grdDetalle.Columns(iintColPrecio).Text)
            Else
                ixdbDetalle(intFila, iintColPrecio) = 0
            End If
            If istrCaso = "COMPRA" And Trim(ixdbDetalle(intFila, iintColCodProducto)) <> "" And Trim(ixdbDetalle(intFila, iintColDesProducto)) <> "" Then
                If grdDetalle.Bookmark = ixdbDetalle.UpperBound(1) Then
                    ixdbDetalle.AppendRows
                    ixdbDetalle(intFila + 1, 0) = grdDetalle.Bookmark + 2
                    intFila = intFila + 1
                    ixdbDetalle(intFila, iintColQFactura) = 0
                    ixdbDetalle(intFila, iintColQFisica) = 0
                    ixdbDetalle(intFila, iintColPrecio) = 0
                End If
            End If
            
        Case iintColLote
            ixdbDetalle(intFila, iintColLote) = Mid(Trim(grdDetalle.Columns(iintColLote).Text), 1, 40)
            
        Case iintColVcmto
            If IsNumeric(grdDetalle.Columns(iintColVcmto).Text) Then
                ixdbDetalle(intFila, iintColVcmto) = Mid(Trim(grdDetalle.Columns(iintColVcmto).Text), 1, 6)
            Else
                ixdbDetalle(intFila, iintColVcmto) = ""
            End If
            
            If Len(ixdbDetalle(intFila, iintColVcmto)) <> 6 Then
                ixdbDetalle(intFila, iintColVcmto) = ""
            Else
                lstrAno = Mid(ixdbDetalle(intFila, iintColVcmto), 1, 4)
                lstrMes = Mid(ixdbDetalle(intFila, iintColVcmto), 5, 2)
                
                'Validamos el vencimiento (debe tener formato de periodo)
                If Not (Val(lstrAno) >= 2000) Or _
                   Not (Val(lstrMes) >= 1 And Val(lstrMes) <= 12) Then
                    ixdbDetalle(intFila, iintColVcmto) = ""
                End If
            End If
    
    End Select
            
    'Consultar si el producto fracciona
    If istrCaso = "COMPRA" Then
        lstrCodProducto = Trim(grdDetalle.Columns(iintColCodProducto).Text)
        If Not Fracciona(lstrCodProducto) Then
            ixdbDetalle(intFila, iintColQFactura) = 0
        Else
            ixdbDetalle(intFila, iintColQFisica) = 0
        End If
    End If
    
    RenumeraItem
    grdDetalle.Rebind
    grdDetalle.Bookmark = intFila
   
    Select Case ColIndex
        Case iintColDesProducto
            grdDetalle.Col = ColIndex
        Case iintColQFactura
            grdDetalle.Col = ColIndex
        Case iintColQFisica
            grdDetalle.Col = ColIndex
        Case iintColPrecio
            If istrCaso = "RECEPCION" Then
                grdDetalle.Col = ColIndex
            End If
        Case iintColLote
            grdDetalle.Col = ColIndex
        Case iintColVcmto
            'grdDetalle.Col = iintColDesProducto
    End Select
    
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
    
End Sub

Private Sub grdDetalle_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If (ColIndex = iintColPrecio And istrCaso = "COMPRA") Or (ColIndex = iintColVcmto And istrCaso = "RECEPCION") Then
        grdDetalle.DirectionAfterEnter = dbgMoveNone
    Else
        grdDetalle.DirectionAfterEnter = dbgMoveRight
    End If
End Sub

Private Sub grdDetalle_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)

On Error GoTo Control

    Dim ldblValor As Double
    Dim ldblOValor As Double

    ldblValor = Val(grdDetalle.Columns(Col).CellValue(Bookmark))

    Select Case Col
    Case iintColQFactura
        If istrCaso = "RECEPCION" Then
            ldblOValor = Val(grdDetalle.Columns(iintColQOFactura).CellValue(Bookmark))
            If ldblValor <> 0 And ldblValor <> ldblOValor Then
                'CellStyle.BackColor = RGB(204, 255, 204)
                CellStyle.ForeColor = &HC0&
                CellStyle.Font.Bold = True
            End If
        End If
    Case iintColQFisica
        If istrCaso = "RECEPCION" Then
            ldblOValor = Val(grdDetalle.Columns(iintColQOFisica).CellValue(Bookmark))
            If ldblValor <> 0 And ldblValor <> ldblOValor Then
                'CellStyle.BackColor = RGB(204, 255, 204)
                CellStyle.ForeColor = &HC0&
                CellStyle.Font.Bold = True
            End If
        End If
    Case iintColPrecio
        If istrCaso = "RECEPCION" Then
            ldblOValor = Val(grdDetalle.Columns(iintColOPrecio).CellValue(Bookmark))
            If ldblValor <> 0 And ldblValor <> ldblOValor Then
                'CellStyle.BackColor = RGB(204, 255, 204)
                CellStyle.ForeColor = &HC0&
                CellStyle.Font.Bold = True
            End If
        End If
    End Select
    Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub grdDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown As Boolean
Dim lintFila As Integer

On Error GoTo Control

CtrlDown = (Shift And vbCtrlMask) > 0

Select Case KeyCode
    Case vbKeyDelete And CtrlDown
        If istrCaso = "COMPRA" Then
            lintFila = grdDetalle.Bookmark
            If Trim(ixdbDetalle(lintFila, iintColDesProducto)) = "" And lintFila > 0 Then
                EliminaItem
            Else
                If MsgBox("¿Está seguro que desea eliminar el item?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación") = vbYes Then
                    EliminaItem
                End If
            End If
        End If
    Case vbKeyReturn
        If grdDetalle.Col = iintColDesProducto Then
            grdDetalle.DirectionAfterEnter = dbgMoveRight
        End If
End Select
   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub grdDetalle_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
On Error GoTo Control

    totaliza
   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub totaliza()

    Dim i As Integer
    Dim lintFraccionamiento As Integer
    Dim lintUnidades As Integer
    Dim lstrCodProducto As String
    Dim ldblValorLinea As Double
    Dim ldblImpuestoLinea As Double
    Dim ldblBaseLinea As Double
    Dim ldblSubtotal As Double
    Dim ldblImpuesto As Double
    Dim ldblInafecto As Double
    Dim ldblTotal As Double
    
    ldblSubtotal = 0
    ldblImpuesto = 0
    ldblInafecto = 0
    ldblTotal = 0
    
    While i <= ixdbDetalle.UpperBound(1)
        lstrCodProducto = ixdbDetalle(i, iintColCodProducto)
        lintFraccionamiento = 1
        If istrCaso = "COMPRA" Then
            If Fracciona(lstrCodProducto) Then
                lintFraccionamiento = Fraccionamiento(lstrCodProducto)
                lintUnidades = Val(ixdbDetalle(i, iintColQFisica)) * lintFraccionamiento + Val(ixdbDetalle(i, iintColQFactura))
            Else
                lintUnidades = Val(ixdbDetalle(i, iintColQFisica))
            End If
            ldblValorLinea = Val(ixdbDetalle(i, iintColPrecio))
        End If
        If istrCaso = "RECEPCION" Then
            lintUnidades = Val(ixdbDetalle(i, iintColQFisica))
            ldblValorLinea = Round(lintUnidades * Val(ixdbDetalle(i, iintColPrecio)), 2)
        End If
        
        If Inafecto(lstrCodProducto) Then
            ldblBaseLinea = ldblValorLinea
            ldblImpuestoLinea = 0
            ldblInafecto = ldblInafecto + ldblValorLinea
        Else
            ldblBaseLinea = Round(ldblValorLinea / (1 + (objUsuario.IGV / 100)), 2)
            ldblImpuestoLinea = ldblValorLinea - ldblBaseLinea
            ldblSubtotal = ldblSubtotal + ldblBaseLinea
        End If
        
        ldblImpuesto = ldblImpuesto + ldblImpuestoLinea
        ldblTotal = ldblTotal + ldblBaseLinea + ldblImpuestoLinea
        
        i = i + 1
    Wend
    lblSubtotal.Caption = Format(ldblSubtotal, "###,##0.00")
    lblImpuesto.Caption = Format(ldblImpuesto, "###,##0.00")
    lblInafecto.Caption = Format(ldblInafecto, "###,##0.00")
    lblTotal.Caption = Format(ldblTotal, "###,##0.00")
    
    'Calculamos el redondeo
    Redondeo
    
End Sub

Private Sub RenumeraItem()
    Dim i As Long
    
    For i = 0 To ixdbDetalle.UpperBound(1)
        ixdbDetalle(i, 0) = i
    Next i
End Sub

Private Sub EliminaItem()

    If grdDetalle.ApproxCount > 1 Then
        grdDetalle.Delete
    ElseIf grdDetalle.ApproxCount = 1 Then
        grdDetalle.Delete
        ixdbDetalle.ReDim 0, -1, 0, iintColumnas
    End If
    RenumeraItem

End Sub

Private Function AntesDeGrabar() As Boolean

    Dim lstrMensaje As String
    Dim lstrOCompra  As String
    Dim lstrFormaPago As String
    Dim lstrMotivo  As String
    Dim lstrProveedor  As String
    Dim lstrDocumento  As String
    Dim lstrNumDocumento As String
    Dim lstrFecha As String
    Dim ldblSubtotal As Double
    Dim ldblImpuesto As Double
    Dim ldblInafecto As Double
    Dim ldblTotal As Double
    Dim lstrCodProducto As String
    Dim lintQFactura As Integer
    Dim lintQFisica As Integer
    Dim ldblPrecio As Double
    Dim lintItemValido As Integer
    Dim lbooExito As Boolean
    
    lstrOCompra = Trim(txtOCompra.Text)
    lstrMotivo = ctlcboMotivo.BoundText
    lstrFormaPago = Mid(lblFormaPago.Caption, 1, 3)
    lstrProveedor = Trim(txtProveedor.Text)
    lstrDocumento = ctlcboDocumento.BoundText
    lstrNumDocumento = Trim(txtDocumento.Text)
    lstrFecha = Format(dtpFecha.Value, "dd/mm/yyyy")
    ldblSubtotal = Val(txtSubtotal.Text)
    ldblImpuesto = Val(txtImpuesto.Text)
    ldblInafecto = Val(txtInafecto.Text)
    ldblTotal = Val(txtTotal.Text)

    lstrMensaje = ""
    
    '---------------------------------------------------------------------------
    'Validaciones de cabecera
    '---------------------------------------------------------------------------
    If istrCaso = "RECEPCION" Then
        
        If lstrMensaje = "" Then
            If lstrOCompra = "" Then
                lstrMensaje = "Debe indicar la Orden de Compra"
                txtOCompra.Focus
            End If
        End If
    
    End If
    
    If lstrMensaje = "" Then
        If lstrMotivo = "" Or lstrMotivo = "*" Then
            lstrMensaje = "Debe indicar el Motivo"
            ctlcboMotivo.SetFocus
        End If
    End If
    If lstrMensaje = "" Then
        If lstrProveedor = "" Then
            lstrMensaje = "Debe indicar el Proveedor"
            txtProveedor.Focus
        End If
    End If
    If lstrMensaje = "" Then
        If lstrDocumento = "" Or lstrDocumento = "*" Then
            lstrMensaje = "Debe indicar el Tipo de Documento"
            ctlcboDocumento.SetFocus
        End If
    End If
    If lstrMensaje = "" Then
        If lstrNumDocumento = "" Then
            lstrMensaje = "Debe indicar el Numero de Documento"
            txtDocumento.Focus
        End If
    End If
    'Aqui debe ir la validacion de fechas
    If lstrMensaje = "" Then
        If (ldblTotal = 0) Or _
           ((ldblSubtotal <> 0 And ldblImpuesto = 0) Or (ldblSubtotal = 0 And ldblImpuesto <> 0)) Or _
           (ldblSubtotal = 0 And ldblInafecto = 0) Then
            If txtSubtotal.Visible Then
                lstrMensaje = "Debe indicar los Totales"
                txtSubtotal.Focus
            End If
        End If
    End If
    If lstrMensaje = "" Then
        If ldblSubtotal < 0 Or ldblImpuesto < 0 Or ldblInafecto < 0 Or ldblTotal < 0 Then
            If txtSubtotal.Visible Then
                lstrMensaje = "No se permiten negativos en los Totales"
                txtSubtotal.Focus
            End If
        End If
    End If
    '---------------------------------------------------------------------------
    
    '---------------------------------------------------------------------------
    'Validaciones del detalle
    '---------------------------------------------------------------------------
    If lstrMensaje = "" Then
        Do While True
            i = 0
            lbooExito = False
            Do While i <= ixdbDetalle.UpperBound(1)
                If Trim(ixdbDetalle(i, iintColDesProducto)) = "" And ixdbDetalle.UpperBound(1) > 0 Then
                    lbooExito = True
                    Exit Do
                End If
                i = i + 1
            Loop
            If Not lbooExito Then
                Exit Do
            End If
            grdDetalle.Bookmark = i
            EliminaItem
        Loop
        grdDetalle.Bookmark = 0
        grdDetalle.Rebind
        Me.Refresh
    End If

    If lstrMensaje = "" Then
        i = 0
        lintItemValido = 0
        Do While i <= ixdbDetalle.UpperBound(1)
            lstrCodProducto = Trim(ixdbDetalle(i, iintColCodProducto))
            lintQFactura = Val(ixdbDetalle(i, iintColQFactura))
            lintQFisica = Val(ixdbDetalle(i, iintColQFisica))
            ldblPrecio = Val(ixdbDetalle(i, iintColPrecio))
            
            If lstrCodProducto = "" Then
                lstrMensaje = "Debe indicar el Producto"
                Exit Do
            End If
            If istrCaso = "RECEPCION" Then
                'Caso contado
                If lstrFormaPago = "005" Then
                    If Not (lintQFactura = 0 And lintQFisica = 0) Then
                        If lintQFisica > lintQFactura Then
                            lstrMensaje = "La cantidad fisica no debe ser mayor que la cantidad facturada"
                            Exit Do
                        End If
                        If lintQFisica < lintQFactura Then
                            istrMensajePostGraba = "Genere boleta por los productos faltantes"
                        End If
                    End If
                End If
                'Caso credito
                If lstrFormaPago <> "005" Then
                    If Not (lintQFactura = 0 And lintQFisica = 0) Then
                        If lintQFisica > lintQFactura Then
                            lstrMensaje = "La cantidad fisica no debe ser mayor que la cantidad facturada"
                            Exit Do
                        End If
                    End If
                End If
                If ldblPrecio <= 0 Then
                    lstrMensaje = "Debe indicar el precio"
                    Exit Do
                End If
            End If
            If istrCaso = "COMPRA" Then
                
'                lstrFormaPago = "003"
                
                If lintQFactura < 0 Then
                    lstrMensaje = "La cantidad en fracciones debe ser positiva"
                    Exit Do
                End If
                If lintQFisica < 0 Then
                    lstrMensaje = "La cantidad en unidades debe ser positiva"
                    Exit Do
                End If
                If lintQFactura = 0 And lintQFisica = 0 Then
                    lstrMensaje = "Debe indicar cantidad en fracciones o unidades"
                    Exit Do
                End If
                If ldblPrecio <= 0 Then
                    lstrMensaje = "Debe indicar el subtotal"
                    Exit Do
                End If
            End If
            
            'Contamos el numero de items validos
            If Not (lintQFactura = 0 And lintQFisica = 0) Then
                lintItemValido = lintItemValido + 1
            End If
            
            i = i + 1
        Loop
        If lstrMensaje <> "" Then
            grdDetalle.Bookmark = i
            grdDetalle.SetFocus
        End If
        If lstrMensaje = "" And i = 0 Then
            lstrMensaje = "Debe indicar el detalle"
        End If
        If lstrMensaje = "" And lintItemValido = 0 Then
            lstrMensaje = "Debe indicar cantidades a registrar"
        End If
    End If
    '---------------------------------------------------------------------------
    
    AntesDeGrabar = True
    If lstrMensaje <> "" Then
        MsgBox lstrMensaje, vbExclamation, "Validaciones"
        AntesDeGrabar = False
    End If

End Function

Private Function Fracciona(ByVal pstrProducto As String) As Boolean
On Error GoTo Excepcion
    Dim lstrFracciona As String

    'Consultar si el producto fracciona
    lstrFracciona = objProducto.ListaDevFracciona(pstrProducto, objUsuario.CodigoLocal, objVenta.CodModalidadVenta)
    Fracciona = (lstrFracciona <> "0")
    Exit Function
Excepcion:
    Fracciona = False
End Function

Private Function Fraccionamiento(ByVal pstrProducto As String) As Integer
On Error GoTo Excepcion
    Fraccionamiento = objProducto.intCtdFrac(pstrProducto)
    Exit Function
Excepcion:
    Fraccionamiento = 0
End Function

Private Function Inafecto(ByVal pstrProducto As String) As Boolean
On Error GoTo Excepcion
    Inafecto = (objProducto.FlgInafecto(pstrProducto) = "1")
    Exit Function
Excepcion:
    Inafecto = False
End Function

Private Sub txtProveedor_Validate(Cancel As Boolean)

    Dim lstrProveedor As String
    
    lblProveedor.Caption = ""
    lstrProveedor = Trim(txtProveedor.Text)
    
    If lstrProveedor <> "" Then
        lblProveedor.Caption = Trim(objProveedor.Nombre(lstrProveedor))
        If lblProveedor.Caption = "" Then
            MsgBox "Proveedor no encontrado", vbExclamation, "Aviso"
            Cancel = True
        End If
    End If

End Sub

Private Sub grabar()

    Dim lstrCia As String
    Dim lstrOCompra As String
    Dim lstrMotivo As String
    Dim lstrFormaPago As String
    Dim lstrLocal As String
    Dim lstrProveedor As String
    Dim lstrDocumento As String
    Dim lstrNumDocumento As String
    Dim lstrFecha As String
    Dim lstrModulo As String
    Dim lstrUsuario As String
    Dim ldblAfecto As Double
    Dim ldblImpuesto As Double
    Dim ldblInafecto As Double
    Dim ldblTotal As Double
    Dim ldblRedondeo As Double
    Dim lstrFlgContabilizar As String
    
    Dim lparProductoProv As Variant
    Dim lparCantidad As Variant
    Dim lparCantidadFrac As Variant
    Dim lparProducto As Variant
    Dim lparCantidadRec As Variant
    Dim lparCantidadFracRec As Variant
    Dim lparPrecio As Variant
    Dim lparLote As Variant
    Dim lparFechaVcmto As Variant
    
    Dim lintFraccionamiento As Integer
    
    Dim i As Integer
    Dim lintFilas As Integer
    
    lstrCia = Trim(objUsuario.CodigoEmpresa)
    lstrOCompra = Trim(txtOCompra.Text)
    lstrMotivo = Trim(ctlcboMotivo.BoundText)
    lstrFormaPago = Mid(lblFormaPago.Caption, 1, 3)
    lstrLocal = Trim(objUsuario.CodigoLocal)
    lstrProveedor = Trim(txtProveedor.Text)
    lstrDocumento = Trim(ctlcboDocumento.BoundText)
    lstrNumDocumento = Trim(txtDocumento.Text)
    lstrFecha = Format(dtpFecha.Value, "dd/mm/yyyy")
    lstrModulo = Trim(objUsuario.NombrePC)
    lstrUsuario = Trim(objUsuario.Codigo)
    
    ldblAfecto = Val(txtSubtotal.Text)
    ldblImpuesto = Val(txtImpuesto.Text)
    ldblInafecto = Val(txtInafecto.Text)
    ldblTotal = Val(txtTotal.Text)
    ldblRedondeo = Val(txtRedondeo.Text)
    '------------------------------------------------
    'Ccieza fecha 24/03/2010
    'Motivo para recepcion de OC en locales y no hacer el quiebre
    If istrCaso = "COMPRA" Then
        lstrFormaPago = "005"
    End If
    '------------------------------------------------
    
    'Caso de boleta -> los totales se envian de forma distinta
    If istrCaso = "COMPRA" And ctlcboDocumento.BoundText = "BOL" Then
        ldblAfecto = 0
        ldblImpuesto = 0
        ldblInafecto = ldblTotal
    End If
    
    lintFilas = 0
    Do While i <= ixdbDetalle.UpperBound(1)
        If Not (Val(ixdbDetalle(i, iintColQFactura)) = 0 And Val(ixdbDetalle(i, iintColQFisica)) = 0) Then
            lintFilas = lintFilas + 1
        End If
        i = i + 1
    Loop
    
    'Contabilizar
    lstrFlgContabilizar = "0"
    If (istrCaso = "COMPRA") Or (istrCaso = "RECEPCION" And lstrFormaPago = "005") Then
        lstrFlgContabilizar = "1"
    End If
    
    ReDim lparProductoProv(1 To lintFilas)
    ReDim lparProducto(1 To lintFilas)
    ReDim lparPrecio(1 To lintFilas)
    ReDim lparLote(1 To lintFilas)
    ReDim lparFechaVcmto(1 To lintFilas)
    ReDim lparCantidad(1 To lintFilas)
    ReDim lparCantidadFrac(1 To lintFilas)
    ReDim lparCantidadRec(1 To lintFilas)
    ReDim lparCantidadFracRec(1 To lintFilas)
    
    i = 0
    lintFilas = 1
    Do While i <= ixdbDetalle.UpperBound(1)
        If Not (Val(ixdbDetalle(i, iintColQFactura)) = 0 And Val(ixdbDetalle(i, iintColQFisica)) = 0) Then
            lparProductoProv(lintFilas) = Trim(ixdbDetalle(i, iintColProductoProv))
            lparProducto(lintFilas) = Trim(ixdbDetalle(i, iintColCodProducto))
            lparLote(lintFilas) = Trim(ixdbDetalle(i, iintColLote))
            lparFechaVcmto(lintFilas) = Trim(ixdbDetalle(i, iintColVcmto))
            
            'Formato requerido por el package
            If Len(lparFechaVcmto(lintFilas)) = 6 Then
                lparFechaVcmto(lintFilas) = "01/" & Mid(lparFechaVcmto(lintFilas), 5, 2) & "/" & Mid(lparFechaVcmto(lintFilas), 1, 4)
            End If
            
            If istrCaso = "RECEPCION" Then
                lparCantidad(lintFilas) = Val(ixdbDetalle(i, iintColQFactura))
                lparCantidadFrac(lintFilas) = 0
                lparCantidadRec(lintFilas) = Val(ixdbDetalle(i, iintColQFisica))
                lparCantidadFracRec(lintFilas) = 0
                '09/10/2007 : Razuri hizo un cambio para que se ahora se envie el precio por la linea (item)
                'lparPrecio(lintFilas) = Val(ixdbDetalle(i, iintColPrecio))
                lparPrecio(lintFilas) = Val(ixdbDetalle(i, iintColPrecio)) * Val(ixdbDetalle(i, iintColQFisica))
            End If
            If istrCaso = "COMPRA" Then
                lparCantidad(lintFilas) = Val(ixdbDetalle(i, iintColQFisica))
                lparCantidadFrac(lintFilas) = Val(ixdbDetalle(i, iintColQFactura))
                lparCantidadRec(lintFilas) = Val(ixdbDetalle(i, iintColQFisica))
                lparCantidadFracRec(lintFilas) = Val(ixdbDetalle(i, iintColQFactura))
                
                '09/10/2007 : Razuri hizo un cambio para que se ahora se envie el precio por la linea (item)
                lparPrecio(lintFilas) = 0
                If lparCantidad(lintFilas) <> 0 Then
                    'lparPrecio(lintFilas) = Round(Val(ixdbDetalle(i, iintColPrecio)) / lparCantidad(lintFilas), 2)
                    lparPrecio(lintFilas) = Round(Val(ixdbDetalle(i, iintColPrecio)), 2)
                End If
                If lparCantidadFrac(lintFilas) <> 0 Then
                    lintFraccionamiento = Fraccionamiento(lparProducto(lintFilas))
                    'lparPrecio(lintFilas) = Round((Val(ixdbDetalle(i, iintColPrecio)) * lintFraccionamiento) / lparCantidadFrac(lintFilas), 2)
                    lparPrecio(lintFilas) = Round(Val(ixdbDetalle(i, iintColPrecio)), 2)
                End If
            End If
            lintFilas = lintFilas + 1
        End If
        i = i + 1
    Loop
    
    Dim objCompras As New clsCompra
    objCompras.GrabaRecepcion lstrCia, lstrOCompra, lstrMotivo, lstrFormaPago, lstrLocal, _
                              lstrProveedor, lstrDocumento, lstrNumDocumento, lstrFecha, _
                              lstrModulo, lstrUsuario, ldblAfecto, ldblImpuesto, ldblInafecto, _
                              ldblTotal, ldblRedondeo, "0", lstrFlgContabilizar, istrGuiaGenerada, _
                              istrRegCompra, _
                              lparProductoProv, lparCantidad, lparCantidadFrac, _
                              lparProducto, lparCantidadRec, lparCantidadFracRec, lparPrecio, _
                              lparLote, lparFechaVcmto

End Sub

Private Sub NuevoRegistro()

    txtOCompra.Text = ""
    lblFormaPago.Caption = ""
    lblObservaciones.Caption = ""
    ctlcboMotivo.BoundText = "*"
    txtProveedor.Text = ""
    lblProveedor.Caption = ""
    ctlcboDocumento.BoundText = "*"
    txtDocumento.Text = ""
    dtpFecha.Value = CDate(Format(Now, "dd/mm/yyyy"))
    txtProducto.Text = ""
    txtSubtotal.Text = 0
    txtImpuesto.Text = 0
    txtInafecto.Text = 0
    txtTotal.Text = 0
    txtRedondeo.Text = 0
    
    'Setamos Grilla
    SetArray

    'Fila por defecto
    If istrCaso = "COMPRA" Then
        ixdbDetalle.AppendRows
        ixdbDetalle(0, 0) = 1
        grdDetalle.Rebind
        RenumeraItem
        grdDetalle.Bookmark = 0
    End If

End Sub

Private Sub Redondeo()

    Dim ldblRedondeo As Double

    If istrCaso = "COMPRA" And ctlcboDocumento.BoundText = "BOL" Then
        ldblRedondeo = Round(Val(lblTotal.Caption) - Val(txtTotal.Text), 2)
        txtRedondeo.Text = ldblRedondeo
    End If

End Sub

Private Sub txtTotal_LostFocus()
    Redondeo
End Sub

Private Sub SoloLectura()

    txtOCompra.Enabled = (Not ibooConsulta)
    cmdBuscar.Enabled = (Not ibooConsulta)
    ctlcboMotivo.Enabled = (Not ibooConsulta)
    txtProveedor.Enabled = (Not ibooConsulta)
    ctlcboDocumento.Enabled = (Not ibooConsulta)
    txtDocumento.Enabled = (Not ibooConsulta)
    dtpFecha.Enabled = (Not ibooConsulta)
    txtProducto.Enabled = (Not ibooConsulta)
    cmdBuscarPrd.Enabled = (Not ibooConsulta)
    
    grdDetalle.AllowAddNew = (Not ibooConsulta)
    grdDetalle.AllowUpdate = (Not ibooConsulta)
    grdDetalle.AllowDelete = (Not ibooConsulta)
    
    txtSubtotal.Enabled = (Not ibooConsulta)
    txtImpuesto.Enabled = (Not ibooConsulta)
    txtInafecto.Enabled = (Not ibooConsulta)
    txtTotal.Enabled = (Not ibooConsulta)
    txtRedondeo.Enabled = (Not ibooConsulta)
    
    cmdGrabar.Enabled = (Not ibooConsulta)

End Sub

Private Sub ConsultaRegistro()

On Error GoTo Control

    Dim objCompra As New clsCompra
    
    Set rstCab = objCompra.ParteRecepcion(istrProveedor, istrNumDocumento, istrTipDocumento)
    Set rstDet = objCompra.ParteRecepcionItems(istrProveedor, istrNumDocumento, istrTipDocumento)
    
    Set objCompra = Nothing
    
    'Cabecera y Pie
    If rstCab.EOF Then
        Exit Sub
    End If

    txtOCompra.Text = rstCab("NUM_ORDEN_COMPRA").Value
    lblFormaPago.Caption = rstCab("COD_FORMA_PAGO").Value & "-" & rstCab("DES_FORMA_PAGO").Value
    ctlcboMotivo.BoundText = Trim(rstCab("COD_TIPO_ORD_COM").Value)
    txtProveedor.Text = rstCab("RUC_PROVEEDOR").Value
    lblProveedor.Caption = rstCab("DES_PROVEEDOR").Value
    ctlcboDocumento.BoundText = rstCab("TIP_DOCUMENTO").Value
    txtDocumento.Text = rstCab("NUM_DOCUMENTO").Value
    dtpFecha.Value = CDate(Format(rstCab("FCH_EMISION").Value, "dd/mm/yyyy"))
    txtProducto.Text = ""
    txtSubtotal.Text = rstCab("MTO_BASE_IMP_MAN").Value
    txtImpuesto.Text = rstCab("MTO_IMPUESTO_MAN").Value
    txtInafecto.Text = rstCab("MTO_INAFECTO_MAN").Value
    txtTotal.Text = rstCab("MTO_TOTAL_MAN").Value
    txtRedondeo.Text = rstCab("MTO_REDONDEO_DOC").Value
    
    'Ocultar campos caso contado
    MostrarCampos
    
    'Detalle
    If rstDet.EOF Then
        Exit Sub
    End If
    
    SetArray

    While Not rstDet.EOF
        ixdbDetalle.AppendRows
        ixdbDetalle(i, 0) = i
        
        ixdbDetalle(i, iintColCodProducto) = rstDet("COD_PRODUCTO").Value
        ixdbDetalle(i, iintColDesProducto) = rstDet("DES_PRODUCTO").Value
        ixdbDetalle(i, iintColLote) = rstDet("LOTE").Value
        ixdbDetalle(i, iintColVcmto) = "" & rstDet("VEN").Value
        
        If istrCaso = "COMPRA" Then
            ixdbDetalle(i, iintColQFactura) = rstDet("CTD_PRODUCTO_FRAC_ING").Value
            ixdbDetalle(i, iintColQFisica) = rstDet("CTD_PRODUCTO_ING").Value
            ixdbDetalle(i, iintColPrecio) = rstDet("PRC_FINAL_SUBTOTAL").Value
        End If
        If istrCaso = "RECEPCION" Then
            ixdbDetalle(i, iintColQFactura) = rstDet("CTD_PRODUCTO_DOC").Value
            ixdbDetalle(i, iintColQFisica) = rstDet("CTD_PRODUCTO_ING").Value
            ixdbDetalle(i, iintColPrecio) = rstDet("PRC_FINAL_UNIT").Value
        End If
        
        ixdbDetalle(i, iintColQOFactura) = 0
        ixdbDetalle(i, iintColQOFisica) = 0
        ixdbDetalle(i, iintColOPrecio) = 0
        ixdbDetalle(i, iintColProductoProv) = ""

        i = i + 1
        rstDet.MoveNext
    Wend
    grdDetalle.Rebind
    totaliza

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub MostrarCampos()

    'Columnas no visibles
    grdDetalle.Columns(0).Visible = False
    grdDetalle.Columns(iintColQOFactura).Visible = False
    grdDetalle.Columns(iintColQOFisica).Visible = False
    grdDetalle.Columns(iintColOPrecio).Visible = False
    grdDetalle.Columns(iintColProductoProv).Visible = False

    If istrCaso = "RECEPCION" Then
        grdDetalle.Columns(iintColPrecio).Visible = (Mid(lblFormaPago.Caption, 1, 3) = "005")
        
        txtSubtotal.Visible = (Mid(lblFormaPago.Caption, 1, 3) = "005")
        txtImpuesto.Visible = (Mid(lblFormaPago.Caption, 1, 3) = "005")
        txtInafecto.Visible = (Mid(lblFormaPago.Caption, 1, 3) = "005")
        txtTotal.Visible = (Mid(lblFormaPago.Caption, 1, 3) = "005")
        txtRedondeo.Visible = (Mid(lblFormaPago.Caption, 1, 3) = "005")
        
        lblMSubtotal.Visible = (Mid(lblFormaPago.Caption, 1, 3) = "005")
        lblMImpuesto.Visible = (Mid(lblFormaPago.Caption, 1, 3) = "005")
        lblMInafecto.Visible = (Mid(lblFormaPago.Caption, 1, 3) = "005")
        lblMTotal.Visible = (Mid(lblFormaPago.Caption, 1, 3) = "005")
        lblMRedondeo.Visible = (Mid(lblFormaPago.Caption, 1, 3) = "005")
    End If

    If istrCaso = "COMPRA" Then
        grdDetalle.Columns(iintColLote).Visible = False
        grdDetalle.Columns(iintColVcmto).Visible = False
        grdDetalle.Columns(iintColObs).Visible = False
        
        lblObservaciones.Visible = False
        lblForma.Visible = False
        
        txtSubtotal.Visible = (ctlcboDocumento.BoundText <> "BOL")
        txtImpuesto.Visible = (ctlcboDocumento.BoundText <> "BOL")
        txtInafecto.Visible = (ctlcboDocumento.BoundText <> "BOL")
        txtRedondeo.Enabled = (ctlcboDocumento.BoundText <> "BOL")
        
        lblMSubtotal.Visible = (ctlcboDocumento.BoundText <> "BOL")
        lblMImpuesto.Visible = (ctlcboDocumento.BoundText <> "BOL")
        lblMInafecto.Visible = (ctlcboDocumento.BoundText <> "BOL")
        lblMRedondeo.Visible = (ctlcboDocumento.BoundText <> "BOL")
    End If

    'Ajuste de tamaño de las columnas
    grdDetalle.Columns(0).AllowSizing = grdDetalle.Columns(0).Visible
    grdDetalle.Columns(iintColQOFactura).AllowSizing = grdDetalle.Columns(iintColQOFactura).Visible
    grdDetalle.Columns(iintColQOFisica).AllowSizing = grdDetalle.Columns(iintColQOFisica).Visible
    grdDetalle.Columns(iintColOPrecio).AllowSizing = grdDetalle.Columns(iintColOPrecio).Visible
    grdDetalle.Columns(iintColProductoProv).AllowSizing = grdDetalle.Columns(iintColProductoProv).Visible
    grdDetalle.Columns(iintColPrecio).AllowSizing = grdDetalle.Columns(iintColPrecio).Visible
    grdDetalle.Columns(iintColLote).AllowSizing = grdDetalle.Columns(iintColLote).Visible
    grdDetalle.Columns(iintColVcmto).AllowSizing = grdDetalle.Columns(iintColVcmto).Visible
    grdDetalle.Columns(iintColObs).AllowSizing = grdDetalle.Columns(iintColObs).Visible

    grdDetalle.Refresh

End Sub
