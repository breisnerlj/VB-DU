VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_VTA_NotaCredito 
   BorderStyle     =   0  'None
   ClientHeight    =   7305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlGrillaArray grdNC 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4471
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin MSComctlLib.ImageList ilsProcesos 
      Left            =   5040
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_NotaCredito.frx":0000
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_NotaCredito.frx":059A
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_NotaCredito.frx":0B34
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_NotaCredito.frx":10CE
            Key             =   "GrabaDoc"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_NotaCredito.frx":1668
            Key             =   "Usuario"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_NotaCredito.frx":1C02
            Key             =   "Calc"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_NotaCredito.frx":219C
            Key             =   "Aplicacion"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_NotaCredito.frx":2736
            Key             =   "Zoom"
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid70.TDBGrid oldgrdNC 
      Height          =   615
      Left            =   1680
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5760
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   1085
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
      Columns(1).DataField=   "COD_PRODUCTO"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Producto"
      Columns(2).DataField=   "DES_PRODUCTO"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Ctd Und"
      Columns(3).DataField=   "CTD_UNIDADES"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Ctd Frac"
      Columns(4).DataField=   "CTD_FRACCIONES"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Sub Tot"
      Columns(5).DataField=   "IMP_SUBTOTAL"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Ctd Dev"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Prec Dev"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Ctd Fracc"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Precio"
      Columns(9).DataField=   "PRC_UNIT_VTA"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=503"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=423"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1138"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1058"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=4075"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3995"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=1217"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1138"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=1217"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1138"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=1217"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1138"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=1482"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1402"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=1561"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=1482"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=1455"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=1376"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=172,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=63,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=72,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=64,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=65,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=66,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=68,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=67,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=69,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=70,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=71,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=73,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=74,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=16,.parent=63"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=64"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=65"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=67"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=20,.parent=63"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=17,.parent=64"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=18,.parent=65"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=19,.parent=67"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=24,.parent=63"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=21,.parent=64"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=22,.parent=65"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=23,.parent=67"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=63"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=64"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=65"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=67"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=63"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=64"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=65"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=67"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=63"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=64"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=65"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=67"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=63"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=64"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=65"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=67"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=54,.parent=63"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=64"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=65"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=67"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=58,.parent=63"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=64"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=65"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=67"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=62,.parent=63"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=59,.parent=64"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=60,.parent=65"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=61,.parent=67"
      _StyleDefs(76)  =   "Named:id=33:Normal"
      _StyleDefs(77)  =   ":id=33,.parent=0"
      _StyleDefs(78)  =   "Named:id=34:Heading"
      _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(80)  =   ":id=34,.wraptext=-1"
      _StyleDefs(81)  =   "Named:id=35:Footing"
      _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(83)  =   "Named:id=36:Selected"
      _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(85)  =   "Named:id=37:Caption"
      _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(87)  =   "Named:id=38:HighlightRow"
      _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&HFFF2FF&,.fgcolor=&H400040&,.bold=-1,.fontsize=825"
      _StyleDefs(89)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(90)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(91)  =   "Named:id=39:EvenRow"
      _StyleDefs(92)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(93)  =   "Named:id=40:OddRow"
      _StyleDefs(94)  =   ":id=40,.parent=33"
      _StyleDefs(95)  =   "Named:id=41:RecordSelector"
      _StyleDefs(96)  =   ":id=41,.parent=34"
      _StyleDefs(97)  =   "Named:id=42:FilterBar"
      _StyleDefs(98)  =   ":id=42,.parent=33"
   End
   Begin vbp_Ventas.ctlGrilla grdCajas 
      Height          =   2055
      Left            =   0
      TabIndex        =   25
      Top             =   1080
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3625
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   3195
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   7185
      Begin vbp_Ventas.ctlTextBox txtObservacion 
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   1085
         Tipo            =   2
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
      Begin VB.CheckBox chkDevEfectivo 
         Caption         =   "Devolver Efectivo"
         Height          =   375
         Left            =   5760
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1095
      End
      Begin vbp_Ventas.ctlTextBox txtUsuario 
         Height          =   315
         Left            =   3240
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1965
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin vbp_Ventas.ctlTextBox txtNombrePC 
         Height          =   315
         Left            =   960
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1965
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin vbp_Ventas.ctlDataCombo ctlCboConcepto 
         Height          =   315
         Left            =   3360
         TabIndex        =   5
         Top             =   1155
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin vbp_Ventas.ctlTextBox TxtFactura 
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   180
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         Enabled         =   0   'False
         EnabledFoco     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin vbp_Ventas.ctlTextBox TxtRazonSocial 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   495
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
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
      Begin vbp_Ventas.ctlTextBox TxtDireccion 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   825
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
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
      Begin vbp_Ventas.ctlTextBox TxtRuc 
         Height          =   315
         Left            =   5400
         TabIndex        =   3
         Top             =   810
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
      Begin vbp_Ventas.ctlTextBox TxtImpTot 
         Height          =   315
         Left            =   960
         TabIndex        =   9
         Top             =   1545
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Enabled         =   0   'False
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
      Begin vbp_Ventas.ctlTextBox TxtImpReal 
         Height          =   315
         Left            =   4080
         TabIndex        =   10
         Top             =   210
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         ColorDefault    =   -2147483633
         ColorDefault    =   -2147483633
         Enabled         =   0   'False
         Bloqueado       =   -1  'True
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
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   5640
         TabIndex        =   28
         Top             =   2040
         Width           =   1335
      End
      Begin vbp_Ventas.ctlTextBox txtTelefono 
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Top             =   1155
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Tipo            =   2
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
      Begin vbp_Ventas.ctlDataCombo cboSubMotivo 
         Height          =   315
         Left            =   3360
         TabIndex        =   39
         Top             =   1545
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         MatchEntry      =   1
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Motivo"
         Height          =   195
         Left            =   2640
         TabIndex        =   40
         Top             =   1605
         Width           =   480
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Telefono:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   1215
         Width           =   675
      End
      Begin VB.Label Label15 
         Caption         =   "Observacion"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Left            =   2640
         TabIndex        =   26
         Top             =   2025
         Width           =   540
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Caja"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   2025
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   195
         Left            =   2640
         TabIndex        =   17
         Top             =   1215
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total Real  S/."
         Height          =   195
         Left            =   3000
         TabIndex        =   16
         Top             =   285
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total  S/."
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1605
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "R.U.C."
         Height          =   195
         Left            =   4560
         TabIndex        =   14
         Top             =   885
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   885
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Señor(es)"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   555
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   540
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   600
      Left            =   0
      TabIndex        =   29
      Top             =   480
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1058
      ButtonWidth     =   1191
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ilsProcesos"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Key             =   "Grabar"
            Object.ToolTipText     =   "Graba Documento"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Imprimir"
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel"
            Key             =   "Cancel"
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Aplicacion"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Asigna"
            Key             =   "Asigna"
            ImageIndex      =   5
         EndProperty
      EndProperty
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   3720
         TabIndex        =   30
         Top             =   -120
         Width           =   3495
         Begin VB.Label LblFchSysdate 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label LblFactura 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1920
            TabIndex        =   33
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Actual"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   120
            Width           =   945
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Factura"
            Height          =   195
            Left            =   1920
            TabIndex        =   31
            Top             =   120
            Width           =   1035
         End
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "F. Actual"
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
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   780
      End
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
      Index           =   1
      Left            =   0
      TabIndex        =   22
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "Nº"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   158
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frm_VTA_NotaCredito.frx":2CD0
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Anulación de Documentos"
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
      Index           =   0
      Left            =   480
      TabIndex        =   19
      Top             =   120
      Width           =   2760
   End
   Begin VB.Label LblNumNC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFBFA&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   330
      Left            =   4320
      TabIndex        =   18
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frm_VTA_NotaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pOpt As String
Public pCol As String
Dim objDocumento As New clsDocumento
Dim intCant%

Private Sub ctlCboConcepto_Change()
    Set cboSubMotivo.RowSource = objDocumento.ListaMotivoNCDet(ctlCboConcepto.BoundText)
    cboSubMotivo.ListField = "DES"
    cboSubMotivo.BoundColumn = "COD"

End Sub

Private Sub ctlCboConcepto_Click(Area As Integer)
    If ctlCboConcepto.BoundText = "027" Then
        'chkDevEfectivo.Visible = True
        chkDevEfectivo.Value = 1
    ElseIf ctlCboConcepto.BoundText = "028" Then
        chkDevEfectivo.Visible = False
        chkDevEfectivo.Value = 0
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            grdNC.SetFocus
    End Select
End Sub

Private Sub Form_Load()
        setteaFormulario Me
        
    SeteaGrilla
    Setea_Array gxdbNC
        
    chkDevEfectivo.Visible = False
    Frame3.Enabled = False
        
        
    Set ctlCboConcepto.RowSource = objDocumento.ListaMotivoNC
    ctlCboConcepto.ListField = "DES"
    ctlCboConcepto.BoundColumn = "COD"
    'ctlCboConcepto.BoundText = "*"

    LblNumNC.Caption = objDocumento.ListaNumeroDisponible(objUsuario.CodigoEmpresa, objUsuario.NombrePC, objVenta.CodigoDocumentoVenta)
    
End Sub

Sub Setea_Array(ByVal xdbArray As XArrayDB)
    xdbArray.ReDim 0, -1, 0, 12
    'grdNC.Array = xdbArray
    grdNC.Array1 = xdbArray
    grdNC.Rebind
End Sub

Private Sub grdCajas_DblClick()
    If grdCajas.ApproxCount = 0 Then Exit Sub
    txtNombrePC.Text = grdCajas.Columns("COD_MAQUINA").Value
    txtUsuario.Text = grdCajas.Columns("USU_TEC").Value
    grdCajas.Visible = False
    
End Sub

Private Sub grdCajas_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            grdCajas_DblClick
        Case vbKeyEscape
            grdCajas.Visible = False
    End Select
End Sub

Private Sub grdNC_DblClick()
    'If ctlCboConcepto.BoundText = "" Then MsgBox "Seleccione un concepto de Nota de Credito", vbCritical, App.ProductName: ctlCboConcepto.SetFocus: Exit Sub
    If ctlCboConcepto.BoundText = "" Then MsgBox "Seleccione un concepto de Anulación de Documento", vbCritical, App.ProductName: ctlCboConcepto.SetFocus: Exit Sub
    If grdNC.ApproxCount <= 0 Then Exit Sub
    pCol = grdNC.Col
    pOpt = ctlCboConcepto.BoundText
    If ctlCboConcepto.BoundText = "027" Then
         frm_VTA_Tipo_NotaCredito.FraDev.Visible = True
         frm_VTA_Tipo_NotaCredito.FraDscto.Visible = False
         frm_VTA_Tipo_NotaCredito.Show vbModal
    ElseIf ctlCboConcepto.BoundText = "028" Then
         frm_VTA_Tipo_NotaCredito.FraDev.Visible = False
         frm_VTA_Tipo_NotaCredito.FraDscto.Visible = True
         frm_VTA_Tipo_NotaCredito.Show vbModal
    End If
End Sub

Private Sub grdNC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            grdNC_DblClick
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strNumNC As String
    Select Case Button.Key
        Case "Grabar"
            GrabarNc
        Case "Imprimir"
            MsgBox "Opción habilitada en la consulta de Documento", vbExclamation, App.ProductName: Exit Sub
            
            'strNumNC = InputBox("Ingresar Número de la Nota de Crédito", App.ProductName)
            
            'If strNumNC <> "" Then objDocumento.ImprimirDocumento objUsuario.TipoDocNC, strNumNC
            
        Case "Cancel"
            Cancel
        Case "Aplicacion"
            Unload Me
            
        Case "Asigna"
            Asigna
    End Select
End Sub

Sub Cancel()
    If MsgBox("Desea borrar los datos de la Anulación de Documento", vbQuestion + vbYesNo) = vbYes Then
        txtDireccion.Text = ""
        TxtRazonSocial.Text = ""
        TxtRuc.Text = ""
        txtNombrePC.Text = ""
        txtUsuario.Text = ""
        TxtFactura.Text = "": txtDireccion.Text = ""
        TxtImpReal.Text = "": TxtImpTot.Text = "": txtTelefono.Text = ""
        txtObservacion.Text = ""
        cboSubMotivo.BoundText = ""
        ctlCboConcepto.BoundText = "*"
        Setea_Array gxdbNC
        grdNC.Columns(7).FooterText = ""
        grdNC.Columns(8).FooterText = ""
        frm_VTA_Concepto_NotaCredito.Show vbModal
        If TxtRazonSocial.Enabled = True Then TxtRazonSocial.Focus
    End If
End Sub

Private Sub TxtDireccion_KeyPress(KeyAscii As Integer)
    txtDireccion.Tipo = AlfaNumerico
End Sub

Private Sub TxtFactura_KeyPress(KeyAscii As Integer)
    TxtFactura.Tipo = Entero
End Sub

Private Sub TxtImpReal_KeyPress(KeyAscii As Integer)
    TxtImpReal.Tipo = Real
End Sub

Private Sub TxtImpTot_KeyPress(KeyAscii As Integer)
    TxtImpTot.Tipo = Real
End Sub

Private Sub txtRazonSocial_KeyPress(KeyAscii As Integer)
    TxtRazonSocial.Tipo = Mayusculas
End Sub

Private Sub TxtRuc_KeyPress(KeyAscii As Integer)
    TxtRuc.Tipo = Entero
End Sub

Sub GrabarNc()
On Error GoTo Handle
    Dim oTipDoc As OraParamArray
    Dim oNumDoc As OraParamArray
    Dim oTipDocCo As OraParamArray
    Dim oNumDocCo As OraParamArray
    Dim auxTipDoc As OraParamArray
    Dim auxNumDoc As OraParamArray
    Dim RetPromoMensaje As String
    Dim varMsgDoc As Variant
    Dim varMsgDocCo As Variant
    Dim objImpresion As New clsImpresiones
    Dim strOldCodMaquina As String
    Dim strOldCodUsuario As String
        
    Dim UltDocEmitido As oraDynaset
    Dim strUltDocEmi As String
    Dim i As Integer
    
    If txtNombrePC.Text = "" Or txtUsuario.Text = "" Then MsgBox "Asignar una caja", vbCritical, App.ProductName: ctlCboConcepto.SetFocus: Exit Sub
        
    If TxtRazonSocial.Text = "" Then MsgBox "Debe indicar el nombre o la razón social del cliente", vbCritical, App.ProductName: TxtRazonSocial.Focus: Exit Sub
    If txtDireccion.Text = "" Then MsgBox "Debe indicar la dirección del cliente", vbCritical, App.ProductName: txtDireccion.Focus: Exit Sub
    If TxtRuc.Text = "" Then MsgBox "Debe indicar el D.N.I o la R.U.C. del cliente", vbCritical, App.ProductName: TxtRuc.Focus: Exit Sub
    If ctlCboConcepto.BoundText = "" Then MsgBox "Seleccione un concepto de anulación", vbCritical, App.ProductName: ctlCboConcepto.SetFocus: Exit Sub
    If cboSubMotivo.BoundText = "" Then MsgBox "Debe seleccionar un sub concepto de Anulación.", vbOKOnly + vbExclamation, "Advertencia": Exit Sub
    If grdNC.ApproxCount <= 0 Then Exit Sub

    objVenta.CodConcepto = objDocumento.ConceptoCredito
    objVenta.CodSubConcepto = ctlCboConcepto.BoundText
    objVenta.DevEfectivoNC = chkDevEfectivo.Value

    objVenta.Ruc = TxtRuc.Text
    objVenta.NumeroDocumentoID = TxtRuc.Text

    objVenta.RazonSocial = TxtRazonSocial.Text
    objVenta.DesAuxCliDirecc = txtDireccion.Text
    objVenta.NombreCliente = TxtRazonSocial.Text
    ''''''
    objVenta.DireccionClienteDLV = txtDireccion.Text
    objVenta.NombreClienteDLV = txtDireccion.Text
    objVenta.DesAuxCliTlf = Trim(txtTelefono.Text)
    objVenta.DesAuxCliNombre = Trim(TxtRazonSocial.Text)
    objVenta.DireccionCliente = Trim(txtDireccion.Text)

    strOldCodMaquina = objUsuario.NombrePC
    strOldCodUsuario = objUsuario.Codigo

    objUsuario.NombrePC = txtNombrePC.Text
    objUsuario.Codigo = txtUsuario.Text
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Autor : Arturo Escate
    'Fecha : 10/11/2009
    'Proposito: Esto es para validar si necesita autorizacion previa
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Dim ObjValidacion As New clsAprobacion
    Dim strNumeroSolicitud As String
    Dim strAccion As String
    Dim strMensaje As String
    Dim strCodigoAutorizacion As String
    Dim srtCodigoAUTH As String
    Dim strStore As String
    srtCodigoAUTH = ""
valida:
'TxtImpTot.text

    strStore = ObjValidacion.Solicita("1", strAccion, strMensaje, srtCodigoAUTH, objUsuario.CodigoLocal, objUsuario.CodigoLiquidacion, objVenta.CodigoCliente, objVenta.CodigoDocumentoVenta, LblNumNC.Caption, objVenta.CodDocRef, objVenta.NumDocRef, objVenta.Totales(0), "", "", objUsuario.Codigo, ctlCboConcepto.Text & " " & cboSubMotivo.BoundText & ":<br>" & Trim(txtObservacion.Text) & "<br>" & "Cliente:" & TxtRazonSocial.Text & "<br>Dirección:" & txtDireccion.Text & "<br>RUC:" & TxtRuc.Text, strCodigoAutorizacion, LblFactura.Caption, TxtImpTot.Text, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    If Not strStore = "" Then
        MsgBox strStore, vbCritical, App.ProductName
        Exit Sub
    Else
        Select Case strAccion
            Case 0
                    MsgBox strMensaje, vbInformation, App.ProductName
            Case 1
                   MsgBox strMensaje, vbCritical, App.ProductName
                   Exit Sub
            Case 2
                   MsgBox strMensaje, vbInformation, App.ProductName
                   Exit Sub
            Case 3
                If MsgBox(strMensaje & Chr(13) & "¿Desea ingresar el codigo de autorización?", vbYesNo + vbInformation, App.ProductName) = vbYes Then
                    srtCodigoAUTH = frmAprobacion.carga
                    If Not srtCodigoAUTH = "" Then
                        GoTo valida
                        Exit Sub
                    End If
                   Exit Sub
                Else
                    Exit Sub
                End If
            Case Else
                   MsgBox "no esta implementado", vbInformation, App.ProductName
                   Exit Sub
        End Select
    End If

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    objVenta.MotivoNotaCredito = cboSubMotivo.BoundText
    objVenta.Observacion = Trim(txtObservacion.Text)
    objVenta.GrabarDoc gclsOracle.ODataBase, oTipDoc, oNumDoc, oTipDocCo, oNumDocCo, RetPromoMensaje, True
    
    objUsuario.NombrePC = strOldCodMaquina
    objUsuario.Codigo = strOldCodUsuario
    
    
    objVenta.OrdenaDoc gclsOracle.ODataBase, oTipDoc, oNumDoc, oTipDocCo, oNumDocCo, auxTipDoc, auxNumDoc
    
    For i = 0 To oTipDoc.ArraySize - 1
        If oTipDoc.get_Value(i) <> "" Then
        varMsgDoc = varMsgDoc & oTipDoc.get_Value(i) & " " & oNumDoc.get_Value(i) & Chr(13)
        End If
    Next i

    For i = 0 To oTipDocCo.ArraySize - 1
        If oTipDocCo.get_Value(i) <> "" Then
        varMsgDocCo = varMsgDocCo & oTipDocCo.get_Value(i) & " " & oNumDocCo.get_Value(i) & Chr(13)
        End If
    Next i
    
    MsgBox "Se realizo la transacción satisfactoriamente  - " & Chr(13) & varMsgDoc & _
            "Por convenio - " & varMsgDocCo, vbInformation + vbOKOnly, "Graba"
            
    strUltDocEmi = ""
    
    For i = 0 To auxTipDoc.ArraySize - 1
        Set UltDocEmitido = objDocumento.UltDocEmitido(objUsuario.CodigoEmpresa, objUsuario.NombrePC)
        If Not UltDocEmitido.EOF Then strUltDocEmi = UltDocEmitido("COD_TIPO_DOCUMENTO").Value
        If auxTipDoc.get_Value(i) <> "" Then
            If auxTipDoc.get_Value(i) <> strUltDocEmi Then
                MsgBox "Sirvase poner la palanca de la impresora" + Chr(13) + _
                        "en posicion de " & auxTipDoc.get_Value(i)
            End If
            objDocumento.ImprimirDocumento auxTipDoc.get_Value(i), auxNumDoc.get_Value(i)
            objDocumento.GrabaUltDocEmitido objUsuario.CodigoEmpresa, objUsuario.NombrePC, auxTipDoc.get_Value(i)
        End If
    Next i
    subNuevo
Exit Sub
Handle:
     objUsuario.NombrePC = strOldCodMaquina
     objUsuario.Codigo = strOldCodUsuario

    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Sub subNuevo()
    TxtFactura.Text = "": txtDireccion.Text = ""
    TxtImpReal.Text = "": TxtImpTot.Text = ""
    ctlCboConcepto.BoundText = "*"
    TxtRazonSocial.Text = ""
    TxtRuc.Text = ""
    txtTelefono.Text = ""
    txtObservacion.Text = ""
    cboSubMotivo.BoundText = ""
    txtUsuario.Text = ""
    txtNombrePC.Text = ""
    Setea_Array gxdbNC
    grdNC.Columns(7).FooterText = ""
    grdNC.Columns(8).FooterText = ""

    Set objVenta = Nothing
    frm_VTA_Concepto_NotaCredito.Show vbModal
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant

'    grdNC.AllowUpdate = False
'    grdNC False
'    grdNC.AllowColSelect = False
'    grdNC.Style.VerticalAlignment = dbgVertCenter
'    grdNC.RowHeight = 1.2 * grdNC.RowHeight
'    grdNC.Columns("DES_PRODUCTO").WrapText = True
'
'    grdNC.MarqueeStyle = dbgHighlightRowRaiseCell
'    Dim i%
'    For i = 0 To grdNC.Columns.Count - 1
'        grdNC.Columns(i).AllowFocus = False
'    Next i
'    grdNC.Columns(6).AllowFocus = True
'    grdNC.Columns(7).AllowFocus = True
'    grdNC.Columns(8).Visible = False
'    grdNC.RecordSelectors = False
'
    '    grdNC.Columns(7).NumberFormat = "#0." & String(gintDecTot, "0")
    
    arrCampos = Array("", _
                      "", _
                      "", _
                      "", _
                      "", _
                      "", _
                      "", _
                      "", _
                      "", _
                      "")
    
    arrCaption = Array("#", _
                        "Código", _
                        "Descripción", _
                        "Un", _
                        "Fr", _
                        "Prec", _
                        "STot", _
                        "Ctd.Dev", _
                        "Devol", _
                        "Real")
                        
    arrAncho = Array(260, _
                     700, _
                     2800, _
                     420, _
                     420, _
                     700, _
                     520, _
                     700, _
                     700, _
                     900)
                     
    arrAlineacion = Array(dbgLeft, _
                          dbgCenter, _
                          dbgLeft, _
                          dbgRight, _
                          dbgRight, _
                          dbgRight, _
                          dbgRight, _
                          dbgRight, _
                          dbgRight, _
                          dbgRight)
    
    grdNC.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdNC.Columns(3).NumberFormat = "###"
    grdNC.Columns(4).NumberFormat = "###"
    grdNC.Columns(5).NumberFormat = "#,##0.00"
    grdNC.Columns(8).NumberFormat = "#,##0.00"
    grdNC.Columns(6).Visible = False
    grdNC.ColumnFooter = True
    grdNC.Columns(1).FooterAlignment = dbgLeft
    grdNC.Columns(1).FooterText = "Total"
    
    '----------
    
    arrCampos = Array("COD_MAQUINA", "FCH_INICIO", _
                      "USU_TEC", "NOM_TEC", _
                      "FLG_ESTADO_CAJA", "COD_LIQUIDACION")
                      
    arrCaption = Array("Maquina", "Fch Inicio", _
                       "Código", "Nombre Depen", _
                       "Estado", "Liquidación")
    
    arrAncho = Array(900, 1000, _
                     600, 2000, _
                     850, 1100)
    
    arrAlineacion = Array(dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft)
    
    grdCajas.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    grdCajas.ColumnFooter = True
    grdCajas.Columns(1).FooterAlignment = dbgLeft
    grdCajas.Columns(1).FooterText = "<ESC> Salir"
    grdCajas.Columns(3).FooterAlignment = dbgLeft
    grdCajas.Columns(3).FooterText = "<ENTER> Seleccionar"

    grdCajas.Visible = False
    
End Sub

Private Sub Asigna()

On Error GoTo Handle

            Set grdCajas.DataSource = objLiquidacion.ListaCajasPrecerradas(objUsuario.CodigoEmpresa, _
                                                                           objUsuario.CodigoLocal, _
                                                                           LblFchSysdate.Caption)
            grdCajas.Visible = True
            grdCajas.SetFocus
            
Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub
