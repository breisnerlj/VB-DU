VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frm_VTA_Solicitud_Ajuste 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Ajustes"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11145
   Icon            =   "frm_VTA_Solicitud_Ajuste.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtObservacion 
      Height          =   855
      Left            =   120
      MaxLength       =   399
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   6360
      Width           =   10935
   End
   Begin vbp_Ventas.ctlGrillaArray grdProducto 
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2143
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin ORADCLibCtl.ORADC oradcTipo 
      Height          =   255
      Left            =   4200
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
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
   Begin TrueDBGrid70.TDBDropDown drdTipo 
      Bindings        =   "frm_VTA_Solicitud_Ajuste.frx":000C
      Height          =   1455
      Left            =   4200
      TabIndex        =   12
      Top             =   3840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2566
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Cod."
      Columns(0).DataField=   "COD_TIP_AJUSTE"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripción"
      Columns(1).DataField=   "DES_TIP_AJUSTE"
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
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
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
   Begin ORADCLibCtl.ORADC oradcMotivo 
      Height          =   255
      Left            =   6480
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
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
   Begin vbp_Ventas.ctlTextBox txtProducto 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
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
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "..."
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
      Left            =   4800
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OptionButton optCodigo 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton optDescri 
      Caption         =   "Descripción"
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   1320
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1215
   End
   Begin vbp_Ventas.ctlTextBox txtSolicitud 
      Height          =   315
      Left            =   1080
      TabIndex        =   8
      Top             =   780
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      ColorDefault    =   -2147483633
      ColorDefault    =   -2147483633
      Tipo            =   3
      Alignment       =   2
      Enabled         =   0   'False
      EnabledFoco     =   0   'False
      Bloqueado       =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TrueDBGrid70.TDBDropDown drdMotivo 
      Bindings        =   "frm_VTA_Solicitud_Ajuste.frx":0024
      Height          =   1455
      Left            =   6480
      TabIndex        =   6
      Top             =   3840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2566
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Cod."
      Columns(0).DataField=   "COD_MOTIVO_AJUSTE"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripción"
      Columns(1).DataField=   "DES_MOTIVO_AJUSTE"
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
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
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
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   1058
      ModoBotones     =   6
   End
   Begin vbp_Ventas.ctlGrillaArray grdSolicitud 
      Height          =   3975
      Left            =   0
      TabIndex        =   1
      Top             =   2055
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7011
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin VB.Label lblObservaciones 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   6120
      Width           =   1065
   End
   Begin VB.Label lblSolicitud 
      AutoSize        =   -1  'True
      Caption         =   "# Solicitud :"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   840
   End
End
Attribute VB_Name = "frm_VTA_Solicitud_Ajuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objAjustes As New clsAjuste
Dim objProducto As New clsProducto
Dim odynTipo, odynMotivo As oraDynaset
Dim ixdbSolicitud As New XArrayDB
Dim ixdbProducto As New XArrayDB
Dim rstProducto, rstValor As oraDynaset
Dim varBookMark As Variant
Dim strProducto As String
Dim vValAlta, vValBaja As Double

Private Sub cmdBuscar_Click()
Dim objProducto As New clsProducto

On Error GoTo Control
        
    Set rstProducto = objAjustes.ListaProductos(objUsuario.CodigoLocal, strProducto)
    Set objProducto = Nothing
    
    If rstProducto.RecordCount > 0 Then
       ixdbProducto.LoadRows rstProducto.GetRows
       If ixdbProducto(0, 0) = "-1" Then
           MsgBox "El producto no se encuentra en la maestro de productos: " & TxtProducto.Text, vbCritical, App.ProductName
           TxtProducto.Text = ""
           Exit Sub
       End If
    End If
    
    grdProducto.Rebind
    grdProducto.MoveFirst
    grdProducto_DblClick
   
   Exit Sub

Control:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error : " & Err.Number
End Sub

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
On Error GoTo Control
    Select Case Index
        Case 1
            Grabar
        Case 2
            Unload Me
    End Select

   Exit Sub

Control:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error : " & Err.Number
End Sub

Private Sub Form_Load()

    ctlToolBar1.Buttons(11).Visible = False
    IniciaArray
    SetSolicitud
    txtSolicitud.Text = objAjustes.DevNroSolicitud

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objAjustes = Nothing
    Set objProducto = Nothing
    Set odynTipo = Nothing
    Set odynMotivo = Nothing
    Set ixdbSolicitud = Nothing
    Set ixdbProducto = Nothing
    Set rstProducto = Nothing
    Set rstValor = Nothing
End Sub

Private Sub grdProducto_DblClick()
Dim IntRow As Integer
Dim strMensaje As String
Dim blnAdd  As Boolean
Dim btnActualizar As Boolean
Dim inTf As Integer
Dim i As Integer

    strMensaje = ""
    btnActualizar = False
    IntRow = 0
    
    Screen.MousePointer = vbHourglass
        
    If Existe(grdProducto.Columns(0).Value) Then
        strMensaje = strMensaje & CStr("" & grdProducto.Columns(0).Value) & " - " & CStr("" & grdProducto.Columns(1).Value) & Chr(13)
    Else
        btnActualizar = True
        ixdbSolicitud.AppendRows 1
        ixdbSolicitud(ixdbSolicitud.UpperBound(1), 0) = grdProducto.Columns(0).Value
        ixdbSolicitud(ixdbSolicitud.UpperBound(1), 1) = grdProducto.Columns(1).Value
        ixdbSolicitud(ixdbSolicitud.UpperBound(1), 2) = grdProducto.Columns(2).Value
        ixdbSolicitud(ixdbSolicitud.UpperBound(1), 5) = 0
        ixdbSolicitud(ixdbSolicitud.UpperBound(1), 8) = grdProducto.Columns(3).Value

        If grdProducto.Columns(4).Value = "0" Then
            grdSolicitud.Columns(5).AllowFocus = False
            ixdbSolicitud(ixdbSolicitud.UpperBound(1), 5) = 0
        Else
            grdSolicitud.Columns(6).AllowFocus = True
            ixdbSolicitud(ixdbSolicitud.UpperBound(1), 6) = 0
        End If

        ixdbSolicitud(ixdbSolicitud.UpperBound(1), 7) = "0.00"

    End If
    
    grdProducto.MoveNext
    Screen.MousePointer = vbDefault

    If btnActualizar = True Then
        grdSolicitud.Rebind
        grdSolicitud.MoveLast
        'grdSolicitud.SetFocus
        TxtProducto.Text = ""
        grdProducto.Limpiar
    End If

    If strMensaje <> "" Then
        Me.Refresh
        MsgBox "El siguiente producto ya se encuentra en la lista: " & Chr(13) & strMensaje, vbCritical, App.ProductName
        grdSolicitud.Bookmark = inTf
        TxtProducto.SetFocus
        TxtProducto.Text = ""
        grdProducto.Limpiar
        Exit Sub
    End If

   Exit Sub

Control:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error : " & Err.Number
End Sub

Private Sub grdProducto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then grdProducto_DblClick
End Sub

Private Sub grdSolicitud_AfterColUpdate(ByVal ColIndex As Integer)
On Error GoTo Control

    grdSolicitud.MoveNext
    grdSolicitud.MovePrevious
    
    Select Case ColIndex
        Case 5
            Calcula_TotalValor
        Case 6
            Calcula_TotalValor
    End Select
    
  Exit Sub
Control:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error : " & Err.Number
End Sub

Private Sub grdSolicitud_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
On Error GoTo Control

    Select Case ColIndex
        Case 5
            If Trim(grdSolicitud.Columns(ColIndex).Value) = "" Then
                grdSolicitud.Columns(ColIndex).Value = 0
                If grdSolicitud.Columns(6).Value = 0 Then
                    grdSolicitud.Columns(7).Value = 0
                  Else
                    grdSolicitud.Columns(7).Value = objAjustes.DevValor(objUsuario.CodigoLocal, grdSolicitud.Columns(0).Value, _
                                                    Val(grdSolicitud.Columns(5).Value), Val(grdSolicitud.Columns(6).Value))
                End If
                
                ElseIf Not IsNumeric(grdSolicitud.Columns(ColIndex).Value) Then
                    MsgBox "La cantidad ingresada no es válido", vbCritical, "Error"
                    Cancel = True
                    grdSolicitud.col = ColIndex
                    Else
                    grdSolicitud.Columns(7).Value = objAjustes.DevValor(objUsuario.CodigoLocal, grdSolicitud.Columns(0).Value, _
                                                    Val(grdSolicitud.Columns(5).Value), Val(grdSolicitud.Columns(6).Value))
            End If

        Case 6
            If Trim(grdSolicitud.Columns(ColIndex).Value) = "" Then
                grdSolicitud.Columns(ColIndex).Value = 0
                    If grdSolicitud.Columns(5).Value = 0 Then
                       grdSolicitud.Columns(7).Value = 0
                      Else
                        grdSolicitud.Columns(7).Value = objAjustes.DevValor(objUsuario.CodigoLocal, grdSolicitud.Columns(0).Value, _
                                                                            Val(grdSolicitud.Columns(5).Value), Val(grdSolicitud.Columns(6).Value))
                    End If
                ElseIf Not IsNumeric(grdSolicitud.Columns(ColIndex).Value) Then
                    MsgBox "La cantidad ingresada no es válido", vbCritical, "Error"
                    Cancel = True
                    grdSolicitud.col = ColIndex
                    Else
                    grdSolicitud.Columns(7).Value = objAjustes.DevValor(objUsuario.CodigoLocal, grdSolicitud.Columns(0).Value, _
                                                                        Val(grdSolicitud.Columns(5).Value), Val(grdSolicitud.Columns(6).Value))
            End If
   End Select
   grdSolicitud.SetFocus
   Exit Sub

Control:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error : " & Err.Number
End Sub

Private Sub grdSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo CtrlErr
    
    If grdSolicitud.ApproxCount < 1 Then Exit Sub
    
    If Not grdSolicitud.EditActive Then
    
        If KeyCode = vbKeyDelete Then
            If MsgBox("¿Desea eliminar el registro de la lista?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Item") = vbYes Then
                grdSolicitud.Delete
                grdSolicitud.Refresh
                TxtProducto.SetFocus
            End If
        End If
    End If
    
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub grdSolicitud_LostFocus()
    grdSolicitud.Refresh
End Sub

Private Sub grdSolicitud_RegistroSeleccionado(ByVal DatoColumna0 As String)
    If grdSolicitud.ApproxCount <= 0 Then
       Exit Sub
    Else
        CargaTipos
        
        If grdSolicitud.Columns(8).Value = 0 Then
           grdSolicitud.Columns(6).AllowSizing = False
           grdSolicitud.Columns(6).AllowFocus = False
        Else
           grdSolicitud.Columns(6).AllowSizing = True
           grdSolicitud.Columns(6).AllowFocus = True
        End If
    End If
End Sub

Private Sub txtObservacion_GotFocus()
    txtObservacion.BackColor = TxtProducto.ColorFoco
End Sub

Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtObservacion_LostFocus()
    txtObservacion.BackColor = TxtProducto.ColorDefault
End Sub

Private Sub TxtProducto_KeyPress(KeyAscii As Integer)
    On Error GoTo handle
    
    If KeyAscii = vbKeyReturn Then

        If Len(Trim(TxtProducto.Text)) < 3 Then
            MsgBox "use por lo menos 3 caracteres", vbExclamation, "Aviso"
            Exit Sub
        End If

        Dim frm As New frm_ADM_ProductoDatos
        frm.Dato = Trim(TxtProducto.Text)
        frm.Show vbModal

        If frm.Salida(1) <> "" Then
            grdSolicitud.SetFocus
        End If

        strProducto = frm.Salida(1)
        
        Set frm = Nothing

        Call cmdBuscar_Click

    End If

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub CargaTipos()

    drdTipo.RowHeight = 0
    drdTipo.RowHeight = drdTipo.RowHeight * 1.7
    
    drdTipo.DataField = "COD_TIP_AJUSTE"
    drdTipo.ListField = "DES_TIP_AJUSTE"
    drdTipo.AllowRowSizing = False
    drdTipo.AllowColMove = False
    drdTipo.EmptyRows = True
        
    Set oradcTipo.Recordset = objAjustes.ListaTipoAju
    Set odynTipo = oradcTipo.Recordset.Clone
    
    odynTipo.MoveFirst
    
    While Not odynTipo.EOF
        psub_Grilla_Traslate grdSolicitud, 3, odynTipo("COD_TIP_AJUSTE").Value, odynTipo("DES_TIP_AJUSTE").Value
        odynTipo.MoveNext
    Wend
        
    drdTipo.Height = 0
    drdTipo.Height = 3 * drdTipo.Height + drdTipo.RowHeight * (IIf(oradcTipo.Recordset.RecordCount > 8, 8, oradcTipo.Recordset.RecordCount))
    
    drdTipo.Appearance = dbgFlat
    drdTipo.Columns(0).BackColor = RGB(240, 240, 240)

    Set objAjustes = Nothing
    CargaMotivos

End Sub

Private Sub CargaMotivos()
Dim strTipoDev As String

    drdMotivo.RowHeight = 0
    drdMotivo.RowHeight = drdMotivo.RowHeight * 2

    drdMotivo.DataField = "COD_MOTIVO_AJUSTE"
    drdMotivo.ListField = "DES_MOTIVO_AJUSTE"
    drdMotivo.AllowRowSizing = False
    drdMotivo.AllowColMove = False
    drdMotivo.EmptyRows = True

    If IsNull(grdSolicitud.Columns(3).Value) Then
        strTipoDev = "-"
    ElseIf IsEmpty(grdSolicitud.Columns(3).Value) Or grdSolicitud.Columns(3).Value = "" Then
        strTipoDev = "-"
    Else
        strTipoDev = grdSolicitud.Columns(3).Value
    End If
    
    Set oradcMotivo.Recordset = objAjustes.ListaMotivoTipoAju(strTipoDev, objUsuario.CodigoLocal)
    Set odynMotivo = oradcMotivo.Recordset.Clone

    odynMotivo.MoveFirst

    While Not odynMotivo.EOF
        psub_Grilla_Traslate grdSolicitud, 4, odynMotivo("COD_MOTIVO_AJUSTE").Value, odynMotivo("DES_MOTIVO_AJUSTE").Value
        odynMotivo.MoveNext
    Wend

    drdMotivo.Height = 0
    drdMotivo.Height = 3 * drdMotivo.Height + drdMotivo.RowHeight * (IIf(oradcMotivo.Recordset.RecordCount > 8, 8, oradcMotivo.Recordset.RecordCount))

    drdMotivo.Appearance = dbgFlat
    drdMotivo.Columns(0).BackColor = RGB(240, 240, 240)

    Set objAjustes = Nothing
    Set odynMotivo = Nothing

End Sub

Private Sub IniciaArray()
On Error GoTo Control

    ixdbSolicitud.ReDim 0, -1, 0, 8
    grdSolicitud.Array1 = ixdbSolicitud
    grdSolicitud.Rebind
   
    ixdbProducto.ReDim 0, -1, 0, 8
    grdProducto.Array1 = ixdbProducto
    grdProducto.Rebind
   
   Exit Sub

Control:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error : " & Err.Number
End Sub

Private Sub Grabar()
Dim larrTipo() As String
Dim larrMotipo() As String
Dim larrProducto() As String
Dim larrCtdUnd() As Double
Dim larrCtdFrac() As Double
Dim i As Integer
Dim strMensaje As String
Dim strNumDocumento As String
Dim arrDocs As Variant

On Error GoTo Control

    ReDim larrTipo(0 To 0)
    ReDim larrMotipo(0 To 0)
    ReDim larrProducto(0 To 0)
    ReDim larrCtdUnd(0 To 0)
    ReDim larrCtdFrac(0 To 0)

    If ixdbSolicitud.UpperBound(1) <> "-1" Then

        For i = ixdbSolicitud.LowerBound(1) To ixdbSolicitud.UpperBound(1)
            
            larrProducto(UBound(larrProducto)) = ixdbSolicitud(i, 0)
    
            larrTipo(UBound(larrTipo)) = ixdbSolicitud(i, 3)
                If ixdbSolicitud(i, 3) = "" Then
                   MsgBox "Debe indicar el Tipo de Ajuste ", vbCritical, App.ProductName
                   grdSolicitud.SetFocus
                   Exit Sub
                End If
            
            larrMotipo(UBound(larrMotipo)) = ixdbSolicitud(i, 4)
                If ixdbSolicitud(i, 4) = "" Then
                   MsgBox "Debe indicar el Motivo del Ajuste", vbCritical, App.ProductName
                   grdSolicitud.SetFocus
                   Exit Sub
                End If
                
            larrCtdUnd(UBound(larrCtdUnd)) = ixdbSolicitud(i, 5)
            larrCtdFrac(UBound(larrCtdFrac)) = ixdbSolicitud(i, 6)

             If (ixdbSolicitud(i, 5) = "" Or ixdbSolicitud(i, 5) = "0") And (ixdbSolicitud(i, 6) = "" Or ixdbSolicitud(i, 6) = "0") Then
                   MsgBox "Debe ingresar las Unidades y/o Fracciones para el ajuste", vbCritical, App.ProductName
                   grdSolicitud.SetFocus
                   Exit Sub
              End If
    
            If ixdbSolicitud(i, 7) = "0.00" Or ixdbSolicitud(i, 7) = "0" Then
                  MsgBox "El siguente del producto " & ixdbSolicitud(i, 0) & " su estado es " & ixdbSolicitud(i, 2) & Chr(13) _
                          & "y su Valor es 0", vbCritical, App.ProductName
            End If
    
            If Mid(grdSolicitud.Columns(7).FooterText, 5, 4) < objVenta.ParametroValor("MTOMINAJU") Then
                  MsgBox "El Total de la solicitud quedara en negativo ", vbCritical, App.ProductName
            End If
            
            ReDim Preserve larrProducto(UBound(larrProducto) + 1)
            ReDim Preserve larrTipo(UBound(larrTipo) + 1)
            ReDim Preserve larrMotipo(UBound(larrMotipo) + 1)
            ReDim Preserve larrCtdUnd(UBound(larrCtdUnd) + 1)
            ReDim Preserve larrCtdFrac(UBound(larrCtdFrac) + 1)
        Next i
    
        ReDim Preserve larrProducto(UBound(larrProducto) - 1)
        ReDim Preserve larrTipo(UBound(larrTipo) - 1)
        ReDim Preserve larrMotipo(UBound(larrMotipo) - 1)
        ReDim Preserve larrCtdUnd(UBound(larrCtdUnd) - 1)
        ReDim Preserve larrCtdFrac(UBound(larrCtdFrac) - 1)

    Else
        MsgBox "Debe ingresar el o los producto(s) para el ajuste", vbCritical, "Error"
        Exit Sub
    End If

    If MsgBox("¿Seguro(a) de Guardar la Solicitud de Ajuste?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbNo Then
        Exit Sub
    End If

    strNumDocumento = ""

    strMensaje = objAjustes.GrabaSolicitud(strNumDocumento, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, objUsuario.Codigo, _
                                           txtObservacion.Text, larrTipo, larrMotipo, larrProducto, larrCtdUnd, larrCtdFrac)

    If strMensaje = "" Then
        If InStr(strNumDocumento, "|") = 0 Then strNumDocumento = strNumDocumento & "|"
        arrDocs = Split(strNumDocumento, "|")
        MsgBox "Se grabo satisfactoriamente la solicitud N°" & vbCrLf & Join(arrDocs, vbCrLf), vbExclamation, App.ProductName
        Unload Me
    Else
        MsgBox strMensaje, vbCritical, App.ProductName
    End If

    Exit Sub

Control:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error : " & Err.Number
End Sub

Private Sub SetSolicitud()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim i As Integer

    arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO", "COD_ESTADO", "COD_TIPO", "COD_MOTIVO", "CTD_UND", "CTD_FRAC", "VALOR", "")
    'arrCampos = Array("", "", "", "", "", "", "")
    arrCaption = Array("Código", "Descripción", "Estado", "Tipo", "Motivo", "Ctd Und", "Ctd Frac", "Valor", "")
    arrAncho = Array(700, 3000, 600, 2000, 2000, 800, 800, 1000, 800)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgCenter, dbgLeft, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgCenter)

    With grdSolicitud
        .FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        .AllowUpdate = True
        .HeadLines = 2
        .RowHeight = 0
        .RowHeight = grdSolicitud.RowHeight * 2
        .ColumnFooter = True
        .Columns(8).Visible = False

        .Columns(0).AllowFocus = False
        .Columns(1).AllowFocus = False
        .Columns(2).AllowFocus = False
        .Columns(3).AllowFocus = True
        .Columns(4).AllowFocus = True
        .Columns(5).AllowFocus = True
        .Columns(6).AllowFocus = True
        .Columns(7).AllowFocus = False
        .Columns(8).AllowFocus = False
        .Columns(8).AllowSizing = False
        
        .Columns(5).EditMask = "###"
        .Columns(6).EditMask = "###"

        .Columns(3).DropDown = drdTipo
        .Columns(3).AutoCompletion = True
        .Columns(3).AutoDropDown = True
        .Columns(3).DropDownList = True

        .Columns(4).DropDown = drdMotivo
        .Columns(4).AutoCompletion = True
        .Columns(4).AutoDropDown = True
        .Columns(4).DropDownList = True
        
        .Columns(4).FooterText = "Total"
        .Columns(4).FooterDivider = False
        .Columns(5).FooterDivider = False
        .Columns(4).FooterAlignment = dbgCenter
        .Columns(5).FooterText = "->"
        
        .Columns(7).NumberFormat = "#,###,##0.00"
        .Columns(7).FooterBackColor = &HC0E0FF
        .Columns(7).FooterText = "S/." & " " & Format(0, "###,###0.00")
        .Rebind
    End With

'****************************************************************************************************************
'****************************************************************************************************************
'****************************************************************************************************************

    arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO", "COD_ESTADO", "FLG_FRACCIONAMIENTO", "", "")
    arrCaption = Array("Código", "Descripción", "Estado", "Fracciona", "", "")
    arrAncho = Array(800, 4500, 1000, 1000, 1000, 1000)
    arrAlineacion = Array(dbgCenter, dbgLeft, dbgCenter, dbgCenter, dbgCenter, dbgCenter)
    
    With grdProducto
         .FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
         .AllowUpdate = False
         For i = 2 To 3
             .Columns(i).Visible = False
         Next i
    
        For i = 2 To 3
            .Columns(i).AllowSizing = False
        Next i
    End With
End Sub

Private Function Existe(ByVal strProducto As String) As Boolean
Dim i As Integer
On Error GoTo Control
     i = ixdbSolicitud.Find(0, 0, strProducto, XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
     If i >= 0 Then
        Existe = True
     Else
        Existe = False
     End If
   Exit Function
Control:
    Existe = False
End Function

Private Sub Calcula_TotalValor()
Dim i As Integer

    vValAlta = 0
    vValBaja = 0
    
    For i = ixdbSolicitud.LowerBound(1) To ixdbSolicitud.UpperBound(1)
    
         If ixdbSolicitud(i, 3) = "JCA" Then
            vValAlta = vValAlta + ixdbSolicitud(i, 7)
         ElseIf ixdbSolicitud(i, 3) = "JCB" Then
                vValBaja = vValBaja + ixdbSolicitud(i, 7)
         End If

    Next i

    grdSolicitud.Columns(7).FooterText = "S/." & " " & Format((vValAlta - vValBaja), "###,###0.00")

End Sub
