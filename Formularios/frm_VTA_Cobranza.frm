VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frm_VTA_Cobranza 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_VTA_Cobranza.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCodMoneda 
      Height          =   375
      Left            =   5760
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin vbp_Ventas.ctlGrilla grdEfectivo 
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1720
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin TrueDBGrid70.TDBGrid grdDocumentos 
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4683
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "TD"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Fecha"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Documento"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "M"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Importe"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "###,##0.00"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Saldo"
      Columns(5).DataField=   ""
      Columns(5).NumberFormat=   "###,##0.00"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "I. a Pagar"
      Columns(6).DataField=   ""
      Columns(6).NumberFormat=   "###,##0.00"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=714"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=609"
      Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8193"
      Splits(0)._ColumnProps(6)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1931"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1826"
      Splits(0)._ColumnProps(11)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(13)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Width=2831"
      Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2725"
      Splits(0)._ColumnProps(18)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=8193"
      Splits(0)._ColumnProps(20)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(22)=   "Column(3).Width=688"
      Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=582"
      Splits(0)._ColumnProps(25)=   "Column(3).AllowSizing=0"
      Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(27)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(29)=   "Column(4).Width=1588"
      Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=1482"
      Splits(0)._ColumnProps(32)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=8194"
      Splits(0)._ColumnProps(34)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(36)=   "Column(5).Width=1482"
      Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=1376"
      Splits(0)._ColumnProps(39)=   "Column(5).AllowSizing=0"
      Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=8194"
      Splits(0)._ColumnProps(41)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(42)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(43)=   "Column(6).Width=1852"
      Splits(0)._ColumnProps(44)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(6)._WidthInPix=1746"
      Splits(0)._ColumnProps(46)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(47)=   "Column(6).FetchStyle=1"
      Splits(0)._ColumnProps(48)=   "Column(6).Order=7"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
      DirectionAfterEnter=   2
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2,.locked=0"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1,.locked=-1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1,.locked=-1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1,.bgcolor=&H80000018&"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17,.bgcolor=&H80000018&"
      _StyleDefs(64)  =   "Named:id=33:Normal"
      _StyleDefs(65)  =   ":id=33,.parent=0"
      _StyleDefs(66)  =   "Named:id=34:Heading"
      _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   ":id=34,.wraptext=-1"
      _StyleDefs(69)  =   "Named:id=35:Footing"
      _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(71)  =   "Named:id=36:Selected"
      _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=37:Caption"
      _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(75)  =   "Named:id=38:HighlightRow"
      _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=39:EvenRow"
      _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(79)  =   "Named:id=40:OddRow"
      _StyleDefs(80)  =   ":id=40,.parent=33"
      _StyleDefs(81)  =   "Named:id=41:RecordSelector"
      _StyleDefs(82)  =   ":id=41,.parent=34"
      _StyleDefs(83)  =   "Named:id=42:FilterBar"
      _StyleDefs(84)  =   ":id=42,.parent=33"
   End
   Begin VB.TextBox txtCodigoCliente 
      Height          =   375
      Left            =   6480
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtNumDocumento 
      Height          =   360
      Left            =   6360
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdBuscarCliente 
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   4920
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      Picture         =   "frm_VTA_Cobranza.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5145
      Picture         =   "frm_VTA_Cobranza.frx":0894
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   1095
   End
   Begin vbp_Ventas.ctlDataCombo ctlCboTipCliente 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      MatchEntry      =   1
   End
   Begin vbp_Ventas.ctlTextBox TxtCliente 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   480
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
   Begin VB.Label lblImporteDeuda 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      TabIndex        =   22
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label lblImporteSaldo 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      TabIndex        =   21
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Importe S/. : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Saldo S/. : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   360
      TabIndex        =   16
      Top             =   480
      Width           =   795
   End
   Begin VB.Label LblCliente 
      BackColor       =   &H00DBFBFA&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   5655
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
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1215
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
      Left            =   3780
      TabIndex        =   13
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
      Index           =   11
      Left            =   5490
      TabIndex        =   12
      Top             =   6900
      Width           =   390
   End
   Begin VB.Label Label5 
      Caption         =   "Abono S/. : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Documentos pendientes de pago :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   360
      TabIndex        =   9
      Top             =   2295
      Width           =   3570
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
      Index           =   8
      Left            =   60
      TabIndex        =   8
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   6600
      Picture         =   "frm_VTA_Cobranza.frx":0E1E
      Top             =   120
      Width           =   240
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
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
End
Attribute VB_Name = "frm_VTA_Cobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objCliente As New clsCliente
Dim objCobranza As New clsCobranza

Dim xCobranza As New XArrayDB

Private Sub cmdAceptar_Click()
Dim strTipoDocumento, strNumeroDocumento, strMonedaVenta, _
    strImporteVenta, strImporteAbono, _
    strMonedaAbono, strImporteAbonoMoneda As String

Dim f As Integer
f = 0
While f < xCobranza.UpperBound(1)

    If Val(xCobranza(f, 6)) > 0 Then
        strTipoDocumento = strTipoDocumento & CStr(xCobranza(f, 0)) & "|"
        strNumeroDocumento = strNumeroDocumento & xCobranza(f, 2) & "|"
        strMonedaVenta = strMonedaVenta & xCobranza(f, 3) & "|"
        strImporteVenta = CStr(strImporteVenta) & CStr(xCobranza(f, 4)) & "|"
        If TxtCodMoneda.Text = "002" Then
            strImporteAbono = strImporteAbono & (Val(xCobranza(f, 6)) * objUsuario.TipoCambio) & "|"
        Else
            strImporteAbono = strImporteAbono & (Val(xCobranza(f, 6)) * objUsuario.TipoCambio) & "|"
        End If
        strMonedaAbono = strMonedaAbono & TxtCodMoneda.Text & "|"
        strImporteAbonoMoneda = strImporteAbonoMoneda & xCobranza(f, 6) & "|"
    End If
        f = f + 1
Wend
    
    objCobranza.Grabar objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, _
                       txtNumDocumento.Text, strTipoDocumento, strNumeroDocumento, _
                       strMonedaVenta, strImporteVenta, strImporteAbono, _
                        strMonedaAbono, strImporteAbonoMoneda, objUsuario.Codigo
End Sub

Private Sub cmdBuscarCliente_Click()
frm_VTA_Cliente_Bus.Show vbModal
''TxtCliente.Text = "" & frm_VTA_Cliente_Bus.Out_CodigoCliente
''LblCliente.Caption = "" & frm_VTA_Cliente_Bus.Out_NombreCliente
''txtNumDocumento.Text = "" & frm_VTA_Cliente_Bus.Out_NumeroId
''txtCodigoCliente.Text = "" & frm_VTA_Cliente_Bus.Out_CodigoCliente

txtCliente.Text = "" & objVenta.Out_CodigoCliente
LblCliente.Caption = "" & objVenta.Out_NombreCliente
txtNumDocumento.Text = "" & objVenta.Out_NumeroId
txtCodigoCliente.Text = "" & objVenta.Out_CodigoCliente



CargaDatos
''''''''Comentado por el nuevo cambio de cliente
''''''''Dim pstrFlgCli As String
''''''''Dim pstrDatoCliente  As String
''''''''    If ctlCboTipCliente.BoundText = "*" Then MsgBox "Seleccione un tipo de Cliente", vbCritical, Caption: Exit Sub
''''''''    pstrFlgCli = ctlCboTipCliente.BoundText
''''''''    pstrDatoCliente = Trim(TxtCliente.Text)
''''''''    Set frm_VTA_ClienteDatos.GrdBusCliente.DataSource = objCliente.ListaClientesGen(pstrDatoCliente, pstrFlgCli)
''''''''    frm_VTA_ClienteDatos.Pantalla = 1
''''''''    frm_VTA_ClienteDatos.Show vbModal
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
    'objVenta.CancelarVenta
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control
    psub_KeyDownAplicacion KeyCode, Shift
    Select Case KeyCode
        Case vbKeyReturn
            
        Case vbKeyF1
            cmdBuscarCliente.SetFocus
        Case vbKeyF2
            grdDocumentos.SetFocus
        Case vbKeyEscape
            cmdCancelar_Click
    End Select

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_Load()
SetteaFormulario Me
xCobranza.ReDim 0, -1, 0, 6
    Set ctlCboTipCliente.RowSource = objCliente.ListaTipo
    ctlCboTipCliente.ListField = "DES"
    ctlCboTipCliente.BoundColumn = "COD"
    ctlCboTipCliente.BoundText = "*"
    Dim ObjFormaPago As New clsFormaPago
    Dim arrCampo, arrCaption, arrWith, ArrAlineamiento As Variant
    arrCampo = Array("COD_MONEDA", "DES_HIJO")
    arrCaption = Array("Codigo", "Moneda")
    arrWith = Array(800, 2500)
    ArrAlineamiento = Array(vbAlignLeft, vbAlignLeft)
    grdEfectivo.FormatoGrilla arrCampo, arrCaption, arrWith, ArrAlineamiento
    Set grdEfectivo.DataSource = ObjFormaPago.ListaHijo("001")
    Set ObjFormaPago = Nothing
    'lblTipoCambio = objUsuario.TipoCambio
''''    Dim arrCampo, arrCaption, arrWidth, arrAlign As Variant
''''    arrCampo = Array("TIP_DOC", "NRO_DOC", "MONEDA", "IMPORTE", "SALDO")
''''    arrCaption = Array("TD", "Número", "M", "Importe", "Saldo")
''''    arrWidth = Array(500, 1500, 300, 1500, 1500)
''''    arrAlign = Array(dbgCenter, dbgCenter, dbgCenter, dbgLeft, dbgLeft)
''''    grdDocumentos.FormatoGrilla arrCampo, arrCaption, arrWidth, arrAlign
End Sub

Sub totaliza()
Dim i As Integer
Dim Saldo As Double
Dim Pago As Double

While i < xCobranza.UpperBound(1)
    If TxtCodMoneda.Text = objUsuario.Parametros("COD_MONEDA")(0, 2) Then
        Saldo = Saldo + Val(xCobranza(i, 5))
        Pago = Pago + Val(xCobranza(i, 6))
    Else
        Saldo = Saldo + (Val(xCobranza(i, 5)) * objUsuario.TipoCambio)
        Pago = Pago + (Val(xCobranza(i, 6)) * objUsuario.TipoCambio)
    End If
    i = i + 1
Wend
lblImporte.Caption = Format(Pago, "###,##0.00")
lblImporteDeuda.Caption = Format(Saldo, "###,##0.00")
lblImporteSaldo.Caption = Format(Val(Saldo) - Val(Pago), "###,##0.00")

End Sub

Private Sub grdDocumentos_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'grdDocumentos.Update
Select Case ColIndex
Case 6
    If Val(grdDocumentos.Columns(6)) > Val(grdDocumentos.Columns(5)) Then
        Cancel = 1
        MsgBox "El abono no puede ser mayor que la deuda", vbCritical
        grdDocumentos.Col = 6
    Else
        grdDocumentos.Columns(6) = Val(grdDocumentos.Columns(6))
    End If
    'grdDocumentos.Rebind
End Select
End Sub

Private Sub grdDocumentos_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
    If Val(grdDocumentos.Columns(6)) > 0 Then CellStyle.Font.Bold = True
End Sub

Private Sub grdDocumentos_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)
   If Val(grdDocumentos.Columns(6)) > 0 Then RowStyle.Font.Italic = True
End Sub

Private Sub grdDocumentos_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeySpace
        frm_VTA_CobranzaMonto.Show vbModal
End Select
End Sub

Sub CargaDatos()
    Dim rsDatos As oraDynaset
    Dim i As Integer
    Set rsDatos = objCobranza.Lista(objUsuario.CodigoEmpresa, txtNumDocumento.Text)
    
    While Not rsDatos.EOF
        xCobranza.AppendRows
        xCobranza(i, 0) = rsDatos("TIP_DOC").Value
        xCobranza(i, 1) = rsDatos("FEC_EMI").Value
        xCobranza(i, 2) = rsDatos("NRO_DOC").Value
        xCobranza(i, 3) = rsDatos("MONEDA").Value
        xCobranza(i, 4) = rsDatos("IMPORTE").Value
        xCobranza(i, 5) = rsDatos("SALDO").Value
        xCobranza(i, 6) = "0.00"
        i = i + 1
        rsDatos.MoveNext
    Wend
    grdDocumentos.Array = xCobranza
    grdDocumentos.Rebind
    'Set grdDocumentos.DataSource = objCobranza.Lista(objUsuario.CodigoEmpresa, txtNumDocumento.Text)
End Sub

Private Sub grdDocumentos_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii, "."
End Sub

Private Sub grdDocumentos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    totaliza
End Sub


Private Sub grdEfectivo_RegistroSeleccionado(ByVal DatoColumna0 As String)
    If xCobranza.UpperBound(1) = -1 Then Exit Sub
    If MsgBox("Esta a punto de cambiar la moneda esto elimiminara cualquier dato ingresado", vbYesNo, App.ProductName) = vbNo Then Exit Sub
    TxtCodMoneda.Text = grdEfectivo.Columns(0).Value
    Dim j As Integer
    While j < xCobranza.UpperBound(1)
        xCobranza(j, 6) = "0.00"
        j = j + 1
    Wend
    'grdDocumentos.Rebind
    
End Sub
