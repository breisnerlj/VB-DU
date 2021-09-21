VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_VTA_RepKardex 
   BorderStyle     =   0  'None
   Caption         =   "Movimiento de Kardex"
   ClientHeight    =   7020
   ClientLeft      =   2520
   ClientTop       =   -150
   ClientWidth     =   7230
   ForeColor       =   &H00000000&
   Icon            =   "frm_VTA_RepKardex.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Ordenar por Fecha"
      Height          =   735
      Left            =   5400
      TabIndex        =   40
      Top             =   2400
      Width           =   1695
      Begin VB.OptionButton Option1 
         Caption         =   "Ascendente"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Descendente"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   1455
      End
   End
   Begin vbp_Ventas.ctlGrilla grdKardex 
      Height          =   3135
      Left            =   0
      TabIndex        =   21
      Top             =   3480
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   5530
      MenuPopUp       =   0   'False
      Resalte         =   0   'False
   End
   Begin vbp_Ventas.ctlTextBox TxtProducto 
      Height          =   330
      Left            =   1200
      TabIndex        =   20
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
   Begin VB.Frame Frame5 
      Caption         =   "Movimiento"
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
      Height          =   735
      Left            =   3240
      TabIndex        =   11
      Top             =   600
      Width           =   3950
      Begin vbp_Ventas.ctlDataCombo ctlCboMov 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   280
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         MatchEntry      =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Destino"
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
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   3255
      Begin vbp_Ventas.ctlDataCombo ctlCboOrigDest 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   280
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         MatchEntry      =   1
      End
   End
   Begin TrueDBGrid70.TDBGrid grdStock 
      Bindings        =   "frm_VTA_RepKardex.frx":000C
      Height          =   2175
      Left            =   7200
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1320
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3836
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
      Splits(0).DividerColor=   12632256
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
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=97,.bold=0,.fontsize=825,.italic=0"
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
   Begin MSComctlLib.Toolbar TblProd 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   1111
      ButtonWidth     =   1191
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "IlsImagen"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin VB.Frame Frame2 
         Caption         =   "Fechas"
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
         Height          =   735
         Left            =   3240
         TabIndex        =   15
         Top             =   0
         Width           =   3950
         Begin MSComCtl2.DTPicker dtpFechaI 
            Height          =   315
            Left            =   480
            TabIndex        =   16
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   16384001
            CurrentDate     =   37950
         End
         Begin MSComCtl2.DTPicker dtpFechaF 
            Height          =   315
            Left            =   2160
            TabIndex        =   17
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   16384001
            CurrentDate     =   37950
         End
         Begin VB.Label Label9 
            Caption         =   "Al"
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
            Height          =   255
            Left            =   1920
            TabIndex        =   13
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            Caption         =   "Del"
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
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   375
         End
      End
   End
   Begin MSComctlLib.ImageList IlsImagen 
      Left            =   1080
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_RepKardex.frx":0025
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_RepKardex.frx":05BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_RepKardex.frx":0B59
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_RepKardex.frx":10F3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMaximo 
      AutoSize        =   -1  'True
      Caption         =   "C. SAP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   240
      Index           =   2
      Left            =   5400
      TabIndex        =   45
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblCodSap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
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
      Height          =   195
      Left            =   6315
      TabIndex        =   44
      Top             =   1580
      Width           =   825
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   240
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   600
   End
   Begin VB.Line Line6 
      X1              =   4140
      X2              =   4140
      Y1              =   6600
      Y2              =   6840
   End
   Begin VB.Line Line5 
      X1              =   5955
      X2              =   5955
      Y1              =   6600
      Y2              =   6840
   End
   Begin VB.Line Line4 
      X1              =   5040
      X2              =   5040
      Y1              =   6600
      Y2              =   6840
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   3240
      X2              =   3240
      Y1              =   6600
      Y2              =   6840
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6855
      X2              =   6855
      Y1              =   6600
      Y2              =   6840
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   3240
      X2              =   6840
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label lblIngreso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   4140
      TabIndex        =   39
      Top             =   6600
      Width           =   900
   End
   Begin VB.Label lblFinal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   5970
      TabIndex        =   38
      Top             =   6600
      Width           =   900
   End
   Begin VB.Label lblSalida 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   5055
      TabIndex        =   37
      Top             =   6600
      Width           =   900
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "registro(S)"
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
      Left            =   255
      TabIndex        =   35
      Top             =   6720
      Width           =   885
   End
   Begin VB.Label lblNumRegistro 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   120
      TabIndex        =   34
      Top             =   6690
      Width           =   135
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "FECHA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2400
      TabIndex        =   33
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "ORI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1920
      TabIndex        =   32
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "#DOC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   960
      TabIndex        =   31
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "MOV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   480
      TabIndex        =   30
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "TD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "FINAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   5940
      TabIndex        =   28
      Top             =   3240
      Width           =   900
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "SALIDA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   5040
      TabIndex        =   27
      Top             =   3240
      Width           =   900
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "INGRESO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   4140
      TabIndex        =   26
      Top             =   3240
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "INICIAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   25
      Top             =   3240
      Width           =   900
   End
   Begin VB.Label lblMaximo 
      AutoSize        =   -1  'True
      Caption         =   "Maximo"
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
      Index           =   1
      Left            =   5160
      TabIndex        =   24
      Top             =   4680
      Width           =   810
   End
   Begin VB.Label lblMaximo 
      AutoSize        =   -1  'True
      Caption         =   "Máximo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   240
      Index           =   0
      Left            =   4200
      TabIndex        =   23
      Top             =   4680
      Width           =   810
   End
   Begin VB.Label lblStock1 
      AutoSize        =   -1  'True
      Caption         =   "Stock"
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
      Height          =   240
      Left            =   4440
      TabIndex        =   22
      Top             =   1560
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ACT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6120
      TabIndex        =   8
      Top             =   2040
      Width           =   975
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLab 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   2400
      Width           =   4035
   End
   Begin VB.Label Label7 
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblLinea 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   2760
      Width           =   4035
   End
   Begin VB.Label Label8 
      Caption         =   "Linea"
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
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lblCod_Producto 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label LblProducto 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Producto :"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label lblStock 
      AutoSize        =   -1  'True
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   240
      Left            =   3600
      TabIndex        =   0
      Top             =   1560
      Width           =   600
   End
   Begin VB.Label lblInicial 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   3240
      TabIndex        =   36
      Top             =   6600
      Width           =   900
   End
End
Attribute VB_Name = "frm_VTA_RepKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objKardex As New clsKardex
Dim objProducto As New clsProducto
Public pstrCodProd As String

Private Sub ctlTextBox1_KeyPress(KeyAscii As Integer)

End Sub

Private Sub Form_Load()
On Error GoTo handle
    Me.top = 0
    Me.left = 0
    setteaFormulario Me
    SeteaGrilla
    Set ctlCboMov.RowSource = objKardex.ListaMovimientos(objUsuario.CodigoLocal)
    ctlCboMov.ListField = "DES"
    ctlCboMov.BoundColumn = "COD"
    ctlCboMov.BoundText = "*"
        
    Set ctlCboOrigDest.RowSource = objKardex.ListaOrigenDestino(objUsuario.CodigoLocal)
    ctlCboOrigDest.ListField = "DES"
    ctlCboOrigDest.BoundColumn = "COD"
    ctlCboOrigDest.BoundText = "*"
    
    dtpFechaI.Value = Format(objUsuario.sysdate, "dd/mm/yyyy")
    dtpFechaF.Value = Format(objUsuario.sysdate, "dd/mm/yyyy")
    
    lblNumRegistro.Caption = "0"
    
    lblFinal.Caption = ""
    lblInicial.Caption = ""
    lblIngreso.Caption = ""
    lblSalida.Caption = ""


    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub grdKardex_DblClickRegistro(ByVal DatoColumna0 As String)

    If grdKardex.ApproxCount = 0 Then Exit Sub
    If grdKardex.Columns("COD_TIPODOC").Value = objUsuario.TipoDocBol Or grdKardex.Columns("COD_TIPODOC").Value = objUsuario.TipoDocFac Or grdKardex.Columns("COD_TIPODOC").Value = objVenta.TipoDocTKB Or grdKardex.Columns("COD_TIPODOC").Value = objVenta.TipoDocTKF Then
        frm_ADM_PreviewDoc.datos objUsuario.CodigoEmpresa, _
                    objUsuario.CodigoLocal, _
                    grdKardex.Columns("COD_TIPODOC").Value, _
                    grdKardex.Columns("NUM_DOCUMENTO").Value, _
                    frm_VTA_ProductoDatos.pCodProd, ""
    End If
                    
End Sub

Private Sub grdKardex_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)

''''If Condition = 0 Then
''''    Select Case Col
''''        Case 6, 7
''''            CellStyle.BackColor = &HC0C0FF
''''        Case 8, 9
''''            CellStyle.BackColor = &HC0E0FF
''''        Case 10, 11
''''            CellStyle.BackColor = &HC0FFFF
''''        Case 12, 13
''''            CellStyle.BackColor = &HC0FFC0
''''    End Select
''''    CellStyle.ForeColor = vbBlack
''''    CellStyle.Font.Bold = True
''''
''''End If
''''
''''If Condition = 2 Or Condition = 3 Then
''''    CellStyle.ForeColor = vbYellow
''''    CellStyle.Font.Bold = True
''''
''''End If

End Sub


Private Sub TblProd_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Nuevo"
            Nuevo
        Case "Buscar"
            Buscar
        Case "Imprimir"
            ImprimirKardex
        Case "Salir"
            Unload Me
    End Select
End Sub

Sub Nuevo()
On Error GoTo handle
    grdKardex.Limpiar
    ctlCboMov.BoundText = "*": ctlCboOrigDest.BoundText = "*"
    lblCod_Producto.Caption = "": LblProducto.Caption = ""
    lblEstado.Caption = "": lblLab.Caption = ""
    lblLinea.Caption = "": lblStock1.Caption = ""
    lblMaximo(1).Caption = ""
    TxtProducto.Text = "": TxtProducto.SetFocus
    frm_VTA_ProductoDatos.pCodProd = ""

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Public Sub Buscar()
On Error GoTo handle
Dim rsKardex As oraDynaset
'Dim intCtdFraccionamiento As Integer
Dim intCtdFraccionamiento As Long

'Dim intCtdInicial As Integer
'Dim intCtdInicialFrac As Integer

Dim intCtdInicial As Long
Dim intCtdInicialFrac As Long


'Dim intCtdFinal As Integer
'Dim intCtdFinalFrac As Integer
Dim intCtdFinal As Long
Dim intCtdFinalFrac As Long

'Dim intCtdIngreso As Integer
'Dim intCtdIngresoFrac As Integer

Dim intCtdIngreso As Long
Dim intCtdIngresoFrac As Long


Dim intCtdSalida As Long
Dim intCtdSalidaFrac As Long

'Dim intCtdSalida As Integer
'Dim intCtdSalidaFrac As Integer

'Dim intInicial As Integer
'Dim intFinal As Integer

Dim intInicial As Long
Dim intFinal As Long

Dim intIngreso As Long
Dim intSalida As Long

'Dim intIngreso As Integer
'Dim intSalida As Integer

Dim strOrden As String
If Option1(0).Value = True Then strOrden = " DESC "
If Option1(1).Value = True Then strOrden = " ASC "

        If Trim(TxtProducto.Text) = "" Then MsgBox "Ingrese el Producto a Consultar", vbCritical, Caption: TxtProducto.selection: Exit Sub
        Set rsKardex = objKardex.Lista(frm_VTA_ProductoDatos.pCodProd, _
                                                   objUsuario.CodigoLocal, _
                                                   ctlCboMov.BoundText, _
                                                   ctlCboOrigDest.BoundText, _
                                                   CStr(Format(dtpFechaI.Value, "dd/mm/yyyy")), _
                                                   CStr(Format(dtpFechaF.Value, "dd/mm/yyyy")), strOrden)
                                                   
    Set grdKardex.DataSource = rsKardex
                                                   
    lblStock1.Caption = objKardex.BuscaStock(frm_VTA_ProductoDatos.pCodProd, objUsuario.CodigoLocal, "1")
    lblMaximo(1).Caption = objProducto.DevMaxProducto(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, frm_VTA_ProductoDatos.pCodProd)
    lblStock1.Visible = True
    
    lblNumRegistro.Caption = "0"
    lblFinal.Caption = ""
    lblInicial.Caption = ""
    lblIngreso.Caption = ""
    lblSalida.Caption = ""
    
                                                           
    intCtdInicial = 0
    intCtdInicialFrac = 0
    intCtdFinal = 0
    intCtdFinalFrac = 0
                                                            
    intCtdIngreso = 0
    intCtdIngresoFrac = 0
    intCtdSalida = 0
    intCtdSalidaFrac = 0
                                                            
                                                           
    intInicial = 0
    intFinal = 0

    intIngreso = 0
    intSalida = 0
                                                           
                                                           
    If grdKardex.ApproxCount > 0 Then
        lblNumRegistro.Caption = rsKardex.RecordCount
        grdKardex.SetFocus
    End If
    
    
    If rsKardex.RecordCount > 0 Then
    
        With rsKardex
    
        If Not .EOF Then
    
            intCtdFraccionamiento = IIf(IsNull(rsKardex("CTD_FRACCIONAMIENTO").Value), 0, rsKardex("CTD_FRACCIONAMIENTO").Value)
            .MoveFirst
            intCtdFinal = IIf(IsNull(rsKardex("CTD_FINAL").Value), 0, rsKardex("CTD_FINAL").Value)
            intCtdFinalFrac = IIf(IsNull(rsKardex("CTD_FINAL_FRAC").Value), 0, rsKardex("CTD_FINAL_FRAC").Value)
            intFinal = intCtdFinal * intCtdFraccionamiento + intCtdFinalFrac
    
    
            While Not .EOF
                intCtdIngreso = intCtdIngreso + IIf(IsNull(rsKardex("CTD_INGRESO").Value), 0, rsKardex("CTD_INGRESO").Value)
                intCtdIngresoFrac = intCtdIngresoFrac + IIf(IsNull(rsKardex("CTD_INGRESO_FRAC").Value), 0, rsKardex("CTD_INGRESO_FRAC").Value)
                intCtdSalida = intCtdSalida + IIf(IsNull(rsKardex("CTD_SALIDA").Value), 0, rsKardex("CTD_SALIDA").Value)
                intCtdSalidaFrac = intCtdSalidaFrac + IIf(IsNull(rsKardex("CTD_SALIDA_FRAC").Value), 0, rsKardex("CTD_SALIDA_FRAC").Value)
            
                intCtdInicial = IIf(IsNull(rsKardex("CTD_INICIAL").Value), 0, rsKardex("CTD_INICIAL").Value)
                intCtdInicialFrac = IIf(IsNull(rsKardex("CTD_INICIAL_FRAC").Value), 0, rsKardex("CTD_INICIAL_FRAC").Value)
                intInicial = intCtdInicial * intCtdFraccionamiento + intCtdInicialFrac
                .MoveNext
            Wend
    
            intIngreso = intCtdIngreso * intCtdFraccionamiento + intCtdIngresoFrac
            intSalida = intCtdSalida * intCtdFraccionamiento + intCtdSalidaFrac
    
        End If
    
    
    
    
        End With
    
        lblFinal.Caption = intFinal \ intCtdFraccionamiento & IIf(intFinal Mod intCtdFraccionamiento = 0, "", "F" & intFinal Mod intCtdFraccionamiento)
        lblInicial.Caption = intInicial \ intCtdFraccionamiento & IIf(intInicial Mod intCtdFraccionamiento = 0, "", "F" & intInicial Mod intCtdFraccionamiento)
        
        lblIngreso.Caption = intIngreso \ intCtdFraccionamiento & IIf(intIngreso Mod intCtdFraccionamiento = 0, "", "F" & intIngreso Mod intCtdFraccionamiento)
        lblSalida.Caption = intSalida \ intCtdFraccionamiento & IIf(intSalida Mod intCtdFraccionamiento = 0, "", "F" & intSalida Mod intCtdFraccionamiento)
    
    End If
                                                            
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName

End Sub

Private Sub TxtProducto_KeyPress(KeyAscii As Integer)
On Error GoTo handle
    TxtProducto.Tipo = AlfaNumerico
    If KeyAscii = 13 Then
'''        lblNumRegistro.Caption = "0"
'''        lblFinal.Caption = ""
'''        lblInicial.Caption = ""
'''        lblIngreso.Caption = ""
'''        lblSalida.Caption = ""
    
        pstrCodProd = Trim(TxtProducto.Text)
        frm_VTA_ProductoDatos.Show
        
        
        
        
    End If

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub

Private Sub SeteaGrilla()
On Error GoTo handle
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  Dim Columna As TrueDBGrid70.Column

  
'    arrCampos = Array("SEC_KARDEX", "FCH_MOVIMIENTO", _
'                      "TIP_MOVIMIENTO", "COD_TIPODOC", _
'                      "NUM_DOCUMENTO", "CTD_INICIAL", _
'                      "CTD_INICIAL_FRAC", "CTD_INGRESO", _
'                      "CTD_INGRESO_FRAC", "CTD_SALIDA", _
'                      "CTD_SALIDA_FRAC", "CTD_FINAL", _
'                      "CTD_FINAL_FRAC", "COD_LOCAL_REF", _
'                      "COD_USUARIO", "NOM_USUARIO")
                      
                      
    arrCampos = Array("SEC_KARDEX", "COD_TIPODOC", _
                      "TIP_MOVIMIENTO", "NUM_DOCUMENTO", _
                      "COD_LOCAL_REF", "FCH_MOVIMIENTO", _
                      "CTD_INICIAL", "CTD_INICIAL_FRAC", _
                      "CTD_INGRESO", "CTD_INGRESO_FRAC", _
                      "CTD_SALIDA", "CTD_SALIDA_FRAC", _
                      "CTD_FINAL", "CTD_FINAL_FRAC", _
                      "COD_USUARIO", "NOM_USUARIO")
                      
                      
    arrCaption = Array("SEC", "DC", _
                       "MV", "NUM DOC", _
                       "ORI", "FCH DOC", _
                       "Und", "Fra", _
                       "Und", "Fra", _
                       "Und", "Fra", _
                       "Und", "Fra", _
                       "Codigo", "Usuario")
    
    arrAncho = Array(410, 410, _
                     500, 1050, _
                     430, 800, _
                     460, 460, _
                     460, 460, _
                     460, 460, _
                     460, 460, _
                     0, 0)
                     
    arrAlineacion = Array(dbgRight, dbgCenter, _
                          dbgCenter, dbgLeft, _
                          dbgCenter, dbgLeft, _
                          dbgRight, dbgLeft, _
                          dbgRight, dbgLeft, _
                          dbgRight, dbgLeft, _
                          dbgRight, dbgLeft, _
                          dbgRight, dbgRight)
                          
    grdKardex.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    grdKardex.ColumnHeaders = False
    
    grdKardex.Columns(14).Visible = False
    grdKardex.Columns(15).Visible = False
    
    grdKardex.Columns(0).Visible = False
    
    grdKardex.Columns(6).FetchStyle = True
    grdKardex.Columns(7).FetchStyle = True
    grdKardex.Columns(8).FetchStyle = True
    grdKardex.Columns(9).FetchStyle = True
    grdKardex.Columns(10).FetchStyle = True
    grdKardex.Columns(11).FetchStyle = True
    grdKardex.Columns(12).FetchStyle = True
    grdKardex.Columns(13).FetchStyle = True
    
    grdKardex.Columns(6).DividerStyle = dbgNoDividers
    grdKardex.Columns(8).DividerStyle = dbgNoDividers
    grdKardex.Columns(10).DividerStyle = dbgNoDividers
    grdKardex.Columns(12).DividerStyle = dbgNoDividers
    
    
    For Each Columna In grdKardex.Columns
        Columna.AllowSizing = False
    
    Next
    

    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
    
End Sub



Sub ImprimirKardex()
Dim Impresora As Printer
Dim intNumLineas As Integer
Dim intNumPaginas As Integer
Dim intNumLotes As Integer
Dim intMaxPaginas As Integer
Dim intMaxLineas As Integer
Const MAX_PAGINAS_X_LOTES = 20
Const MAX_LINEAS_X_PAGINAS = 60 '59---75 lineas
Dim intLineas, intPaginas, intLotes As Integer
Dim intK As Integer
Dim strCadenaKardex As String
Dim strSQL As String
Dim intRespuesta As Integer
Dim lngI As Long
Dim lngIF As Long
Dim lngS As Long
Dim lngSF As Long
Dim rstImpresion As oraDynaset
    On Error GoTo ErrorImpresora
    Set rstImpresion = grdKardex.DataSource
    If (rstImpresion.RecordCount = 0) Then
        MsgBox "NO EXISTEN REGISTROS A IMPRIMIR", vbInformation, "IMPRESION"
        Exit Sub
    Else
        rstImpresion.MoveFirst
    End If
   'rstImpresion.MoveFirst
  ' rstImpresion.MoveLast
   Set Impresora = Printer
   Impresora.Height = 280 * 56.7
   Impresora.Width = 320 * 56.7
   Impresora.FontName = "Draft 17cpi"
   ''''ESTOS DATOS HAY QUE CALCULARLOS
   intNumLineas = rstImpresion.RecordCount
   rstImpresion.MoveFirst
   intNumPaginas = (intNumLineas \ MAX_LINEAS_X_PAGINAS) + IIf((intNumLineas Mod MAX_LINEAS_X_PAGINAS = 0), 0, 1)
   intNumLotes = (intNumPaginas \ MAX_PAGINAS_X_LOTES) + IIf((intNumPaginas Mod MAX_PAGINAS_X_LOTES = 0), 0, 1)
   '''''''''''''''''''''''''''''''''''
   For intLotes = 1 To intNumLotes
       If intNumPaginas - ((intLotes - 1) * MAX_PAGINAS_X_LOTES) > MAX_PAGINAS_X_LOTES Then
          intMaxPaginas = MAX_PAGINAS_X_LOTES
       Else
          intMaxPaginas = intNumPaginas - ((intLotes - 1) * MAX_PAGINAS_X_LOTES)
       End If
       For intPaginas = 1 To intMaxPaginas
           If intNumLineas - ((intPaginas - 1) * MAX_LINEAS_X_PAGINAS) > MAX_LINEAS_X_PAGINAS Then
              intMaxLineas = MAX_LINEAS_X_PAGINAS
           Else
              intMaxLineas = intNumLineas - ((intPaginas - 1) * MAX_LINEAS_X_PAGINAS)
           End If
           'Impresora.FontName = "Draft 12cpi"
           'Impresora.Print IIf(chkFecha.Value = Checked, String(14, "_") & Space$(10), "") & "BTL" & gblstrCodigoBTL & "  Usuario : " & glbstrUsuario & " - " & strNombreUsuario(glbstrUsuario)
           'Impresora.Print String(14, "_") & Space$(10), "") & "BTL" & gblstrCodigoBTL & "  Usuario : " & glbstrUsuario & " - " & strNombreUsuario(glbstrUsuario)
           'Impresora.FontName = "Draft 17cpi"
           
           Impresora.Print "Desde: " & Format(dtpFechaI.Value, "dd/MM/yyyy") & Space(100) & Date
           Impresora.Print "Hasta: " & Format(dtpFechaF.Value, "dd/MM/yyyy") & Space(100) & Time
           'Impresora.Print IIf(chkFecha.Value = Checked, "| " & Format(dtpFinal.Value, "dd/MM/yyyy") & Format(dtpHoraFinal.Value, " HH:mm") & " |", "") & Space(IIf(chkFecha.Value = Checked, 114, 133) - Len(CStr(Time))) & CStr(Time)
           'Impresora.FontName = "Draft 12cpi"
           Impresora.FontSize = 15
           Impresora.FontBold = True
           Impresora.Print "                                           REPORTE DE KARDEX EN CANTIDADES               "
           Impresora.FontSize = 12
           Impresora.FontBold = False
           Impresora.Print "Producto    : " & lblCod_Producto & " - " & LblProducto
           Impresora.Print "Destino     : " & ctlCboOrigDest.Text & Space(100) & " Max : "; lblMaximo(1).Caption
           Impresora.Print "Movimiento  : " & ctlCboMov.Text & Space(100) & " Min : "; lblStock1.Caption
           Impresora.Print "Laboratorio : " & lblLab
           Impresora.Print
           Impresora.Print " SECUENCIAL               NUMERO           FECHA               _____INICIAL_____ ____INGRESO______ _____SALIDA______ _____FINAL_______"
           Impresora.Print "   KARDEX   DOC MOV       DOCUMENTO  ORI  DOCUMENTO   IMP.VENTA   UNIDAD   FRACC.  UNIDAD   FRACC.   UNIDAD   FRACC.   UNIDAD   FRACC. "
           Impresora.Print "----------- --- ---     ----------- --- ------------  --------- -------- -------- -------- -------- -------- -------- -------- --------"
           Impresora.Print
           
           For intLineas = (1 + ((intPaginas - 1) * MAX_LINEAS_X_PAGINAS)) To (intMaxLineas + ((intPaginas - 1) * MAX_LINEAS_X_PAGINAS))
              
               strCadenaKardex = CME(str(rstImpresion.Fields("sec_kardex").Value), 11, "D") & " " & _
                                 CME(rstImpresion.Fields("cod_tipodoc").Value, 3, "D") & " " & _
                                 CME(rstImpresion.Fields("tip_movimiento").Value, 3) & Space(5) & _
                                 CME(rstImpresion.Fields("num_documento").Value, 11, "D") & " " & _
                                 CME(rstImpresion.Fields("cod_local").Value, 3) & "  " & _
                                 CME(rstImpresion.Fields("fch_movimiento").Value, 11) & "  " & _
                                 CME(Format(rstImpresion.Fields("imp_producto").Value, "###,###.00"), 10) & "" & _
                                 CME(IIf(IsNull(rstImpresion.Fields(8).Value), "", rstImpresion.Fields(8).Value), 8, "D") & " " & _
                                 CME(IIf(IsNull(rstImpresion.Fields(9).Value), "", rstImpresion.Fields(9).Value), 8, "D") & " " & _
                                 CME(IIf(IsNull(rstImpresion.Fields(10).Value), "", rstImpresion.Fields(10).Value), 8, "D") & " " & _
                                 CME(IIf(IsNull(rstImpresion.Fields(11).Value), "", rstImpresion.Fields(11).Value), 8, "D") & " " & _
                                 CME(IIf(IsNull(rstImpresion.Fields(12).Value), "", rstImpresion.Fields(12).Value), 8, "D") & " " & _
                                 CME(IIf(IsNull(rstImpresion.Fields(13).Value), "", rstImpresion.Fields(13).Value), 8, "D") & " " & _
                                 CME(IIf(IsNull(rstImpresion.Fields(14).Value), "", rstImpresion.Fields(14).Value), 8, "D") & " " & _
                                 CME(IIf(IsNull(rstImpresion.Fields(15).Value), "", rstImpresion.Fields(15).Value), 8, "D") & " "
                                 'CME(IIf(IsNull(rstImpresion.Fields(16).Value), "", " F " & rstImpresion.Fields(16).Value), 8, "I")
               Impresora.Print strCadenaKardex
               rstImpresion.MoveNext
           Next intLineas
           Impresora.Print
           If intNumPaginas <> intPaginas Then
              Impresora.Print "                                                                                                                           Página " & intPaginas
           End If
           If intNumPaginas > 1 Then
              Impresora.NewPage
           End If
       Next intPaginas
       
       
      ' Impresora.Print " _RESUMEN________________________________ "
       
       
      ' strSQL = " select count(sec_kardex), cod_movimiento " & Mid(rstImpresion.Source, InStr(rstImpresion.Source, "from"))
      ' strSQL = Left(rstImpresion.Source, InStr(rstImpresion.Source, "order"))
      ' strSQL = strSQL & " group by cod_movimiento"
      '
      ' rstImpresion.Close
      ' rstImpresion.Open strSQL, conPrincipal, adOpenKeyset, adLockOptimistic
       
       Impresora.Print
       
       
       If intLotes = intNumLotes Then
       
         Impresora.Print "----------- --- --- --- ----------- --- ------------ --------- -------- -------- -------- -------- -------- -------- -------- --------"
         'impresora.Print " SECUENCIAL               NUMERO           FECHA               _____INICIAL_____ ____INGRESO______ _____SALIDA______ _____FINAL_______"
         'impresora.Print "   KARDEX   DC  MV  AL   DOCUMENTO  ORI  DOCUMENTO   IMP.VENTA   UNIDAD   FRACC.  UNIDAD   FRACC.   UNIDAD   FRACC.   UNIDAD   FRACC. "
         'impresora.Print "                                                               " & _
         '                    CME(IIf(intFracc > 0, lngIniU + (lngIniF \ intFracc), lngIniU), 8, "D") & IIf(lngIniF > 0, " F " & CME(IIf(lngIniF >= intFracc, lngIniF Mod intFracc, lngIniF), 7, "I"), Space$(10)) & _
         '                    CME(IIf(intFracc > 0, lngIngU + (lngIngF \ intFracc), lngIngU), 8, "D") & IIf(lngIngF > 0, " F " & CME(IIf(lngIngF >= intFracc, lngIngF Mod intFracc, lngIngF), 7, "I"), Space$(10)) & _
         ''                    CME(IIf(intFracc > 0, lngSalU + (lngSalF \ intFracc), lngSalU), 8, "D") & IIf(lngSalF > 0, " F " & CME(IIf(lngSalF >= intFracc, lngSalF Mod intFracc, lngSalF), 7, "I"), Space$(10)) & _
         '                    CME(IIf(intFracc > 0, lngFinU + (lngFinF \ intFracc), lngFinU), 8, "D") & IIf(lngFinF > 0, " F " & CME(IIf(lngFinF >= intFracc, lngFinF Mod intFracc, lngFinF), 7, "I"), Space$(10))
'         Impresora.Print "                                                               " & _
                             CME(Trim(lblTotales(0).Caption), 8, "D") & Space$(1) & CME(IIf(intFracc > 0, Trim(lblTotales(1).Caption), Space$(10)), 9, "I") & _
                             CME(Trim(lblTotales(2).Caption), 8, "D") & Space$(1) & CME(IIf(intFracc > 0, Trim(lblTotales(3).Caption), Space$(10)), 9, "I") & _
                             CME(Trim(lblTotales(4).Caption), 8, "D") & Space$(1) & CME(IIf(intFracc > 0, Trim(lblTotales(5).Caption), Space$(10)), 9, "I") & _
                             CME(Trim(lblTotales(6).Caption), 8, "D") & Space$(1) & CME(IIf(intFracc > 0, Trim(lblTotales(7).Caption), Space$(10)), 9, "I")
        ' Impresora.Print "----------- --- --- --- ----------- --- ------------ --------- -------- -------- -------- -------- -------- -------- -------- --------"
       
       
         'Impresora.Print "   " & lblregistros.Caption
         
         Dim lngPosP As Long
         Dim lngPosU As Long
         Dim lngtmp As Long
         
         rstImpresion.MoveFirst
         lngPosP = rstImpresion.Fields(0).Value
         rstImpresion.MoveLast
         lngPosU = rstImpresion.Fields(0).Value
         If lngPosP > lngPosU Then
            lngtmp = lngPosP
            lngPosP = lngPosU
            lngPosU = lngtmp
         End If
         
         'strSQL = "Select count(sec_kardex), cod_movimiento,sum(ctd_ingreso),sum(ctd_ingresofrac),sum(ctd_salida),sum(ctd_salidafrac) from m_kardex where cod_btl = '" & "002" & "' and cod_almacen = '" & Right(ctlCboOrigDest.Text, 2) & "' and cod_producto = '" & TxtProducto.Text & "' and sec_kardex between " & lngPosP & " and " & lngPosU & " Group by cod_movimiento"
         'strSQL = "Select count(sec_kardex), cod_movimiento,sum(ctd_ingreso),sum(ctd_ingresofrac),sum(ctd_salida),sum(ctd_salidafrac) from m_kardex " & Replace(strWhere, "ORDER BY SEC_KARDEX", "") & " Group by cod_movimiento"
         
         'Set rstImpresion = OraDatabase.CreateDynaset(strSQL, ORADYN_DEFAULT)
         'rstImpresion.Open strSQL, conPrincipal, adOpenKeyset, adLockOptimistic
'         If Not (rstImpresion.EOF) Then
'            Impresora.Print " -----RESUMEN DE MOVIMIENTOS-------------------ING-------------SAL-------- "
'            Impresora.Print " |                                                                       | "
'            While Not (rstImpresion.EOF)
'               If intFracc = 0 Then
'                  lngI = rstImpresion.Fields(2).Value
'                  lngIF = rstImpresion.Fields(3).Value
'                  lngS = rstImpresion.Fields(4).Value
'                  lngSF = rstImpresion.Fields(5).Value
'               Else
'                  lngI = rstImpresion.Fields(2).Value + rstImpresion.Fields(3).Value \ intFracc
'                  lngIF = rstImpresion.Fields(3).Value Mod intFracc
'                  lngS = rstImpresion.Fields(4).Value + rstImpresion.Fields(5).Value \ intFracc
'                  lngSF = rstImpresion.Fields(5).Value Mod intFracc
'               End If
'
'               'lngI = IIf(intFracc = 0, rstImpresion.Fields(2), rstImpresion.Fields(2) + rstImpresion.Fields(3) \ intFracc)
'               'lngIF = IIf(intFracc = 0, rstImpresion.Fields(3), rstImpresion.Fields(3) Mod intFracc)
'               'lngS = IIf(intFracc = 0, rstImpresion.Fields(4), rstImpresion.Fields(4) + rstImpresion.Fields(5) \ intFracc)
'               'lngSF = IIf(intFracc = 0, rstImpresion.Fields(5), rstImpresion.Fields(5) Mod intFracc)
'               Impresora.Print " | " & CME(rstImpresion.Fields(1).Value & " - " & strDameMovimiento(rstImpresion.Fields(1).Value), 25) & CME(rstImpresion.Fields(0).Value, 12, "D") & CME(IIf(lngI > 0, CStr(lngI), ""), 8, "D") & CME(IIf(lngIF > 0, "F " & lngIF, ""), 8) & CME(IIf(lngS > 0, CStr(lngS), ""), 8, "D") & CME(IIf(lngSF > 0, "F " & lngSF, ""), 8) & " |"
'               rstImpresion.MoveNext
'            Wend
           ' Impresora.Print " ------------------------------------------------------------------------- "
            Impresora.Print
            Impresora.Print "-----STOCK--------------------"
            Impresora.Print "|    " & CME(lblStock1.Caption, 23, "I") & "|"
            Impresora.Print "------------------------------"
'         End If
         'Impresora.Print "                                                                                                                           Página " & intPaginas
       End If
              
       Impresora.EndDoc
       
   Next intLotes
   
   Set rstImpresion = Nothing
   
   On Error GoTo 0
   
   Exit Sub

ErrorImpresora:

   intRespuesta = MsgBox("Existe un problema con la impresora" & Chr(13) & _
                  "Desea esperar a ser resuelto?", vbYesNo)
   Err.Clear
   If intRespuesta = vbYes Then
       Resume
   Else
       Exit Sub
   End If

End Sub

