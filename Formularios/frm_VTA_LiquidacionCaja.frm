VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frm_VTA_LiquidacionCaja 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidación de Caja"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   6825
   ClientWidth     =   8430
   ForeColor       =   &H00000000&
   Icon            =   "frm_VTA_LiquidacionCaja.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "[ C ]  Secuencia de Documentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   0
      TabIndex        =   88
      Top             =   5520
      Width           =   8415
      Begin vbp_Ventas.ctlGrillaArray ctlGrdSecDoc 
         Height          =   1095
         Left            =   120
         TabIndex        =   89
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   1931
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "[ B ]  Declarado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   78
      Top             =   3720
      Width           =   8415
      Begin vbp_Ventas.ctlGrilla grdDeclarado 
         Height          =   1335
         Left            =   120
         TabIndex        =   79
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   2355
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total =>"
         Height          =   195
         Left            =   6360
         TabIndex        =   83
         Top             =   1610
         Width           =   585
      End
      Begin VB.Label LblTotDecl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   7080
         TabIndex        =   82
         Top             =   1610
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdGrabaVenta 
      Caption         =   "&Grabar [F2]"
      Height          =   735
      Left            =   7290
      Picture         =   "frm_VTA_LiquidacionCaja.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   7370
      Width           =   1065
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "[ A ]  Sistema --------- (presione F1 para actualizar)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   76
      Top             =   600
      Width           =   8415
      Begin vbp_Ventas.ctlGrilla grdSistema 
         Height          =   2595
         Left            =   120
         TabIndex        =   77
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4577
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total =>"
         Height          =   195
         Left            =   6360
         TabIndex        =   81
         Top             =   2880
         Width           =   585
      End
      Begin VB.Label LblTotSist 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   7080
         TabIndex        =   80
         Top             =   2880
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4050
      TabIndex        =   62
      Top             =   3120
      Visible         =   0   'False
      Width           =   3855
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TC"
         Height          =   195
         Left            =   1320
         TabIndex        =   74
         Top             =   1080
         Width           =   210
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ".- Efectivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "[ Local Declara ]"
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
         Left            =   120
         TabIndex        =   70
         Top             =   600
         Width           =   1440
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Soles     S/."
         Height          =   195
         Left            =   240
         TabIndex        =   69
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Dolares  $."
         Height          =   195
         Left            =   240
         TabIndex        =   68
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label LblSolesDec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   2400
         TabIndex        =   67
         Top             =   840
         Width           =   855
      End
      Begin VB.Label LblDolaresDec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   2400
         TabIndex        =   66
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label LblTCdec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   1680
         TabIndex        =   65
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "(11)"
         Height          =   195
         Left            =   3360
         TabIndex        =   64
         Top             =   840
         Width           =   270
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "(12)"
         Height          =   195
         Left            =   3360
         TabIndex        =   63
         Top             =   1080
         Width           =   270
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4080
      TabIndex        =   52
      Top             =   720
      Visible         =   0   'False
      Width           =   3855
      Begin TrueDBGrid70.TDBGrid grdDonacion 
         Bindings        =   "frm_VTA_LiquidacionCaja.frx":0596
         Height          =   495
         Left            =   360
         TabIndex        =   53
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   873
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "COD"
         Columns(0).DataField=   "COD"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "DES"
         Columns(1).DataField=   "DES"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "MONTO"
         Columns(2).DataField=   "MONTO"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).DividerColor=   16777215
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=476"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
         Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=423"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2831"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2752"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=1588"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerStyle=0"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1535"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   0
         BorderStyle     =   0
         DefColWidth     =   0
         HeadLines       =   0
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   16777215
         RowDividerColor =   16777215
         RowSubDividerColor=   16777215
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.borderColor=&H80000005&"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1,.valignment=2,.fgcolor=&HFFFFFF&"
         _StyleDefs(13)  =   ":id=7,.borderColor=&H80000005&"
         _StyleDefs(14)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(19)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HFFFFFF&"
         _StyleDefs(22)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(25)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(26)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(27)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.appearance=0"
         _StyleDefs(28)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(29)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(30)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1,.bgcolor=&HD8D8D8&"
         _StyleDefs(40)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(43)  =   "Named:id=33:Normal"
         _StyleDefs(44)  =   ":id=33,.parent=0"
         _StyleDefs(45)  =   "Named:id=34:Heading"
         _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   ":id=34,.wraptext=-1"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(52)  =   "Named:id=37:Caption"
         _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(54)  =   "Named:id=38:HighlightRow"
         _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H0&"
         _StyleDefs(56)  =   "Named:id=39:EvenRow"
         _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(58)  =   "Named:id=40:OddRow"
         _StyleDefs(59)  =   ":id=40,.parent=33"
         _StyleDefs(60)  =   "Named:id=41:RecordSelector"
         _StyleDefs(61)  =   ":id=41,.parent=34"
         _StyleDefs(62)  =   "Named:id=42:FilterBar"
         _StyleDefs(63)  =   ":id=42,.parent=33,.appearance=0"
      End
      Begin ORADCLibCtl.ORADC oradcDona 
         Height          =   255
         Left            =   840
         Top             =   720
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
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
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ".- Documento Dscto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   1560
         Width           =   1710
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ".- Donación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label LblDD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   2520
         TabIndex        =   59
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Soles     S/."
         Height          =   195
         Left            =   600
         TabIndex        =   58
         Top             =   1920
         Width           =   840
      End
      Begin VB.Label LblTotDona 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   2520
         TabIndex        =   57
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total =>"
         Height          =   195
         Left            =   1800
         TabIndex        =   56
         Top             =   1215
         Width           =   585
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "(9)"
         Height          =   195
         Left            =   3480
         TabIndex        =   55
         Top             =   1200
         Width           =   180
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "(10)"
         Height          =   195
         Left            =   3480
         TabIndex        =   54
         Top             =   1920
         Width           =   270
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   55
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   3855
      Begin TrueDBGrid70.TDBGrid grdTarjetas 
         Bindings        =   "frm_VTA_LiquidacionCaja.frx":05AE
         Height          =   495
         Left            =   285
         TabIndex        =   17
         Top             =   2520
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   873
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Codigo"
         Columns(0).DataField=   "COD"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descripción"
         Columns(1).DataField=   "DES"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Monto"
         Columns(2).DataField=   "MONTO"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).DividerColor=   16777215
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=529"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
         Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=16777215"
         Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=476"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2858"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2778"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=1614"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerStyle=0"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1561"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   0
         BorderStyle     =   0
         DefColWidth     =   0
         HeadLines       =   0
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   16777215
         RowDividerColor =   16777215
         RowSubDividerColor=   16777215
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.borderColor=&H80000005&"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1,.valignment=2,.fgcolor=&HFFFFFF&"
         _StyleDefs(13)  =   ":id=7,.borderColor=&H80000005&"
         _StyleDefs(14)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(19)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HFFFFFF&"
         _StyleDefs(22)  =   "Splits(0).FooterStyle:id=15,.parent=3,.bgcolor=&HFFFFFF&"
         _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(25)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(26)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(27)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.appearance=0"
         _StyleDefs(28)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(29)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(30)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1,.bgcolor=&HD8D8D8&"
         _StyleDefs(40)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(43)  =   "Named:id=33:Normal"
         _StyleDefs(44)  =   ":id=33,.parent=0"
         _StyleDefs(45)  =   "Named:id=34:Heading"
         _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   ":id=34,.wraptext=-1"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(52)  =   "Named:id=37:Caption"
         _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(54)  =   "Named:id=38:HighlightRow"
         _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H0&"
         _StyleDefs(56)  =   "Named:id=39:EvenRow"
         _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(58)  =   "Named:id=40:OddRow"
         _StyleDefs(59)  =   ":id=40,.parent=33"
         _StyleDefs(60)  =   "Named:id=41:RecordSelector"
         _StyleDefs(61)  =   ":id=41,.parent=34"
         _StyleDefs(62)  =   "Named:id=42:FilterBar"
         _StyleDefs(63)  =   ":id=42,.parent=33,.appearance=0"
      End
      Begin TrueDBGrid70.TDBGrid grdCredito 
         Bindings        =   "frm_VTA_LiquidacionCaja.frx":05C5
         Height          =   495
         Left            =   285
         TabIndex        =   18
         Top             =   4560
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   873
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Codigo"
         Columns(0).DataField=   "COD"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descripcion"
         Columns(1).DataField=   "DES"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "MONTO"
         Columns(2).DataField=   "MONTO"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).DividerColor=   16777215
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=450"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
         Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=397"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2884"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2805"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=1640"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerStyle=0"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1588"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   0
         BorderStyle     =   0
         DefColWidth     =   0
         HeadLines       =   0
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   16777215
         RowDividerColor =   16777215
         RowSubDividerColor=   16777215
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.borderColor=&H80000005&"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1,.valignment=2,.fgcolor=&HFFFFFF&"
         _StyleDefs(13)  =   ":id=7,.borderColor=&H80000005&"
         _StyleDefs(14)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(19)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HFFFFFF&"
         _StyleDefs(22)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(25)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(26)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(27)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.appearance=0"
         _StyleDefs(28)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(29)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(30)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1,.bgcolor=&HD8D8D8&"
         _StyleDefs(40)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(43)  =   "Named:id=33:Normal"
         _StyleDefs(44)  =   ":id=33,.parent=0"
         _StyleDefs(45)  =   "Named:id=34:Heading"
         _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   ":id=34,.wraptext=-1"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(52)  =   "Named:id=37:Caption"
         _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(54)  =   "Named:id=38:HighlightRow"
         _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H0&"
         _StyleDefs(56)  =   "Named:id=39:EvenRow"
         _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(58)  =   "Named:id=40:OddRow"
         _StyleDefs(59)  =   ":id=40,.parent=33"
         _StyleDefs(60)  =   "Named:id=41:RecordSelector"
         _StyleDefs(61)  =   ":id=41,.parent=34"
         _StyleDefs(62)  =   "Named:id=42:FilterBar"
         _StyleDefs(63)  =   ":id=42,.parent=33,.appearance=0"
      End
      Begin ORADCLibCtl.ORADC oradcTar 
         Height          =   135
         Left            =   2085
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   238
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
      Begin ORADCLibCtl.ORADC oradcCredito 
         Height          =   135
         Left            =   2160
         Top             =   4680
         Visible         =   0   'False
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   238
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
      Begin TrueDBGrid70.TDBGrid TDBGrid1 
         Bindings        =   "frm_VTA_LiquidacionCaja.frx":05E0
         Height          =   495
         Left            =   360
         TabIndex        =   75
         Top             =   1440
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   873
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Codigo"
         Columns(0).DataField=   "COD"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descripción"
         Columns(1).DataField=   "DES"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Monto"
         Columns(2).DataField=   "MONTO"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).DividerColor=   16777215
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=529"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
         Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=16777215"
         Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=476"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2858"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2778"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=1614"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerStyle=0"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1561"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   0
         BorderStyle     =   0
         DefColWidth     =   0
         HeadLines       =   0
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   16777215
         RowDividerColor =   16777215
         RowSubDividerColor=   16777215
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.borderColor=&H80000005&"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1,.valignment=2,.fgcolor=&HFFFFFF&"
         _StyleDefs(13)  =   ":id=7,.borderColor=&H80000005&"
         _StyleDefs(14)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(19)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HFFFFFF&"
         _StyleDefs(22)  =   "Splits(0).FooterStyle:id=15,.parent=3,.bgcolor=&HFFFFFF&"
         _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(25)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(26)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(27)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.appearance=0"
         _StyleDefs(28)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(29)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(30)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1,.bgcolor=&HD8D8D8&"
         _StyleDefs(40)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(43)  =   "Named:id=33:Normal"
         _StyleDefs(44)  =   ":id=33,.parent=0"
         _StyleDefs(45)  =   "Named:id=34:Heading"
         _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   ":id=34,.wraptext=-1"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(52)  =   "Named:id=37:Caption"
         _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(54)  =   "Named:id=38:HighlightRow"
         _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H0&"
         _StyleDefs(56)  =   "Named:id=39:EvenRow"
         _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(58)  =   "Named:id=40:OddRow"
         _StyleDefs(59)  =   ":id=40,.parent=33"
         _StyleDefs(60)  =   "Named:id=41:RecordSelector"
         _StyleDefs(61)  =   ":id=41,.parent=34"
         _StyleDefs(62)  =   "Named:id=42:FilterBar"
         _StyleDefs(63)  =   ":id=42,.parent=33,.appearance=0"
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TC"
         Height          =   195
         Left            =   1440
         TabIndex        =   73
         Top             =   7080
         Width           =   210
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TC"
         Height          =   195
         Left            =   1560
         TabIndex        =   72
         Top             =   1080
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ".- Efectivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "[ Sistema ]"
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
         Left            =   285
         TabIndex        =   50
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Soles     S/."
         Height          =   195
         Left            =   405
         TabIndex        =   49
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Dolares  $."
         Height          =   195
         Left            =   405
         TabIndex        =   48
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label LblSolesDoc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   2565
         TabIndex        =   47
         Top             =   840
         Width           =   855
      End
      Begin VB.Label LblDolaresDoc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   2565
         TabIndex        =   46
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ".- Tarjetas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ".- Nota de Credito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   3480
         Width           =   1530
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Soles     S/."
         Height          =   195
         Left            =   360
         TabIndex        =   43
         Top             =   3840
         Width           =   840
      End
      Begin VB.Label LblNC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   2565
         TabIndex        =   42
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ".- Credito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   4200
         Width           =   795
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ".- Cheque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   6240
         Width           =   840
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ".- Crob. Venta a Cred."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   5520
         Width           =   1875
      End
      Begin VB.Label lLblCVC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   2565
         TabIndex        =   38
         Top             =   5880
         Width           =   855
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Soles     S/."
         Height          =   195
         Left            =   360
         TabIndex        =   37
         Top             =   5880
         Width           =   840
      End
      Begin VB.Label LblTotTarj 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   2565
         TabIndex        =   36
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label LblTotCred 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   2565
         TabIndex        =   35
         Top             =   5160
         Width           =   855
      End
      Begin VB.Label lblTotChequeSol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   2565
         TabIndex        =   34
         Top             =   6720
         Width           =   855
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total =>"
         Height          =   195
         Left            =   1845
         TabIndex        =   33
         Top             =   3135
         Width           =   585
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total =>"
         Height          =   195
         Left            =   1845
         TabIndex        =   32
         Top             =   5175
         Width           =   585
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "(2)"
         Height          =   195
         Left            =   3525
         TabIndex        =   31
         Top             =   1080
         Width           =   180
      End
      Begin VB.Label LblTCdoc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   1845
         TabIndex        =   30
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "(1)"
         Height          =   195
         Left            =   3525
         TabIndex        =   29
         Top             =   840
         Width           =   180
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "(3)"
         Height          =   195
         Left            =   3525
         TabIndex        =   28
         Top             =   3120
         Width           =   180
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "(4)"
         Height          =   195
         Left            =   3525
         TabIndex        =   27
         Top             =   3840
         Width           =   180
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "(5)"
         Height          =   195
         Left            =   3525
         TabIndex        =   26
         Top             =   5160
         Width           =   180
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "(6)"
         Height          =   195
         Left            =   3525
         TabIndex        =   25
         Top             =   5880
         Width           =   180
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "(7)"
         Height          =   195
         Left            =   3525
         TabIndex        =   24
         Top             =   6720
         Width           =   180
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Soles     S/."
         Height          =   195
         Left            =   360
         TabIndex        =   23
         Top             =   6720
         Width           =   840
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "(8)"
         Height          =   195
         Left            =   3525
         TabIndex        =   22
         Top             =   7080
         Width           =   180
      End
      Begin VB.Label LblTotChequeDol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   2565
         TabIndex        =   21
         Top             =   7080
         Width           =   855
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Dolares  $."
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   7080
         Width           =   765
      End
      Begin VB.Label LblTCchqDol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Caption         =   "0.00"
         Height          =   225
         Left            =   1800
         TabIndex        =   19
         Top             =   7080
         Width           =   600
      End
   End
   Begin vbp_Ventas.ctlTextBox TxtRepoCaja 
      Height          =   315
      Left            =   7280
      TabIndex        =   0
      Top             =   7000
      Width           =   1095
      _ExtentX        =   1931
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
   Begin VB.CheckBox chkActivar 
      BackColor       =   &H80000009&
      Caption         =   "Reposición Caja"
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
      Height          =   255
      Left            =   5520
      TabIndex        =   87
      Top             =   7040
      Width           =   1815
   End
   Begin VB.Label txtCerrada 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "CERRADA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   3960
      TabIndex        =   86
      Top             =   7080
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label LblLocal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   10
      Width           =   855
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Liquidación Nº"
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
      Left            =   3960
      TabIndex        =   85
      Top             =   7560
      Width           =   1260
   End
   Begin VB.Label LblCodLiquidacion 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   3960
      TabIndex        =   84
      Top             =   7760
      Width           =   3195
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "( B )      Total Declarado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   7440
      Width           =   2100
   End
   Begin VB.Label LblTotDeclarado 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0.00"
      Height          =   225
      Left            =   2760
      TabIndex        =   14
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label LblCaja 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "99999999"
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
      Left            =   7200
      TabIndex        =   13
      Top             =   280
      Width           =   855
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Caja :"
      Height          =   195
      Left            =   6600
      TabIndex        =   12
      Top             =   280
      Width           =   405
   End
   Begin VB.Line Line5 
      BorderStyle     =   3  'Dot
      X1              =   3720
      X2              =   3720
      Y1              =   7080
      Y2              =   8040
   End
   Begin VB.Line Line4 
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   3720
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line3 
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   3720
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   120
      Y1              =   7080
      Y2              =   8040
   End
   Begin VB.Label LblDif 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0.00"
      Height          =   225
      Left            =   2760
      TabIndex        =   11
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label LblTotCobrado 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0.00"
      Height          =   225
      Left            =   2760
      TabIndex        =   10
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "( B - A ) T. Cobrado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   7680
      Width           =   1680
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "( A )      Total Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   7200
      Width           =   1890
   End
   Begin VB.Label LblFecha 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "dd/mm/yyyy"
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
      Left            =   7200
      TabIndex        =   7
      Top             =   30
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fecha :"
      Height          =   195
      Left            =   6600
      TabIndex        =   6
      Top             =   30
      Width           =   540
   End
   Begin VB.Label LblCajero 
      BackColor       =   &H00FFFFFF&
      Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   270
      Width           =   4335
   End
   Begin VB.Label LblQuimico 
      BackColor       =   &H00FFFFFF&
      Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
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
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   30
      Width           =   4455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vendedor :"
      Height          =   195
      Left            =   75
      TabIndex        =   2
      Top             =   270
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quimico    :"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   30
      Width           =   795
   End
End
Attribute VB_Name = "frm_VTA_LiquidacionCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim odynLiqAux As oraDynaset

Dim odynLiquiSist As oraDynaset
Dim odynLiquiDecl As oraDynaset

Dim objImpresion As New clsImpresiones

Dim strCadRetEfect As String
Dim strCadFP As String
Dim strCadFPH As String

'-- Var Totales de remesa donde va la suma del monto x moneda--'
Dim strCadSist As String
Dim strCadEfect As String
Dim strCadEfectTotal As String
'--------------------------------------------------------------'

'-- Var nuevas para la el detalle de la suma de remesa por moneda --'
Dim strCadCodRemesaP As String
Dim strCadCodRemesaH As String
'Dim strCadEfectSol As String
'Dim strCadEfectDol As String
'Dim strCadEfectSoles As String
'Dim strCadEfectDolares As String
'-------------------------------------------------------------------'
Dim strEfect_Val As String
Dim strCadDif As String
Dim strCadValor As String
Dim blnPase As Boolean

Dim dblTotSist As Double
Dim dblTotDecl As Double
Dim strValor As String

Dim blnCerrada As Boolean

Dim strCaja As String
Dim strLiquidacion As String

'-- Var para el detalle de la secuencia de documentos
Dim xarrSecDoc As New XArrayDB
Private lblnCerrar As Boolean

Private Sub chkActivar_Click()
    If chkActivar.Value = 0 Then
        TxtRepoCaja.Enabled = False
    Else
        TxtRepoCaja.Enabled = True
    End If
End Sub

Private Sub cmdGrabaVenta_Click()
    If Not blnCerrada Then
        Graba
    End If
End Sub

Private Sub ctlGrdSecDoc_AfterColUpdate(ByVal ColIndex As Integer)
    ctlGrdSecDoc.MoveNext
    ctlGrdSecDoc.MovePrevious
End Sub

Private Sub ctlGrdSecDoc_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    Select Case ColIndex
        Case 0, 1, 3
            Cancel = 1
    End Select
End Sub

Private Sub Form_Activate()
If lblnCerrar = True Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
    On Error GoTo CtrlErr
    lblnCerrar = False
    Set odynLiqAux = objLiquidacion.fnCalFormaPagos(objUsuario.CodigoEmpresa, _
                                                objUsuario.CodigoLocal, _
                                                strCaja, _
                                                strValor, _
                                                strLiquidacion)

    
    If odynLiqAux.RecordCount <= 0 Then
        MsgBox "La caja no tiene documentos emitidos", vbCritical, App.ProductName
        Set odynLiqAux = Nothing
        lblnCerrar = True
        Exit Sub
    End If
    fnCalcSistema
    
    LblLocal.Caption = objUsuario.CodigoLocal
    LblFecha.Caption = Trim(Fecha)    'Trim(frm_Lista_Caja_PreCerradas.grdCajas.Columns("FCH_INICIO").Value)
    LblQuimico.Caption = objUsuario.Codigo & " " & objLiquidacion.fnDevNomUsu(objUsuario.Codigo)
    LblCajero.Caption = Cajero & " " & objLiquidacion.NomUsuario(Cajero)    'frm_Lista_Caja_PreCerradas.grdCajas.Columns("USU_TEC").Value & " " & objLiquidacion.NomUsuario(frm_Lista_Caja_PreCerradas.grdCajas.Columns("USU_TEC").Value)
    LblCaja.Caption = Caja  'frm_Lista_Caja_PreCerradas.grdCajas.Columns("COD_MAQUINA").Value
        
    SeteaGrilla
    SeteaGrillaRem
    SeteaGrillaSecxDoc
    
    strValor = objUsuario.Parametros("COD_MONEDA")(0, 2)
    
    blnPase = True
    LblCodLiquidacion.Caption = Liquidacion  'Trim(frm_Lista_Caja_PreCerradas.grdCajas.Columns("COD_LIQUIDACION").Value)
    
    TxtRepoCaja.Enabled = False
    chkActivar.Visible = True
    DeshabilitaObjetos
    
    Consulta
    TxtRepoCaja.Text = Val(objLiquidacion.ListaMtoCajaChica(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, Liquidacion))
    
    Exit Sub
CtrlErr:
   MsgBox Err.Description, vbOKOnly, "Error"
    
End Sub

Sub fnSecuenciaDocumentos()
    Dim odynSecDoc As oraDynaset

    Set odynSecDoc = objLiquidacion.Lista_Sec_x_Doc(objUsuario.CodigoEmpresa, _
                                                    objUsuario.CodigoLocal, _
                                                    Liquidacion)
    If odynSecDoc.RecordCount > 0 Then
        xarrSecDoc.LoadRows odynSecDoc.GetRows
    End If
    
    ctlGrdSecDoc.Array1 = xarrSecDoc
    
End Sub

Private Sub SeteaGrillaSecxDoc()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim i As Integer
  
    arrCampos = Array("DOCUMENTO", "SEC_MINIMO_CAB", "SEC_MINIMO", "SEC_MAXIMO_CAB", "SEC_MAXIMO")
                      
    arrCaption = Array("Tipo Doc.", "Sec Minímo", "Desde", "Sec Maxímo", "Hasta")
    
    arrAncho = Array(1500, 1800, 1800, 1800, 1800)
    
    arrAlineacion = Array(vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft, vbAlignLeft)
    
    With ctlGrdSecDoc
        .FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        .EditorStyle.BackColor = vbWhite
        .EditorStyle.ForeColor = RGB(180, 0, 180)
        .EditorStyle.Font.Bold = True
        .AllowUpdate = True
        '.RowHeight = 1.2 * .RowHeight
        .MarqueeStyle = 2
        .col = 3

        For i = 0 To .Columns.Count - 1
            .Columns(i).AllowSizing = False
            .Columns(i).WrapText = False
        Next i
        .Columns(1).Visible = False
        .Columns(3).Visible = False
        
        'Columnas editables
        .Columns(2).BackColor = vbInfoBackground
        .Columns(2).DataWidth = 10
        .Columns(2).EditMask = "##########"
        .Columns(4).BackColor = vbInfoBackground
        .Columns(4).DataWidth = 10
        .Columns(4).EditMask = "##########"
    End With
End Sub

Sub fnCalcSistema()
    Dim odynClonSist As oraDynaset
    'Dim dblTotSist As Double
    
''''''''    Set odynLiquiSist = objLiquidacion.fnCalFormaPagos(objUsuario.CodigoEmpresa, _
''''''''                                                       objUsuario.CodigoLocal, _
''''''''                                                       frm_Lista_Caja_PreCerradas.grdCajas.Columns("COD_MAQUINA").Value, _
''''''''                                                       strValor, _
''''''''                                                       frm_Lista_Caja_PreCerradas.grdCajas.Columns("COD_LIQUIDACION").Value)
        
'' ** Comentado el 06/01/2009 Por Cristhian Rueda
        
''    Set odynLiquiSist = objLiquidacion.fnCalFormaPagos(objUsuario.CodigoEmpresa, _
''                                                       objUsuario.CodigoLocal, _
''                                                       Caja, _
''                                                       strValor, _
''                                                       Liquidacion)
        
'''''''    Set grdSistema.DataSource = objLiquidacion.fnCalFormaPagos(objUsuario.CodigoEmpresa, _
'''''''                                                         objUsuario.CodigoLocal, _
'''''''                                                         frm_Lista_Caja_PreCerradas.grdCajas.Columns("COD_MAQUINA").Value, _
'''''''                                                         strValor, _
'''''''                                                         frm_Lista_Caja_PreCerradas.grdCajas.Columns("COD_LIQUIDACION").Value)
    
'' ** Comentado el 06/01/2009 Por Cristhian Rueda

''    Set grdSistema.DataSource = objLiquidacion.fnCalFormaPagos(objUsuario.CodigoEmpresa, _
''                                                         objUsuario.CodigoLocal, _
''                                                         Caja, _
''                                                         strValor, _
''                                                         Liquidacion)
    'podynLiq.Refresh
    Set grdSistema.DataSource = odynLiqAux
    Set odynClonSist = odynLiqAux.Clone
    dblTotSist = 0
    odynClonSist.MoveFirst
    'Todas las formas de pagos se suma ecepto el redondeo'
        While Not odynClonSist.EOF
           If (odynClonSist("COD_FORMA_PAGO").Value = "001") Then
              dblTotSist = dblTotSist + odynClonSist("TOTAL").Value
           End If

           If odynClonSist("COD_FORMA_PAGO").Value = "002" Then
                If odynClonSist("FLG_RETIRO").Value = "1" Then
                   dblTotSist = dblTotSist - Val(odynClonSist("TOTAL").Value)
                End If
            End If
            
           If (odynClonSist("COD_FORMA_PAGO").Value = "008") Then   'agregado por jmelgar
              dblTotSist = dblTotSist - odynClonSist("TOTAL").Value
           End If

           odynClonSist.MoveNext
        Wend

    LblTotSist.Caption = Format(dblTotSist, "#,###,##0.00")
    
End Sub

Sub fnCalcDeclRemesa()
    Dim odynClonDecl As oraDynaset
    
'    Set odynLiquiDecl = objLiquidacion.fnTotDeclLocal(objUsuario.CodigoEmpresa, _
'                                                      objUsuario.CodigoLocal, _
'                                                      frm_Lista_Caja_PreCerradas.grdCajas.Columns("COD_MAQUINA").Value, _
'                                                      Trim(LblCodLiquidacion.Caption))
'
'    Set grdDeclarado.DataSource = objLiquidacion.fnTotDeclLocal(objUsuario.CodigoEmpresa, _
'                                                                objUsuario.CodigoLocal, _
'                                                                frm_Lista_Caja_PreCerradas.grdCajas.Columns("COD_MAQUINA").Value, _
'                                                                Trim(LblCodLiquidacion.Caption))
    
    Set odynLiquiDecl = objLiquidacion.fnTotDeclLocal(objUsuario.CodigoEmpresa, _
                                                      objUsuario.CodigoLocal, _
                                                      Caja, _
                                                      Trim(LblCodLiquidacion.Caption))
                                                      
    Set grdDeclarado.DataSource = objLiquidacion.fnTotDeclLocal(objUsuario.CodigoEmpresa, _
                                                                objUsuario.CodigoLocal, _
                                                                Caja, _
                                                                Trim(LblCodLiquidacion.Caption))
    
    Set odynClonDecl = odynLiquiDecl.Clone
    odynClonDecl.MoveFirst
    While Not odynClonDecl.EOF
        dblTotDecl = dblTotDecl + odynClonDecl("SUBTOTAL").Value
        odynClonDecl.MoveNext
    Wend
    LblTotDecl.Caption = Format(dblTotDecl, "#,###,##0.00")
End Sub
    
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            If Not blnCerrada Then
                Consulta
            End If
        Case vbKeyF2
            Call cmdGrabaVenta_Click
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Sub Consulta()
    dblTotDecl = 0
    'dblTotSist = 0
    
    'fnCalcSistema
    fnCalcDeclRemesa
    fnSecuenciaDocumentos
    
    LblTotCobrado.Caption = Format(dblTotSist, "#,#,##0.00")
    LblTotDeclarado.Caption = Format(dblTotDecl, "#,###,##0.00")
    LblDif.Caption = Format(dblTotDecl - dblTotSist, "#,###,##0.00")
End Sub

Sub Graba()
    Dim i%
    Dim CodLiquidacion$
    Dim strCadCodDoc As String, strCadMinDoc As String, strCadMaxDoc As String
    
    strCadDif = "": strCadFP = "": strCadFPH = ""
    strCadEfectTotal = "": strCadEfect = "": strCadSist = "":
    strCadRetEfect = ""
    strValor = ""
    strEfect_Val = ""
    strCadCodDoc = ""
    strCadMinDoc = ""
    strCadMaxDoc = ""
    
    If ctlGrdSecDoc.EditActive = True Then Call ctlGrdSecDoc_AfterColUpdate(0)
    If xarrSecDoc.Count(1) > 0 Then
        For i = xarrSecDoc.LowerBound(1) To xarrSecDoc.UpperBound(1)
            If IsNull(Trim(xarrSecDoc.Value(i, 2))) Or LenB(Trim(xarrSecDoc.Value(i, 2))) = 0 Or _
               IsNull(Trim(xarrSecDoc.Value(i, 4))) Or LenB(Trim(xarrSecDoc.Value(i, 4))) = 0 Then
                MsgBox "Debe ingresar los correlativos de " & xarrSecDoc.Value(i, 0) & ".", vbCritical, App.ProductName
                ctlGrdSecDoc.SetFocus
                Exit Sub
            End If
            If StrComp(Trim(xarrSecDoc.Value(i, 1)), Trim(xarrSecDoc.Value(i, 2)), vbBinaryCompare) <> 0 Or _
               StrComp(Trim(xarrSecDoc.Value(i, 3)), Trim(xarrSecDoc.Value(i, 4)), vbBinaryCompare) <> 0 Then
                MsgBox "Error en la secuencia de documentos." & vbCrLf & _
                       "Debe corregir los correlativos de " & xarrSecDoc.Value(i, 0) & ".", vbCritical, App.ProductName
                grdSistema.SetFocus
                Exit Sub
            End If
            strCadCodDoc = strCadCodDoc & Trim(xarrSecDoc.Value(i, 0)) & "|"
            strCadMinDoc = strCadMinDoc & Trim(xarrSecDoc.Value(i, 2)) & "|"
            strCadMaxDoc = strCadMaxDoc & Trim(xarrSecDoc.Value(i, 4)) & "|"
        Next i
    End If
    strCadCodRemesaP = "": strCadCodRemesaH = ""
    'strCadEfectSol = "": strCadEfectDol = ""
    'strCadEfectSoles = "": strCadEfectDolares = ""
    
    If grdSistema.ApproxCount <= 0 Then MsgBox "No tiene ventas asociadas a esta caja", vbCritical, App.ProductName: Exit Sub
    If grdDeclarado.ApproxCount <= 0 Then MsgBox "Ingrese la remesa para esta caja", vbCritical, App.ProductName: Exit Sub
    
    odynLiqAux.MoveFirst
    
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/'
    '**** Valores para la cadena de la venta ***'
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/'
    While Not odynLiqAux.EOF
        strCadFP = strCadFP & odynLiqAux("COD_FORMA_PAGO").Value & "|"
        strCadFPH = strCadFPH & odynLiqAux("COD_HIJO").Value & "|"
        strCadRetEfect = strCadRetEfect & odynLiqAux("FLG_RETIRO").Value & "|"
        
''        While Not odynLiquiDecl.EOF
''            '** Cambios hechos para las suma de remesa por moneda **'
''            '**              fecha de cambio 06/03/2007           **'
''            '*******************************************************'
''            'Dim x%
''           ' If Not odynLiquiSist.EOF Then
''            ' If grdSistema.ApproxCount > 0 Then
''                    Dim Y%
''                    For Y = 0 To odynLiquiDecl.RecordCount - 1
''                       If odynLiquiDecl("COD_CONCEPTO").Value = "01" Then
''                          strEfect_Val = Val(strEfect_Val) + Val(odynLiquiDecl("SUBTOTAL").Value)
''                          strCadEfectSol = Val(strCadEfectSol) + odynLiquiDecl("MONTO").Value
''                          strCadEfectSoles = Val(strCadEfectSoles) + Val(odynLiquiDecl("SUBTOTAL").Value)
''                       End If
''                       If odynLiquiDecl("COD_CONCEPTO").Value = "02" Then
''                          strEfect_Val = Val(strEfect_Val) + Val(odynLiquiDecl("SUBTOTAL").Value)
''                          strCadEfectDol = Val(strCadEfectDol) + odynLiquiDecl("MONTO").Value
''                          strCadEfectDolares = Val(strCadEfectDolares) + Val(odynLiquiDecl("SUBTOTAL").Value)
''                       End If
''                       odynLiquiDecl.MoveNext
''                    Next Y
''
''                    strCadEfect = Val(strCadEfectSol) & "|" & Val(strCadEfectDol)
''                    strCadEfectTotal = Val(strCadEfectSoles) & "|" & Val(strCadEfectDolares)
''
''                    'odynLiquiDecl.MoveNext
''            '  Else
''
''                    'strEfect_Val = "0"
''                    'strCadEfect = "0" & "|"
''                    'strCadEfectTotal = strCadEfectTotal & "0" & "|"
''
''            'End If
''        Wend
        
        If odynLiqAux("COD_FORMA_PAGO").Value = "001" And odynLiqAux("COD_HIJO").Value = "001" Then
            'If strEfect_Val > 0 Then
             '       strCadDif = strCadDif & (Val(strEfect_Val) - odynLiquiSist("TOTAL").Value) & "|"
                strCadSist = strCadSist & odynLiqAux("TOTAL").Value & "|"
             '  Else
             '       strCadDif = IIf(strCadDif = "", "0", strCadDif) & "0|"
              '      strCadSist = IIf(strCadSist = "", "0", strCadSist) & "0|"
            'End If
        ElseIf odynLiqAux("COD_FORMA_PAGO").Value = "001" And odynLiqAux("COD_HIJO").Value = "002" Then
            'If strEfect_Val > 0 Then
                strCadSist = strCadSist & odynLiqAux("TOTAL").Value & "|"
            '  Else
             '   strCadSist = IIf(strCadSist = "", "0", strCadSist) & "0|"
            'End If
        ElseIf (odynLiqAux("FLG_RETIRO").Value = "0") Or (odynLiqAux("FLG_RETIRO").Value = "1") Then
            strCadSist = strCadSist & odynLiqAux("TOTAL").Value & "|"
            
        End If
        odynLiqAux.MoveNext
    Wend
    
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/'
    '**** Valores para la cadena de remesa *****'
    '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/'
    odynLiquiDecl.MoveFirst
    While Not odynLiquiDecl.EOF
        '** Cambios hechos para las suma de remesa por moneda **'
        '**              fecha de cambio 06/03/2007           **'
        '*******************************************************'
        Dim y%
        For y = 0 To odynLiquiDecl.RecordCount - 1
           If odynLiquiDecl("COD_CONCEPTO").Value = "01" Then
              'strEfect_Val = Val(strEfect_Val) + Val(odynLiquiDecl("SUBTOTAL").Value)
              strCadCodRemesaP = strCadCodRemesaP & odynLiquiDecl("COD_FORMA_PAGO").Value & "|"
              strCadCodRemesaH = strCadCodRemesaH & odynLiquiDecl("COD_HIJO").Value & "|"
              'strCadEfectSol = Val(strCadEfectSol) + odynLiquiDecl("MONTO").Value & "|"
              'strCadEfectSoles = Val(strCadEfectSoles) + Val(odynLiquiDecl("SUBTOTAL").Value) & "|"
              strCadEfect = strCadEfect + odynLiquiDecl("MONTO").Value & "|"
              strCadEfectTotal = strCadEfectTotal + odynLiquiDecl("SUBTOTAL").Value & "|"
           End If
           If odynLiquiDecl("COD_CONCEPTO").Value = "02" Then
              'strEfect_Val = Val(strEfect_Val) + Val(odynLiquiDecl("SUBTOTAL").Value)
              strCadCodRemesaP = strCadCodRemesaP & odynLiquiDecl("COD_FORMA_PAGO").Value & "|"
              strCadCodRemesaH = strCadCodRemesaH & odynLiquiDecl("COD_HIJO").Value & "|"
              'strCadEfectDol = Val(strCadEfectDol) + odynLiquiDecl("MONTO").Value & "|"
              'strCadEfectDolares = Val(strCadEfectDolares) + Val(odynLiquiDecl("SUBTOTAL").Value) & "|"
              strCadEfect = strCadEfect + odynLiquiDecl("MONTO").Value & "|"
              strCadEfectTotal = strCadEfectTotal + odynLiquiDecl("SUBTOTAL").Value & "|"
           End If
           odynLiquiDecl.MoveNext
        Next y
        
        'strCadEfect = Val(strCadEfectSol) & "|" & Val(strCadEfectDol)
        'strCadEfectTotal = Val(strCadEfectSoles) & "|" & Val(strCadEfectDolares)
    Wend
    
    '*************************************************************'
    '-- Datos de Remesa para que no llegue nulos--'
'        Dim x%
'        odynLiquiSist.MoveFirst
'        For x = 1 To odynLiquiSist.RecordCount - 2
'            strCadCodRemesaP = strCadCodRemesaP & "0|"
'            strCadCodRemesaH = strCadCodRemesaH & "0|"
'            strCadEfect = strCadEfect & "|0"
'            strCadEfectTotal = strCadEfectTotal & "|0"
'        Next
'        strCadEfect = strCadEfect & "|"
'        strCadEfectTotal = strCadEfectTotal & "|"
    '*************************************************************'
    
    Dim gvarError As String
'    gvarError = objLiquidacion.Graba(gclsOracle.ODataBase, _
'                                     objUsuario.CodigoEmpresa, _
'                                     objUsuario.CodigoLocal, _
'                                     LblCaja.Caption, _
'                                     frm_Lista_Caja_PreCerradas.grdCajas.Columns("USU_TEC").Value, _
'                                     objUsuario.Codigo, _
'                                     strCadSist, _
'                                     "|", _
'                                     strCadFP, _
'                                     strCadFPH, _
'                                     strCadRetEfect, _
'                                     Trim(TxtRepoCaja.Text), _
'                                     "", _
'                                     Trim(LblCodLiquidacion.Caption), _
'                                     strCadCodRemesaP, _
'                                     strCadCodRemesaH, _
'                                     strCadEfect, _
'                                     strCadEfectTotal)
                         
    gvarError = objLiquidacion.Graba(gclsOracle.ODataBase, _
                                     objUsuario.CodigoEmpresa, _
                                     objUsuario.CodigoLocal, _
                                     LblCaja.Caption, _
                                     Cajero, _
                                     objUsuario.Codigo, _
                                     strCadSist, _
                                     "|", _
                                     strCadFP, _
                                     strCadFPH, _
                                     strCadRetEfect, _
                                     Trim(TxtRepoCaja.Text), _
                                     "", _
                                     Trim(LblCodLiquidacion.Caption), _
                                     strCadCodRemesaP, _
                                     strCadCodRemesaH, _
                                     strCadEfect, _
                                     strCadEfectTotal, _
                                     strCadCodDoc, _
                                     strCadMinDoc, _
                                     strCadMaxDoc)
                         
    If gvarError = "" Then
        MsgBox "Se realizo la Liquidación con exito", vbInformation, App.ProductName
        CodLiquidacion = Trim(LblCodLiquidacion.Caption)
        objImpresion.Imprime_Liquidacion objUsuario.CodigoEmpresa, _
                                         objUsuario.CodigoLocal, _
                                         LblCaja.Caption, _
                                         CodLiquidacion
                                         
        Set frm_Lista_Caja_PreCerradas.grdCajas.DataSource = objLiquidacion.ListaCajasPrecerradas(objUsuario.CodigoEmpresa, _
                                                                                                  objUsuario.CodigoLocal, _
                                                                                                  Format(frm_Lista_Caja_PreCerradas.dtpFchIni.Value, "dd/mm/yyyy"))
        Set odynLiquiSist = Nothing
        Set odynLiquiDecl = Nothing
        Set objImpresion = Nothing
        Unload Me
    Else
        MsgBox gvarError, vbCritical, App.ProductName
    End If
    
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_FORMA_PAGO", "COD_HIJO", _
                      "DES_FORMA_PAGO", "DES_HIJO", _
                      "FLG_EFEC", "MONTO", _
                      "IMP_TIPO_CAMBIO", "TOTAL")
    
    arrCaption = Array("Cod.FP", "Cod.FPH", _
                       "Des Padre", "Des Hijo", _
                       "Retiro Efect", "Monto", _
                       "Tip.Camb", "Total")
    
    arrAncho = Array(900, 900, _
                     1500, 3000, _
                     900, 900, _
                     900, 900)
    
    arrAlineacion = Array(dbgLeft, dbgLeft, _
                          dbgLeft, dbgLeft, _
                          dbgCenter, dbgRight, _
                          dbgRight, dbgRight)
    
    grdSistema.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    Dim i%
    For i = 0 To grdSistema.Columns.Count - 1
        grdSistema.Columns(i).Visible = False
    Next i
    
    grdSistema.Columns("DES_FORMA_PAGO").Visible = True
    grdSistema.Columns("DES_HIJO").Visible = True
    grdSistema.Columns("MONTO").Visible = True
    grdSistema.Columns("MONTO").NumberFormat = "####0.00"
    grdSistema.Columns("IMP_TIPO_CAMBIO").Visible = True
    grdSistema.Columns("IMP_TIPO_CAMBIO").NumberFormat = "####0.00"
    grdSistema.Columns("TOTAL").Visible = True
    grdSistema.Columns("TOTAL").NumberFormat = "####0.00"
    grdSistema.Columns("FLG_EFEC").Visible = True
    
    grdSistema.Columns("DES_FORMA_PAGO").Merge = True
    grdSistema.Columns("DES_HIJO").Merge = True
    'grdSistema.RowHeight = 1.2 * grdSistema.RowHeight
End Sub

Private Sub SeteaGrillaRem()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_CONCEPTO", "CONCEPTO", _
                      "MONTO", "IMP_TIPO_CAMBIO", _
                      "SUBTOTAL", "COD_FORMA_PAGO", _
                      "COD_HIJO")
    
    arrCaption = Array("Codigo", "Concepto", _
                       "Monto", "Tip.Camb", _
                       "Total", "FPago P", _
                       "FPago H")
    
    arrAncho = Array(900, 1500, _
                     900, 900, _
                     900, 900, _
                     900)
    
    arrAlineacion = Array(dbgLeft, dbgLeft, _
                          dbgRight, dbgRight, _
                          dbgRight, dbgLeft, _
                          dbgLeft)
    
    grdDeclarado.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    'grdDeclarado.RowHeight = 1.2 * grdDeclarado.RowHeight
    
    grdDeclarado.Columns("MONTO").NumberFormat = "####0.00"
    grdDeclarado.Columns("IMP_TIPO_CAMBIO").NumberFormat = "####0.00"
    grdDeclarado.Columns("SUBTOTAL").NumberFormat = "####0.00"
    
    grdDeclarado.Columns("COD_FORMA_PAGO").Visible = False
    grdDeclarado.Columns("COD_HIJO").Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set odynLiqAux = Nothing
End Sub

Private Sub grdSistema_DblClick()
    If grdSistema.ApproxCount <= 0 Then Exit Sub
    frm_VTA_Liquidacion_FP.pLiquidacion = Liquidacion
    frm_VTA_Liquidacion_FP.pCodFormaPago = grdSistema.Columns("COD_FORMA_PAGO").Value
    frm_VTA_Liquidacion_FP.pCodHijo = grdSistema.Columns("COD_HIJO").Value
    
    
    'frm_VTA_Liquidacion_FP.pCodHijo = grdSistema.Columns("COD_HIJO").Value
    frm_VTA_Liquidacion_FP.Show vbModal
End Sub



Private Sub TxtRepoCaja_Change()
    Dim dblRepCaja As Double
    Dim dblTotCobSist As Double
    Dim dblNueTotCob As Double
    Dim dblNueDif
    
    If Val(TxtRepoCaja.Text) > 0 Then
        dblRepCaja = Val(TxtRepoCaja.Text)
        dblTotCobSist = Val(dblTotSist)
        
        dblNueTotCob = Format(dblTotCobSist - dblRepCaja, "#,#,##0.00")
        LblTotCobrado.Caption = Format(dblTotCobSist - dblRepCaja, "#,#,##0.00")
        dblNueDif = Format(dblTotDecl - dblNueTotCob, "#,#,##0.00")
        LblDif.Caption = dblNueDif
      Else
        LblTotCobrado.Caption = Format(dblTotSist, "#,#,##0.00")
        LblDif.Caption = Format(dblTotDecl - dblTotSist, "#,###,##0.00")
    End If
End Sub

Private Sub TxtRepoCaja_KeyPress(KeyAscii As Integer)
    TxtRepoCaja.Tipo = Real
End Sub

Public Sub Mostrar(ByVal vblnCerrada As Boolean, ByVal vstrCaja As String, ByVal vstrLiquidacion As String)

    blnCerrada = vblnCerrada
    strCaja = vstrCaja
    strLiquidacion = vstrLiquidacion
    Me.Show vbModal

End Sub

Private Function Fecha() As String
    If blnCerrada Then
        Fecha = frm_Lista_Caja_PreCerradas.grdArqAnula.DataSource("FCH_INICIO").Value
    Else
        Fecha = frm_Lista_Caja_PreCerradas.grdCajas.Columns("FCH_INICIO").Value
    End If
End Function

Private Function Cajero() As String
    If blnCerrada Then
        Cajero = frm_Lista_Caja_PreCerradas.grdArqAnula.DataSource("USU_TEC").Value
    Else
        Cajero = frm_Lista_Caja_PreCerradas.grdCajas.Columns("USU_TEC").Value
    End If
End Function

Private Function Caja() As String
    If blnCerrada Then
        Caja = frm_Lista_Caja_PreCerradas.grdArqAnula.DataSource("COD_MAQUINA").Value
    Else
        Caja = frm_Lista_Caja_PreCerradas.grdCajas.Columns("COD_MAQUINA").Value
    End If
End Function

Private Function Liquidacion() As String
    If blnCerrada Then
        Liquidacion = frm_Lista_Caja_PreCerradas.grdArqAnula.DataSource("COD_LIQUIDACION").Value
    Else
        Liquidacion = frm_Lista_Caja_PreCerradas.grdCajas.Columns("COD_LIQUIDACION").Value
    End If
End Function

Private Sub DeshabilitaObjetos()

    If Not blnCerrada Then
        Exit Sub
    End If
    
    If TxtRepoCaja.Enabled Then
        TxtRepoCaja.Enabled = False
    End If
    If Not txtCerrada.Visible Then
        txtCerrada.Visible = True
    End If

    cmdGrabaVenta.Enabled = False
    Frame6.Enabled = False
End Sub
