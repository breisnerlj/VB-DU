VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_ADM_SustentoDiff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sustento de diferencias de inventario"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   10815
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   2640
      TabIndex        =   17
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   4920
         TabIndex        =   20
         Top             =   220
         Width           =   1095
      End
      Begin vbp_Ventas.ctlTextBox txtBuscar 
         Height          =   375
         Left            =   1080
         TabIndex        =   19
         Top             =   225
         Width           =   3735
         _ExtentX        =   6588
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
      Begin VB.CheckBox chkPendientes 
         Caption         =   "Solo pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6240
         TabIndex        =   18
         Top             =   320
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Buscar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   300
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   40
      TabIndex        =   1
      Top             =   720
      Width           =   10740
      Begin MSComDlg.CommonDialog CommonDialogExcel 
         Left            =   3120
         Top             =   3000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin vbp_Ventas.ctlGrilla grdDetImp 
         Height          =   1095
         Left            =   3120
         TabIndex        =   22
         Top             =   3000
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1931
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
      Begin ORADCLibCtl.ORADC oradcTipoAjuste 
         Height          =   255
         Left            =   3120
         Top             =   1080
         Visible         =   0   'False
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   207
         Caption         =   "oradcTipoAjuste"
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
         Bindings        =   "frm_ADM_SustentoDiff.frx":0000
         Height          =   1455
         Left            =   3120
         TabIndex        =   16
         Top             =   1320
         Width           =   2535
         _ExtentX        =   4471
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
      Begin VB.Frame Frame2 
         Caption         =   "Totales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2295
         Left            =   7440
         TabIndex        =   3
         Top             =   4680
         Width           =   3135
         Begin VB.Line Line1 
            Index           =   1
            X1              =   240
            X2              =   2880
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   225
            X2              =   2880
            Y1              =   1365
            Y2              =   1365
         End
         Begin VB.Label lblTotalXCobrar 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   1800
            TabIndex        =   13
            Top             =   1800
            Width           =   1005
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   1800
            TabIndex        =   12
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label lblTotalFalta 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            Height          =   195
            Left            =   1800
            TabIndex        =   11
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label lblTotalSobra 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            Height          =   195
            Left            =   1800
            TabIndex        =   10
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label lblTotalErrCont 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            Height          =   195
            Left            =   1800
            TabIndex        =   9
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total por cobrar:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   8
            Top             =   1800
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
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
            Index           =   3
            Left            =   240
            TabIndex        =   7
            Top             =   1440
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Faltantes:"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   6
            Top             =   1080
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobrantes:"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   5
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Error de conteo:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   1140
         End
      End
      Begin ORADCLibCtl.ORADC oradcMotivo 
         Height          =   255
         Left            =   6000
         Top             =   1080
         Visible         =   0   'False
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
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
      Begin TrueDBGrid70.TDBDropDown drdMotivo 
         Bindings        =   "frm_ADM_SustentoDiff.frx":001E
         Height          =   1455
         Left            =   6000
         TabIndex        =   14
         Top             =   1320
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2566
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Cod."
         Columns(0).DataField=   "COD_TIPO_SUSTENTO"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descripción"
         Columns(1).DataField=   "DES_TIPO_SUSTENTO"
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
      Begin vbp_Ventas.ctlGrillaArray grdCab 
         Height          =   4335
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7646
      End
      Begin vbp_Ventas.ctlGrillaArray grdDet 
         Height          =   2175
         Left            =   120
         TabIndex        =   15
         Top             =   4800
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   3836
      End
   End
   Begin vbp_Ventas.ctlToolBar ctlToolBar1 
      Height          =   600
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1058
      ModoBotones     =   6
      EnabledEfecto   =   0   'False
   End
End
Attribute VB_Name = "frm_ADM_SustentoDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private xCabSustento As New XArrayDB
Private xDetSustento As New XArrayDB
Private xTipoAjuste As New XArrayDB
Private strNumDocAlta As String
Private strNumDocBaja As String
Private intSortColumn As Integer
Private intSortOrder As Integer

Private Enum ColumnCab
    NUM_ITEM
    COD_BTL
    NUM_INVENTARIO
    COD_PRODUCTO
    DES_PRODUCTO
    DES_LABORATORIO
    CTD_FRACCIONAMIENTO
    CTD_PRODUCTO_FIS
    CTD_PRODUCTO_FRAC_FIS
    CTD_PRODUCTO_SIS
    CTD_PRODUCTO_FRAC_SIS
    CTD_DIFE_UNID
    CTD_DIFE_FRAC
    CTD_PEND_UNID
    CTD_PEND_FRAC
    VAL_UNITARIO
    VAL_TOTAL
    CTD_UNID
    CTD_FRAC
    COD_MOTIVO
    COD_TIPO_AJUSTE
    DES_OBS
    BTN_GRABAR
End Enum

Private Enum ColumnDet
    COD_BTL
    NUM_INVENTARIO
    COD_PRODUCTO
    COD_TIPO_SUSTENTO
    DES_TIPO_SUSTENTO
    CTD_SUSTENTO
    CTD_SUSTENTO_FRAC
    COD_TIPO_AJUSTE
    NUM_AJUSTE
    DES_OBSERVACION
End Enum

Private Sub SetGrid()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    Dim i As Integer
    Dim s As TrueDBGrid70.Split
    Dim c As Column
    
    On Error GoTo Control
    
    arrCampos = Array("ITEM", "COD_BTL", "NUM_INVENTARIO", "COD_PRODUCTO", "DES_PRODUCTO", _
                      "DES_LABORATORIO", "CTD_FRACCIONAMIENTO", "CTD_PRODUCTO_FIS", "CTD_PRODUCTO_FRAC_FIS", _
                      "CTD_PRODUCTO_SIS", "CTD_PRODUCTO_FRAC_SIS", "CTD_DIFE_UNID", "CTD_DIFE_FRAC", _
                      "CTD_PEND_UNID", "CTD_PEND_FRAC", "VAL_UNITARIO", "VAL_TOTAL", _
                      "CTD_UNID", "CTD_FRAC", "COD_MOTIVO", "COD_TIPO_AJUSTE", _
                      "DES_OBS", "BTN_GRABAR")
    arrCaption = Array("Item", "Local", "Inventario", "Código", "Descripcion", _
                       "Laboratorio", "Ctd. Frac.", "Unidades Físico", "Fracciones Físico", _
                       "Unidades Sistema", "Fracciones Sistema", "Unidades Diferencia", "Fracciones Diferencia", _
                       "Unidades Pendientes", "Fracciones Pendientes", "Valor por Unidad", "Total", _
                       "Unidades Sustento", "Fracciones Sustento", "Motivo", "Tipo Ajuste", _
                       "Observación", "")
    arrAncho = Array(500, 10, 10, 700, 4500, _
                     2500, 900, 900, 900, _
                     900, 900, 900, 900, _
                     900, 900, 900, 900, _
                     900, 900, 2000, 1500, _
                     2000, 1000)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, dbgCenter, dbgLeft, _
                          dbgLeft, dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgCenter, dbgRight, dbgRight, _
                          dbgCenter, dbgCenter, dbgCenter, dbgLeft, _
                          dbgCenter, dbgCenter)
    
    With grdCab
        .FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        .HeadLines = 2
        .EditorStyle.BackColor = vbWhite
        .EditorStyle.ForeColor = RGB(180, 0, 180)
        .EditorStyle.Font.Bold = True
        .AllowUpdate = True
        .RowHeight = 1.2 * .RowHeight
        .MarqueeStyle = 4
        .col = 0

        For i = 0 To .Columns.Count - 1
            .Columns(i).AllowSizing = False
            .Columns(i).WrapText = False
        Next i
        
        .Columns(ColumnCab.COD_BTL).Visible = False
        .Columns(ColumnCab.NUM_INVENTARIO).Visible = False
'''        .Columns(ColumnCab.CTD_DIFE_UNID).Visible = False
'''        .Columns(ColumnCab.CTD_DIFE_FRAC).Visible = False
        
        .Columns(ColumnCab.COD_MOTIVO).DropDown = drdMotivo
        .Columns(ColumnCab.COD_MOTIVO).AutoCompletion = True
        .Columns(ColumnCab.COD_MOTIVO).AutoDropDown = True
        .Columns(ColumnCab.COD_MOTIVO).DropDownList = True
       
        .Columns(ColumnCab.COD_TIPO_AJUSTE).DropDown = drdTipo
        .Columns(ColumnCab.COD_TIPO_AJUSTE).AutoCompletion = True
        .Columns(ColumnCab.COD_TIPO_AJUSTE).AutoDropDown = True
        .Columns(ColumnCab.COD_TIPO_AJUSTE).DropDownList = True
       
        .Columns(ColumnCab.CTD_PEND_UNID).BackColor = &HC0E0FF
        .Columns(ColumnCab.CTD_PEND_FRAC).BackColor = &HC0E0FF

        .Columns(ColumnCab.VAL_UNITARIO).NumberFormat = "Standard"
        .Columns(ColumnCab.VAL_TOTAL).NumberFormat = "Standard"

        'Columnas editables
        .Columns(ColumnCab.CTD_UNID).BackColor = vbInfoBackground
        .Columns(ColumnCab.CTD_UNID).DataWidth = 4
        .Columns(ColumnCab.CTD_FRAC).BackColor = vbInfoBackground
        .Columns(ColumnCab.CTD_FRAC).DataWidth = 4
        .Columns(ColumnCab.COD_MOTIVO).BackColor = vbInfoBackground
        .Columns(ColumnCab.COD_MOTIVO).DataWidth = 3
        .Columns(ColumnCab.COD_TIPO_AJUSTE).BackColor = vbInfoBackground
        .Columns(ColumnCab.COD_TIPO_AJUSTE).DataWidth = 3
        .Columns(ColumnCab.DES_OBS).BackColor = vbInfoBackground
        .Columns(ColumnCab.DES_OBS).DataWidth = 100
    
        .Columns(ColumnCab.BTN_GRABAR).ButtonText = True
        .Columns(ColumnCab.BTN_GRABAR).ButtonAlways = True
    
        .Rebind
    End With

    arrCampos = Array("COD_BTL", "NUM_INVENTARIO", "COD_PRODUCTO", _
                      "COD_TIPO_SUSTENTO", "DES_TIPO_SUSTENTO", "CTD_SUSTENTO", _
                      "CTD_SUSTENTO_FRAC", "COD_TIPO_AJUSTE", "NUM_AJUSTE", _
                      "DES_OBSERVACION")
    arrCaption = Array("", "", "", _
                       "Código", "Descripcion", "Unidades", _
                       "Fracciones", "Tipo Ajuste", "Numero Ajuste", _
                       "Observaciones")
    arrAncho = Array(100, 100, 100, _
                     700, 3000, 1000, _
                     1000, 1000, 1500, _
                     3000)
    arrAlineacion = Array(dbgCenter, dbgCenter, dbgCenter, _
                          dbgCenter, dbgLeft, dbgCenter, _
                          dbgCenter, dbgCenter, dbgCenter, _
                          dbgLeft)
    
    With grdDet
        .FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
        .HeadLines = 2
        .RowHeight = 1.2 * .RowHeight
        .MarqueeStyle = 4
        .col = 0

        For i = 0 To .Columns.Count - 1
            .Columns(i).AllowSizing = False
            .Columns(i).WrapText = False
        Next i
        
        .Columns(ColumnDet.COD_BTL).Visible = False
        .Columns(ColumnDet.NUM_INVENTARIO).Visible = False
        .Columns(ColumnDet.COD_PRODUCTO).Visible = False
    
        .Rebind
        
    End With
    
    Exit Sub
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub CargaMotivos()
    On Error GoTo Control
    
    With drdMotivo
        .RowHeight = 0
        .RowHeight = .RowHeight * 1.2
        .DataField = "COD_TIPO_SUSTENTO"
        .ListField = "DES_TIPO_SUSTENTO"
        .AllowRowSizing = False
        .AllowColMove = False
        .EmptyRows = False
        .Appearance = dbgFlat
        .ValueTranslate = True
    End With
    
    Set oradcMotivo.Recordset = gclsOracle.FN_Cursor("BTLPROD.PKG_SUSTENTO_DIF.FN_LISTA_TIPO_SUSTENTO", 0, vbNullString, "1")
    
    Exit Sub
Control:
    Set oradcMotivo.Recordset = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub chkPendientes_Click()
    On Error GoTo Control
    
    Call CargarDiferencias

    Exit Sub
Control:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo Control
    
    Call CargarDiferencias

    Exit Sub
Control:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub ctlToolBar1_Click(ByVal boton As tlbTipoBoton, ByVal Index As Integer)
    Dim objDocumento As clsDocumento
    Dim strGraba As String
    
    On Error GoTo Control
    
    Select Case boton
        Case tlbTipoBoton.tb_Excel
            Call ExportarAExcel
        Case tlbTipoBoton.salir
            If MsgBox("Esta acción cerrará el formulario sin confirmar el sustento. " & vbCrLf & _
                      "Las modificaciones realizadas no se perderán y podrá continuar luego." & vbCrLf & _
                      "¿Está seguro que desea contiuar?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                Unload Me
            End If
        Case tlbTipoBoton.Grabar
            ''arturo escate 20/11/2009 se caia cuando no tenia ningun producto
            If grdCab.ApproxCount <= 0 Then MsgBox "No tiene datos para grabar", vbCritical, App.ProductName: Exit Sub
            If MsgBox("Esta acción confirmará el sustento y no podrá ser modificado luego." & vbCrLf & _
                      "¿Está seguro que desea contiuar?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                
                strGraba = GrabarSustento
                If strGraba = vbNullString Then
                    MsgBox "Se guardaron los cambios satisfactoriamente." & vbCrLf & _
                           "Le llegará una comunicación a inventarios indicando el término de la revisión del sustento.", _
                           vbInformation, App.ProductName
                    Set objDocumento = New clsDocumento
                    
                    If strNumDocAlta <> vbNullString Then
                        MsgBox "Se grabó el AJUSTE Nº " & strNumDocAlta & " para las altas." & vbCrLf & _
                               "Sirvase verificar el formato de guia en la impresora.", vbInformation, App.ProductName
                        objDocumento.Imprime_Ajuste_Cje strNumDocAlta
                    End If
                    
                    If strNumDocBaja <> vbNullString Then
                        MsgBox "Se grabó el AJUSTE Nº " & strNumDocBaja & " para las bajas." & vbCrLf & _
                               "Sirvase verificar el formato de guia en la impresora.", vbInformation, App.ProductName
                        objDocumento.Imprime_Ajuste_Cje strNumDocBaja
                    End If
                    
                    Set objDocumento = Nothing
                    Unload Me
                Else
                    MsgBox strGraba, vbCritical, App.ProductName
                    Exit Sub
                End If
            End If
    End Select

    Exit Sub
Control:
    objDocumento = Nothing
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_Load()
    On Error GoTo Control
    
    ctlToolBar1.Buttons(7).Visible = True
    ctlToolBar1.Buttons(11).Visible = False
        
    intSortOrder = XORDER_ASCEND
    intSortColumn = ColumnCab.COD_PRODUCTO
    
    Call CargarTiposAjuste
    Call CargaMotivos
    Call SetGrid
    Call CargarDiferencias
    Call CargarTotales
    
    Exit Sub
Control:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub CargarTiposAjuste()
    Dim objAjuste As clsAjuste
    
    On Error GoTo Control
    
    Set objAjuste = New clsAjuste
    With drdTipo
        .RowHeight = 0
        .RowHeight = .RowHeight * 1.2
        .DataField = "COD_TIP_AJUSTE"
        .ListField = "DES_TIP_AJUSTE"
        .AllowRowSizing = False
        .AllowColMove = False
        .EmptyRows = False
        .Appearance = dbgFlat
        .ValueTranslate = True
    End With
    
    Set oradcTipoAjuste.Recordset = objAjuste.ListaTipoAju
    
    Set objAjuste = Nothing
    
    Exit Sub
Control:
    Set oradcTipoAjuste.Recordset = Nothing
    Set objAjuste = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub CargarDiferencias()
    Dim rsCab As oraDynaset
    Dim i As Integer
    
    On Error GoTo Control
    
    Me.MousePointer = vbHourglass
    
    xCabSustento.ReDim 0, -1, 0, 22

    Set rsCab = gclsOracle.FN_Cursor("BTLPROD.PKG_SUSTENTO_DIF.FN_LISTA_DIFERENCIA", 0, _
                                     objUsuario.CodigoLocal, _
                                     vbNullString, _
                                     vbNullString, _
                                     CInt(chkPendientes.Value), _
                                     Trim(txtBuscar.Text), _
                                     intSortColumn, _
                                     IIf(intSortOrder = XORDER_DESCEND, "DESC", vbNullString))
                                     
    If rsCab.RecordCount > 0 Then xCabSustento.LoadRows rsCab.GetRows
    
    If xCabSustento.UpperBound(2) < ColumnCab.BTN_GRABAR Then xCabSustento.AppendColumns ColumnCab.BTN_GRABAR - xCabSustento.UpperBound(2)
    
    For i = xCabSustento.LowerBound(1) To xCabSustento.UpperBound(1)
        xCabSustento.Value(i, ColumnCab.BTN_GRABAR) = "Grabar"
    Next i
    
    grdCab.Array1 = xCabSustento
    
    Set rsCab = Nothing

    Me.MousePointer = vbDefault
    
    Exit Sub
Control:
    Me.MousePointer = vbDefault
    Set rsCab = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub CargarDetalle()
    Dim rsDet As oraDynaset
    
    On Error GoTo Control
    
    xDetSustento.ReDim 0, -1, 0, 4
    
    Set rsDet = gclsOracle.FN_Cursor("BTLPROD.PKG_SUSTENTO_DIF.FN_LISTA_DET_SUSTENTO_DIF", 0, _
                                                     grdCab.Columns(ColumnCab.COD_BTL).Value, _
                                                     grdCab.Columns(ColumnCab.NUM_INVENTARIO).Value, _
                                                     grdCab.Columns(ColumnCab.COD_PRODUCTO).Value)

    If rsDet.RecordCount > 0 Then xDetSustento.LoadRows rsDet.GetRows

    grdDet.Array1 = xDetSustento
    
    Set rsDet = Nothing

    Exit Sub
Control:
    Set rsDet = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub CargarTotales()
    Dim rsTot As oraDynaset
    
    On Error GoTo Control
    
    Set rsTot = gclsOracle.FN_Cursor("BTLPROD.PKG_SUSTENTO_DIF.FN_LISTA_TOT_SUSTENTO_DIF", 0, _
                                                     grdCab.Columns(ColumnCab.COD_BTL).Value, _
                                                     grdCab.Columns(ColumnCab.NUM_INVENTARIO).Value)

    If rsTot.RecordCount > 0 Then

        lblTotalSobra.Caption = Format$(CDbl(rsTot.Fields("VAL_SOBRANTES").Value), "##,##0.00")
        lblTotalFalta.Caption = Format$(CDbl(rsTot.Fields("VAL_FALTANTES").Value), "##,##0.00")
        lblTotalErrCont.Caption = Format$(CDbl(rsTot.Fields("VAL_ERR_CONTEO").Value), "##,##0.00")
        'lblTotalCruce
        lblTotal.Caption = Format$(CDbl(rsTot.Fields("VAL_TOTAL").Value), "##,##0.00")
        lblTotalXCobrar.Caption = Format$(CDbl(rsTot.Fields("VAL_X_COBRAR").Value), "##,##0.00")
    
    Else
    
        lblTotalSobra.Caption = "0.00"
        lblTotalFalta.Caption = "0.00"
        lblTotalErrCont.Caption = "0.00"
        lblTotal.Caption = "0.00"
        lblTotalXCobrar.Caption = "0.00"
    
    End If
    
    Set rsTot = Nothing

    Exit Sub
Control:
    Set rsTot = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    On Error GoTo Control
    
    If UnloadMode = vbFormControlMenu Then
        If MsgBox("Esta acción cerrará el formulario sin confirmar el sustento. " & vbCrLf & _
                  "Las modificaciones realizadas no se perderán y podrá continuar luego." & vbCrLf & _
                  "¿Está seguro que desea contiuar?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then
            Cancel = 1
        End If
    End If

    Exit Sub
Control:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo Control
    
    xCabSustento.ReDim 0, -1, 0, 21
    xDetSustento.ReDim 0, -1, 0, 10

    Exit Sub
Control:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdCab_AfterColUpdate(ByVal ColIndex As Integer)
    grdCab.MoveNext
    grdCab.MovePrevious
End Sub

Private Sub grdCab_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    Select Case ColIndex
        Case ColumnCab.COD_BTL To ColumnCab.CTD_PEND_FRAC
            Cancel = 1
    End Select
End Sub

Private Sub grdCab_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim intUndPendientes As Double
    Dim intFraPendientes As Double
    Dim strCodMotAjuste As String
    Dim strCodMotivo As String

    On Error GoTo Control
    
    intUndPendientes = Val("" & xCabSustento.Value(grdCab.Bookmark, ColumnCab.CTD_PEND_UNID))
    intFraPendientes = Val("" & xCabSustento.Value(grdCab.Bookmark, ColumnCab.CTD_PEND_FRAC))
    strCodMotAjuste = "" & xCabSustento.Value(grdCab.Bookmark, ColumnCab.COD_MOTIVO)
    strCodMotivo = "" & grdCab.Columns(ColumnCab.COD_MOTIVO).Value

    Select Case ColIndex
        Case ColumnCab.CTD_UNID, ColumnCab.CTD_FRAC
            If grdCab.Columns(ColIndex).Value <> "" And Not IsNumeric("" & grdCab.Columns(ColIndex).Value) Then
                MsgBox "El valor no es válido.", vbCritical, App.ProductName
                Cancel = 1
                Exit Sub
            End If
                
            If Val("" & grdCab.Columns(ColIndex).Value) < 0 Then
                MsgBox "No se aceptan valores negativos.", vbCritical, App.ProductName
                Cancel = 1
                Exit Sub
            End If
'''            If ColIndex = ColumnCab.CTD_UNID Then
'''                If intUndPendientes = 0 And Val("" & grdCab.Columns(ColIndex).Value) <> 0 Then
'''                    MsgBox "No existen unidades por sustentar.", vbCritical, App.ProductName
'''                    Cancel = 1
'''                    Exit Sub
'''                End If
'''            Else
'''                If intFraPendientes = 0 And Val("" & grdCab.Columns(ColIndex).Value) <> 0 Then
'''                    MsgBox "No existen fracciones por sustentar.", vbCritical, App.ProductName
'''                    Cancel = 1
'''                    Exit Sub
'''                End If
'''            End If
    End Select

    Exit Sub
Control:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdCab_ButtonClick(ByVal ColIndex As Integer)
    Dim arrValores As Variant
    Dim arrDireccion As Variant
    Dim strAgregaSus As String
    Dim strMotivoAjuste As String
    Dim intCtdDifeUnid As Double, intCtdDifeFrac As Double
    Dim intUndPendientes As Double, intFraPendientes As Double
    Dim intCtdUnidades As Double, intCtdFraccion As Double
    Dim strCodMotivoS As String, strCodTipoAju As String
    
    On Error GoTo Control
    
    If ColIndex = ColumnCab.BTN_GRABAR Then
        intCtdDifeUnid = Val("" & xCabSustento.Value(grdCab.Bookmark, ColumnCab.CTD_DIFE_UNID))
        intCtdDifeFrac = Val("" & xCabSustento.Value(grdCab.Bookmark, ColumnCab.CTD_DIFE_FRAC))
        intUndPendientes = Val("" & xCabSustento.Value(grdCab.Bookmark, ColumnCab.CTD_PEND_UNID))
        intFraPendientes = Val("" & xCabSustento.Value(grdCab.Bookmark, ColumnCab.CTD_PEND_FRAC))
        intCtdUnidades = Val("" & xCabSustento.Value(grdCab.Bookmark, ColumnCab.CTD_UNID))
        intCtdFraccion = Val("" & xCabSustento.Value(grdCab.Bookmark, ColumnCab.CTD_FRAC))
        strCodMotivoS = "" & xCabSustento.Value(grdCab.Bookmark, ColumnCab.COD_MOTIVO)
        strCodTipoAju = "" & xCabSustento.Value(grdCab.Bookmark, ColumnCab.COD_TIPO_AJUSTE)

        If intCtdUnidades = 0 And intCtdFraccion = 0 Then
            MsgBox "Ingrese las cantidades y/o fracciones a sustentar.", vbCritical, App.ProductName
            Exit Sub
        End If

        If strCodMotivoS = vbNullString Then
            MsgBox "Debe seleccionar el motivo de sustento.", vbCritical, App.ProductName
            Exit Sub
        End If

        If strCodMotivoS = gclsOracle.Const_Val("BTLPROD.PKG_SUSTENTO_DIF.CONS_TIPO_SOBRANTE") Then
            If (intUndPendientes < 0) Or (intFraPendientes < 0) Then
                MsgBox "Las diferencias negativas solo se pueden sustentar como FALTANTES o ERRORES DE CONTEO.", _
                    vbCritical, App.ProductName
                Exit Sub
            End If
'''            If (intFraPendientes < 0) Then
'''                MsgBox "Las diferencias negativas solo se pueden sustentar como FALTANTES o ERRORES DE CONTEO.", _
'''                    vbCritical, App.ProductName
'''                Exit Sub
'''            End If
        End If

        If strCodMotivoS = gclsOracle.Const_Val("BTLPROD.PKG_SUSTENTO_DIF.CONS_TIPO_FALTANTE") Then
            If (intUndPendientes > 0) Or (intFraPendientes > 0) Then
                MsgBox "Las diferencias positivas solo se pueden sustentar como SOBRANTES o ERRORES DE CONTEO.", _
                    vbCritical, App.ProductName
                Exit Sub
            End If
'''            If (intFraPendientes > 0) Then
'''                MsgBox "Las diferencias positivas solo se pueden sustentar como SOBRANTES o ERRORES DE CONTEO.", _
'''                    vbCritical, App.ProductName
'''                Exit Sub
'''            End If
        End If

'''        If strCodMotivoS <> gclsOracle.Const_Val("BTLPROD.PKG_SUSTENTO_DIF.CONS_TIPO_ERROR_CONTEO") Then
'''            If Abs(intUndPendientes) < intCtdUnidades Then
'''                MsgBox "No puede sustentar mas unidades que las pendientes ( " & Abs(intUndPendientes) & " ).", _
'''                    vbCritical, App.ProductName
'''                Exit Sub
'''            End If
'''        End If
'''
'''        If strCodMotivoS <> gclsOracle.Const_Val("BTLPROD.PKG_SUSTENTO_DIF.CONS_TIPO_ERROR_CONTEO") Then
'''            If Abs(intFraPendientes) < intCtdFraccion Then
'''                MsgBox "No puede sustentar mas fracciones que las pendientes ( " & Abs(intFraPendientes) & " ).", _
'''                    vbCritical, App.ProductName
'''                Exit Sub
'''            End If
'''        End If

        If strCodMotivoS = gclsOracle.Const_Val("BTLPROD.PKG_SUSTENTO_DIF.CONS_TIPO_ERROR_CONTEO") Then
            If strCodTipoAju = vbNullString Then
                MsgBox "Debe seleccionar el tipo de ajuste a realizar.", vbCritical, App.ProductName
                Exit Sub
            End If
        End If
        
        arrValores = Array(Trim(xCabSustento.Value(grdCab.Bookmark, ColumnCab.COD_BTL)), _
                           Trim(xCabSustento.Value(grdCab.Bookmark, ColumnCab.NUM_INVENTARIO)), _
                           Trim(xCabSustento.Value(grdCab.Bookmark, ColumnCab.COD_PRODUCTO)), _
                           Trim(xCabSustento.Value(grdCab.Bookmark, ColumnCab.COD_MOTIVO)), _
                           Trim("" & xCabSustento.Value(grdCab.Bookmark, ColumnCab.CTD_UNID)), _
                           Trim("" & xCabSustento.Value(grdCab.Bookmark, ColumnCab.CTD_FRAC)), _
                           xCabSustento.Value(grdCab.Bookmark, ColumnCab.DES_OBS), _
                           strCodTipoAju, _
                           objUsuario.Codigo)
        
        arrDireccion = Array(entrada, entrada, _
                             entrada, entrada, _
                             entrada, entrada, _
                             entrada, entrada, _
                             entrada)
                             
        strAgregaSus = gclsOracle.SP("BTLPROD.PKG_SUSTENTO_DIF.SP_AGREGAR_DETALLE_SUSTENTO", arrValores, arrDireccion)
        
        If strAgregaSus = vbNullString Then
            Call CargarDiferencias
            Call CargarDetalle
            Call CargarTotales
        Else
            MsgBox strAgregaSus, vbCritical, App.ProductName
        End If
    
    End If
    
    Exit Sub
Control:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdCab_HeadClick(ByVal ColIndex As Integer)
    Dim i As Integer
    
    On Error GoTo Control

    Select Case ColIndex
        Case ColumnCab.COD_PRODUCTO, _
             ColumnCab.DES_PRODUCTO, _
             ColumnCab.DES_LABORATORIO, _
             ColumnCab.VAL_TOTAL
            
            Me.MousePointer = vbHourglass
            
            If intSortColumn <> ColIndex Then
                intSortColumn = ColIndex
                intSortOrder = XORDER_ASCEND
            Else
                intSortOrder = IIf(intSortOrder = XORDER_DESCEND, XORDER_ASCEND, XORDER_DESCEND)
            End If
            
            Call CargarDiferencias
            
            Me.MousePointer = vbDefault
    End Select

    Exit Sub
Control:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdCab_RegistroSeleccionado(ByVal DatoColumna0 As String)
    On Error GoTo Control
    
    Call CargarDetalle
    
    Exit Sub
Control:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Function GrabarSustento() As String
    Dim arrValores As Variant
    Dim arrDireccion As Variant
        
    On Error GoTo Control
    
    arrValores = Array(Trim(xCabSustento.Value(xCabSustento.LowerBound(1), ColumnCab.COD_BTL)), _
                       Trim(xCabSustento.Value(xCabSustento.LowerBound(1), ColumnCab.NUM_INVENTARIO)), _
                       lblTotalErrCont.Caption, _
                       lblTotalSobra.Caption, _
                       lblTotalFalta.Caption, _
                       lblTotal.Caption, _
                       lblTotalXCobrar.Caption, _
                       objUsuario.Codigo, _
                       strNumDocAlta, _
                       strNumDocBaja)
        
    arrDireccion = Array(entrada, entrada, _
                         entrada, entrada, _
                         entrada, entrada, _
                         entrada, entrada, _
                         Salida, Salida)
                             
    GrabarSustento = gclsOracle.SP("BTLPROD.PKG_SUSTENTO_DIF.SP_GRABA", arrValores, arrDireccion)
    
    Exit Function
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub grdDet_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim arrValores As Variant
    Dim arrDireccion As Variant
    Dim strQuitarItem As String
    
    On Error GoTo Control
    
    If KeyCode = vbKeyDelete Then
        arrValores = Array(Trim(xDetSustento.Value(grdDet.Bookmark, ColumnDet.COD_BTL)), _
                           Trim(xDetSustento.Value(grdDet.Bookmark, ColumnDet.NUM_INVENTARIO)), _
                           Trim(xDetSustento.Value(grdDet.Bookmark, ColumnDet.COD_PRODUCTO)), _
                           Trim(xDetSustento.Value(grdDet.Bookmark, ColumnDet.COD_TIPO_SUSTENTO)))
            
        arrDireccion = Array(entrada, entrada, entrada, entrada)
                                 
        strQuitarItem = gclsOracle.SP("BTLPROD.PKG_SUSTENTO_DIF.SP_QUITAR_DETALLE_SUSTENTO", arrValores, arrDireccion)
        
        If strQuitarItem = vbNullString Then
            Call CargarDiferencias
            Call CargarDetalle
            Call CargarTotales
        Else
            MsgBox strQuitarItem, vbCritical, App.ProductName
        End If
    End If

    Exit Sub
Control:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub txtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdBuscar_Click
    End If
End Sub

Private Sub ExportarAExcel()
    Dim lstrFile As String
    
    On Error GoTo Control
    
    Set grdDetImp.DataSource = gclsOracle.FN_Cursor("BTLPROD.PKG_SUSTENTO_DIF.FN_REPO_DET_SUSTENTO_DIF", 0, _
                                                    grdCab.Columns(ColumnCab.COD_BTL).Value, _
                                                    grdCab.Columns(ColumnCab.NUM_INVENTARIO).Value)
    
    If grdDetImp.ApproxCount < 1 Then
        MsgBox "No hay datos para exportar", vbCritical, App.ProductName
        Exit Sub
    End If

    With CommonDialogExcel
        .DefaultExt = ".xls"
        .Filter = "Archivos de Excel (*.xls)|*.xls|"
        .FilterIndex = 1
        .ShowSave
        lstrFile = .FileName
    End With
    
    If lstrFile = "" Then Exit Sub

    grdDetImp.ExportToFile lstrFile, False
    
    MsgBox "Los datos fueron exportados correctamente", vbOKOnly + vbInformation, App.ProductName
    
    Set grdDetImp.DataSource = Nothing
    
    Exit Sub
Control:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
