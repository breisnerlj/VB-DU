VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_ADM_HistSolicitud 
   BorderStyle     =   0  'None
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10515
   Icon            =   "frm_ADM_HistSolicitud.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin ORADCLibCtl.ORADC oradcHistSolicitudes 
      Height          =   270
      Left            =   2280
      Top             =   8040
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   476
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
   Begin TrueDBGrid70.TDBGrid grdHistSolicitudes 
      Bindings        =   "frm_ADM_HistSolicitud.frx":0442
      Height          =   4725
      Left            =   60
      TabIndex        =   0
      Top             =   1560
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   8334
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
      Enabled         =   0   'False
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      EmptyRows       =   -1  'True
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=240,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(9)   =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(10)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0"
      _StyleDefs(42)  =   "Named:id=34:Heading"
      _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=34,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
   End
   Begin MSComctlLib.ImageList IlsImagen 
      Left            =   7320
      Top             =   -120
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
            Picture         =   "frm_ADM_HistSolicitud.frx":0465
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistSolicitud.frx":09FF
            Key             =   "buscar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistSolicitud.frx":0F99
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistSolicitud.frx":1533
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistSolicitud.frx":1ACD
            Key             =   "Chek"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistSolicitud.frx":2067
            Key             =   "Bien"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistSolicitud.frx":2601
            Key             =   "Agregar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistSolicitud.frx":2B9B
            Key             =   "Atras"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistSolicitud.frx":3135
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistSolicitud.frx":36CF
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistSolicitud.frx":3C69
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_HistSolicitud.frx":4203
            Key             =   "Hora"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   12
      Top             =   6570
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1111
      ButtonWidth     =   1270
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "IlsImagen"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nueva"
            Object.ToolTipText     =   "Nueva Solicitud"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Anular"
            Object.ToolTipText     =   "Anular Solicitud"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Visuali."
            Object.ToolTipText     =   "Visualizar Solicitud"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Recep."
            Object.ToolTipText     =   "Recepcionar Pedido"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Hist."
            Object.ToolTipText     =   "Ver &Historial"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Key             =   "Salir"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Servicios Generales-Historial de Solicitudes"
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
      TabIndex        =   13
      Top             =   120
      Width           =   4710
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   0
      Picture         =   "frm_ADM_HistSolicitud.frx":479D
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Height          =   195
      Index           =   0
      Left            =   5160
      TabIndex        =   11
      ToolTipText     =   "by Nelson Chumbipuma"
      Top             =   120
      Width           =   225
   End
   Begin VB.Label lblFechaFinal 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblFechaInicial 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblTitFin 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Fin:"
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
      Left            =   4800
      TabIndex        =   8
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label lblTitInicio 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Inicio:"
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
      Left            =   2160
      TabIndex        =   7
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label lblPeriodo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   180
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblCodArea 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1095
      Width           =   735
   End
   Begin VB.Label lblDesArea 
      BackColor       =   &H80000005&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   615
      Width           =   5895
   End
   Begin VB.Label lblArea 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "AREA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   1080
      Width           =   615
   End
   Begin VB.Image imgLogo 
      Height          =   660
      Left            =   8640
      Picture         =   "frm_ADM_HistSolicitud.frx":6497
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label lblFondo2 
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   10400
   End
   Begin VB.Label lblMensaje 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frm_ADM_HistSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lodynConsulta As oraDynaset
Dim lodynConsultaClone As oraDynaset
Dim lstrRetCodigo As String
Dim lstrNumSolicitud As String
Dim lstrEstado As String
Dim lstrPaseInicio As String
Dim lstrFechaInicial As String
Dim lstrFechaFinal As String
Dim lstrPaseAnular As String
Dim lstrFechaActual As String
Public lstrDeNueva As String
Dim lstrCodPeriodo As String
Dim lblnMostrar As Boolean
Dim objSSGG As New clsSSGG

Private Sub CmdAnular_Click()
    On Error GoTo ERROR
    If Trim(lstrNumSolicitud) = "" Then
        MsgBox "No existe Solicitud a anular", vbInformation, "Aviso"
        Exit Sub
    End If
    If lstrEstado = "ANU" Then
        MsgBox "La Solicitud ya está anulada", vbInformation, "Aviso"
        Exit Sub
    End If
    If grdHistSolicitudes.Columns(9).Text <> gstrCodUsuario Then
        MsgBox "Esta solicitud solo puede ser anulada por el usuario que la emitió!", vbExclamation, "Aviso"
        Exit Sub
    End If
        
    If MsgBox("Desea anular la solicitud?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If





    Dim lstrError As String
    Dim lstrMensaje As String
    
    Screen.MousePointer = vbHourglass
    Dim gvarValores  As Variant
    Dim gvarIO  As Variant
                              
     'Add RET_ERROR as an Output parameter and set its initial value.
    godbVentas.Parameters.Add "RET_ERROR", "0", ORAPARM_OUTPUT
    godbVentas.Parameters("RET_ERROR").serverType = ORATYPE_VARCHAR2

     'Add RET_ACCION as an Output parameter and set its initial value.
    godbVentas.Parameters.Add "RET_MENSAJE", "0", ORAPARM_OUTPUT
    godbVentas.Parameters("RET_MENSAJE").serverType = ORATYPE_VARCHAR2

    'Add NUM_SOLICITUD as an Input/Output parameter and set its initial value.
    godbVentas.Parameters.Add "NUM_SOLICITUD", lstrNumSolicitud, ORAPARM_INPUT
    godbVentas.Parameters("NUM_SOLICITUD").serverType = ORATYPE_VARCHAR2

    'Add NUM_SOLICITUD as an Input/Output parameter and set its initial value.
    godbVentas.Parameters.Add "COD_USUARIO", Trim(gstrCodUsuario), ORAPARM_INPUT
    godbVentas.Parameters("COD_USUARIO").serverType = ORATYPE_VARCHAR2

    'Add NUM_SOLICITUD as an Input/Output parameter and set its initial value.
    godbVentas.Parameters.Add "COD_AREA", Trim(gstrCodAreaUsuario), ORAPARM_INPUT
    godbVentas.Parameters("COD_AREA").serverType = ORATYPE_VARCHAR2

    'Execute the Stored Procedure.
    godbVentas.ExecuteSQL ("Begin SSGG.SP_SGE_ANU_SOLICITUD (:RET_ERROR, :RET_MENSAJE, :NUM_SOLICITUD, :COD_USUARIO, :COD_AREA); end;")

    lstrError = godbVentas.Parameters("RET_ERROR").Value
    lstrMensaje = IIf(IsNull(godbVentas.Parameters("RET_MENSAJE").Value), "", godbVentas.Parameters("RET_MENSAJE").Value)

    godbVentas.Parameters.Remove "RET_ERROR"
    godbVentas.Parameters.Remove "RET_MENSAJE"
    godbVentas.Parameters.Remove "NUM_SOLICITUD"
    godbVentas.Parameters.Remove "COD_USUARIO"
    godbVentas.Parameters.Remove "COD_AREA"

    If Trim(lstrError) = "0" Then
        MsgBox "Solicitud Anulada", vbInformation, "Aviso"
    Else
        MsgBox Trim(lstrMensaje), vbCritical, "Error"
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    Call sub_sge_Llena_Grilla
    Screen.MousePointer = vbNormal
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdNueva_Click()
    On Error GoTo ERROR
    lstrDeNueva = "si"
    frm_ADM_Solicitud.sub_sge_Generar_Solicitud
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVisualizar_Click()
    On Error GoTo ERROR
    If lstrNumSolicitud <> "" Then
        frm_ADM_Solicitud.blnNuevaHistorial = True
        frm_ADM_Solicitud.sub_sge_Visualizar_Solicitud (lstrNumSolicitud)
    Else
        MsgBox "No existe Solicitud a visualizar", vbInformation, "Aviso"
    End If
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub cmdRecepcionar_Click()
    On Error GoTo ERROR
    frm_ADM_RecepcionSolicitudes.Show
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub cmdVer_Click()
    On Error GoTo ERROR
    Screen.MousePointer = vbHourglass
    grdHistSolicitudes.Enabled = True
    lstrPaseInicio = "no"
    Call sub_sge_Llena_Grilla
    Screen.MousePointer = vbNormal
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
        frm_ADM_RecepcionSolicitudes.Show
    Case vbKeyEscape
        mdiPrincipal.picComandos.Enabled = True
        Unload Me
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub sub_sge_llena_fechas()
    On Error GoTo ERROR
    Dim lodynConsultaFechas As oraDynaset
    Dim strSqlFechas  As String
    Screen.MousePointer = vbHourglass
    '-Correccion
    'strSqlFechas = " SELECT COD_PERIODO,FCH_INICIO,NVL(FCH_FIN,'01/01/2000') AS FECFIN" & _
                   " FROM SSGG.AUX_PERIODO_PEDIDO WHERE COD_PERIODO = (SELECT MAX(COD_PERIODO) FROM SSGG.AUX_PERIODO_PEDIDO )"
    'Set lodynConsultaFechas = godbVentas.CreateDynaset(strSqlFechas, 0&)
    Set lodynConsultaFechas = objSSGG.llenaFechas
    '---------------------------------------------------
    If lodynConsultaFechas.RecordCount <> 0 Then
        lstrCodPeriodo = Trim(lodynConsultaFechas("COD_PERIODO").Value)
        lblFechaInicial.Caption = lodynConsultaFechas("FCH_INICIO").Value
        If Trim(lodynConsultaFechas("FCH_INICIO").Value) <> "" And Trim(lodynConsultaFechas("FECFIN").Value) = "01/01/2000" Then
            'lblFechaFinal.Caption = "  /  /    "
            '===================================
            Dim lodynFecFinDefault As oraDynaset
            Dim strSqlFecFinDefault As String
            'strSqlFecFinDefault = " SELECT VAL_PARAMETRO" & _
                                  " FROM NUEVO.MAE_PARAMETRO WHERE COD_PARAMETRO = 'NUM_DIASSG'"
            Set lodynFecFinDefault = objSSGG.Parametro
                                  
'            Set lodynFecFinDefault = godbVentas.CreateDynaset(strSqlFecFinDefault, 0&)
            If lodynFecFinDefault.RecordCount <> 0 Then
                lblFechaFinal.ForeColor = RGB(253, 254, 226)
                lblFechaFinal.Caption = lodynConsultaFechas("FCH_INICIO").Value + CInt(Trim(lodynFecFinDefault("VAL_PARAMETRO").Value))
            Else
                lblFechaFinal.Caption = "  /  /    "
            End If
            '===================================
            Toolbar1.Buttons(1).Enabled = True
            lblMensaje.Caption = "Periodo Activo para emitir solicitudes"
        Else
            lblFechaFinal.ForeColor = RGB(211, 34, 14)
            lblFechaFinal.Caption = lodynConsultaFechas("FECFIN").Value
            If lblnMostrar = False Then MsgBox "Periodo de Solicitudes culminado", vbInformation, "Aviso"
            lblnMostrar = True
            lstrPaseAnular = "no"
            lstrPaseInicio = "no"
            Toolbar1.Buttons(1).Enabled = False
        End If
    Else
        Toolbar1.Buttons(7).Enabled = False
        MsgBox "Aún no se ha habilitado periodo alguno", vbInformation, "Aviso"
    End If
    Screen.MousePointer = vbNormal
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub

Private Sub sub_sge_Llena_Grilla()
    On Error GoTo ERROR
    Dim StrSql As String
    Screen.MousePointer = vbHourglass
    '--Correcion
    'StrSql = "SELECT S.NUM_SOLICITUD," & _
             "(SELECT INITCAP(APE_PAT_USUARIO||' '||APE_MAT_USUARIO||' '||DES_NOMBRE) FROM NUEVO.MAE_USUARIO_BTL WHERE COD_USUARIO = S.USU_EMISION) USU_EMISION," & _
             "S.FCH_EMISION," & _
             "(SELECT INITCAP(APE_PAT_USUARIO||' '||APE_MAT_USUARIO||' '||DES_NOMBRE) FROM NUEVO.MAE_USUARIO_BTL WHERE COD_USUARIO = S.USU_RECEPCION) USU_RECEPCION," & _
             "S.FCH_RECEPCION," & _
             "(SELECT INITCAP(APE_PAT_USUARIO||' '||APE_MAT_USUARIO||' '||DES_NOMBRE) FROM NUEVO.MAE_USUARIO_BTL WHERE COD_USUARIO = S.USU_ANULACION) USU_ANULACION," & _
             "S.FCH_ANULACION,S.EST_SOLICITUD,S.COD_PERIODO, S.USU_EMISION COD_EMISION,PE.FCH_INICIO,PE.FCH_FIN " & _
             "FROM SSGG.CAB_SOLICITUD_SGE S,SSGG.AUX_PERIODO_PEDIDO PE " & _
             "WHERE S.COD_PERIODO=PE.COD_PERIODO AND S.COD_AREA = '" & gstrCodAreaUsuario & "' " & _
             "ORDER BY S.FCH_EMISION DESC,S.NUM_SOLICITUD DESC"
    Set lodynConsulta = objSSGG.listaGrilla(CStr(gstrCodAreaUsuario))
    Set lodynConsultaClone = lodynConsulta.Clone
    Set oradcHistSolicitudes.Recordset = lodynConsulta
    '--------------------------------------------------------
    Screen.MousePointer = vbNormal
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub
Private Sub Form_Activate()
    On Error GoTo ERROR
    sub_sge_llena_fechas
    If lstrPaseInicio = "si" Then
        Exit Sub
    End If
    If grdHistSolicitudes.Row <> -1 Then
        If lstrDeNueva = "si" Then
            lstrDeNueva = "no"
            Call sub_sge_Llena_Grilla
        End If
        grdHistSolicitudes.SetFocus
    Else
        'Toolbar1.Buttons(9).SetFocus
    End If
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub
Private Sub Form_Load()
'    On Error GoTo ERROR
    Me.left = 0
    Me.top = 0
    redimensionaForm
    Screen.MousePointer = vbHourglass
    lblnMostrar = False
    'Call sub_sge_Centra_Ventana(Me)
    'grdHistSolicitudes.Array
    Call spSetGrdDetalle(grdHistSolicitudes)
    sub_sge_Area
    lstrPaseInicio = "si"
    Screen.MousePointer = vbNormal
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub

Private Sub sub_sge_Area()
    lblCodArea.Caption = gstrCodAreaUsuario
    lblDesArea.Caption = gstrDesAreaUsuario
End Sub

Private Sub spSetGrdDetalle(ByRef rgrd As TDBGrid)
    Dim pvarAncho As Variant
    Dim pvarTitulo As Variant
    Dim pvarAlinea As Variant
    Dim pvarCampoDato As Variant

   pvarAncho = Array(1800, 1850, 1400, 1500, 1850, 1400, 1850, 1400, 1000, 0, 0, 0)
   pvarTitulo = Array("# SOLICITUD", "Usuario Emisor", "Fecha Emisión", "ESTADO", _
            "Usuario Receptor", "Fecha Recepción", "Usuario Anulador", "Fecha Anulación", _
            "PERIODO", "Cod Emisor", "Fecha Inicio", "Fecha Fin")
   pvarAlinea = Array(2, 0, 2, 2, 0, 2, 0, 2, 2, 0, 0, 0)
   pvarCampoDato = Array("NUM_SOLICITUD", "USU_EMISION", "FCH_EMISION", _
            "EST_SOLICITUD", "USU_RECEPCION", "FCH_RECEPCION", "USU_ANULACION", "FCH_ANULACION", _
            "COD_PERIODO", "COD_EMISION", "FCH_INICIO", "FCH_FIN")

   objSSGG.spGrilla_Carga rgrd, pvarTitulo, pvarAncho, pvarAlinea, pvarCampoDato
   objSSGG.spGrilla_Traslate rgrd, "EST_SOLICITUD", "EMI", "EMITIDA"
   objSSGG.spGrilla_Traslate rgrd, "EST_SOLICITUD", "ANU", "ANULADA"

   rgrd.MarqueeStyle = dbgHighlightRow
   rgrd.HeadBackColor = &H8000000F '&H80000007
   rgrd.HeadForeColor = &H80000012 '&HFFFF&
   rgrd.RowHeight = 0
   rgrd.RowHeight = 400
   rgrd.HeadLines = 2
   rgrd.Font.Size = 8
   rgrd.Styles(5).Font.Size = 8
   rgrd.FetchRowStyle = True            'para toda la fila
   rgrd.BookmarkType = 2
   
    rgrd.Columns("USU_RECEPCION").Visible = False
    rgrd.Columns("FCH_RECEPCION").Visible = False
    rgrd.Columns("NUM_SOLICITUD").Style.Font.Bold = True
    rgrd.Columns("FCH_EMISION").Style.Font.Bold = True
    rgrd.Columns("EST_SOLICITUD").Style.Font.Bold = True
    With grdHistSolicitudes
        .Columns(9).Visible = False
        .Columns(9).AllowSizing = False
        .Columns(10).Visible = False
        .Columns(10).AllowSizing = False
        .Columns(11).Visible = False
        .Columns(11).AllowSizing = False
    End With
    
End Sub
Sub configuraGrilla()

    With grdHistSolicitudes
        .Columns(9).Visible = False
        .Columns(9).AllowSizing = False
        .Columns(10).Visible = False
        .Columns(10).AllowSizing = False
        .Columns(11).Visible = False
        .Columns(11).AllowSizing = False
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
        mdiPrincipal.picComandos.Visible = True

End Sub

'PARA TODA LA FILA
'-----------------
Private Sub grdHistSolicitudes_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)
    On Error GoTo ERROR
    lodynConsultaClone.MoveTo Bookmark - 96
    If lodynConsultaClone("COD_PERIODO").Value = lstrCodPeriodo Then
        RowStyle.BackColor = RGB(254, 243, 207) 'RGB(200, 210, 200)            ' &HC0FF3D
        If Trim(lodynConsultaClone("EST_SOLICITUD").Value) = "ANU" Then
            RowStyle.ForeColor = RGB(2, 65, 117)
        Else
            RowStyle.ForeColor = &H80&
        End If
    End If
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub


Private Sub grdHistSolicitudes_DblClick()
    cmdVisualizar_Click
End Sub

Private Sub grdHistSolicitudes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdVisualizar_Click
End Sub

Private Sub grdHistSolicitudes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   ' On Error GoTo ERROR
    If lstrPaseInicio = "si" Then
        lstrPaseInicio = "no"
        Exit Sub
    End If
    If lodynConsulta.EOF Or lodynConsulta.BOF Then
        MsgBox "No se encontraron registros de Solicitudes", vbInformation, "Aviso"
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
    Else
        lstrNumSolicitud = Trim(lodynConsulta("NUM_SOLICITUD").Value)
        lstrEstado = Trim(lodynConsulta("EST_SOLICITUD").Value)
        lblFechaInicial.Caption = grdHistSolicitudes.Columns(10).Text
        lblFechaFinal.Caption = grdHistSolicitudes.Columns(11).Text
        If lstrEstado = "ANU" Then
            Toolbar1.Buttons(3).Enabled = False
        Else
            Toolbar1.Buttons(3).Enabled = True
        End If
        
        Toolbar1.Buttons(5).Enabled = True
        Toolbar1.Buttons(7).Enabled = True
        grdHistSolicitudes.SetFocus
    End If
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ERROR
Select Case Button.Index
    Case 1 ' Nueva
        lstrDeNueva = "si"
        frm_ADM_Solicitud.blnNuevaHistorial = IIf(lstrDeNueva = "si", True, False)
        frm_ADM_Solicitud.sub_sge_Generar_Solicitud
    Case 3 'Anular
        If Trim(lstrNumSolicitud) = "" Then
        MsgBox "No existe Solicitud a anular", vbInformation, "Aviso"
        Exit Sub
        End If
        If lstrEstado = "ANU" Then
            MsgBox "La Solicitud ya está anulada", vbInformation, "Aviso"
            Exit Sub
        End If
        If grdHistSolicitudes.Columns(9).Text <> gstrCodUsuario Then
            MsgBox "Esta solicitud solo puede ser anulada por el usuario que la emitió!", vbExclamation, "Aviso"
            Exit Sub
        End If
            
        If MsgBox("Desea anular la solicitud?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
            Exit Sub
        End If
        Dim lstrError As String
        Dim lstrMensaje As String
        Screen.MousePointer = vbHourglass
        Dim gvarValores  As Variant
        Dim gvarIO  As Variant
                                  
         'Add RET_ERROR as an Output parameter and set its initial value.
        godbVentas.Parameters.Add "RET_ERROR", "0", ORAPARM_OUTPUT
        godbVentas.Parameters("RET_ERROR").serverType = ORATYPE_VARCHAR2
    
         'Add RET_ACCION as an Output parameter and set its initial value.
        godbVentas.Parameters.Add "RET_MENSAJE", "0", ORAPARM_OUTPUT
        godbVentas.Parameters("RET_MENSAJE").serverType = ORATYPE_VARCHAR2
    
        'Add NUM_SOLICITUD as an Input/Output parameter and set its initial value.
        godbVentas.Parameters.Add "NUM_SOLICITUD", lstrNumSolicitud, ORAPARM_INPUT
        godbVentas.Parameters("NUM_SOLICITUD").serverType = ORATYPE_VARCHAR2
    
        'Add NUM_SOLICITUD as an Input/Output parameter and set its initial value.
        godbVentas.Parameters.Add "COD_USUARIO", Trim(gstrCodUsuario), ORAPARM_INPUT
        godbVentas.Parameters("COD_USUARIO").serverType = ORATYPE_VARCHAR2
    
        'Add NUM_SOLICITUD as an Input/Output parameter and set its initial value.
        godbVentas.Parameters.Add "COD_AREA", Trim(gstrCodAreaUsuario), ORAPARM_INPUT
        godbVentas.Parameters("COD_AREA").serverType = ORATYPE_VARCHAR2
    
        'Execute the Stored Procedure.
        godbVentas.ExecuteSQL ("Begin SSGG.SP_SGE_ANU_SOLICITUD (:RET_ERROR, :RET_MENSAJE, :NUM_SOLICITUD, :COD_USUARIO, :COD_AREA); end;")
    
        lstrError = godbVentas.Parameters("RET_ERROR").Value
        lstrMensaje = IIf(IsNull(godbVentas.Parameters("RET_MENSAJE").Value), "", godbVentas.Parameters("RET_MENSAJE").Value)
    
        godbVentas.Parameters.Remove "RET_ERROR"
        godbVentas.Parameters.Remove "RET_MENSAJE"
        godbVentas.Parameters.Remove "NUM_SOLICITUD"
        godbVentas.Parameters.Remove "COD_USUARIO"
        godbVentas.Parameters.Remove "COD_AREA"
    
        If Trim(lstrError) = "0" Then
            MsgBox "Solicitud Anulada", vbInformation, "Aviso"
        Else
            MsgBox Trim(lstrMensaje), vbCritical, "Error"
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        
        Call sub_sge_Llena_Grilla
        Screen.MousePointer = vbNormal
    Case 5 'Visualizar
        If lstrNumSolicitud <> "" Then
            frm_ADM_Solicitud.blnNuevaHistorial = True
            frm_ADM_Solicitud.sub_sge_Visualizar_Solicitud (lstrNumSolicitud)
        Else
            MsgBox "No existe Solicitud a visualizar", vbInformation, "Aviso"
        End If
    Case 7 'Recepcionar
        frm_ADM_RecepcionSolicitudes.blnNuevaForm = True
        frm_ADM_RecepcionSolicitudes.Show
    Case 9
        Screen.MousePointer = vbHourglass
        grdHistSolicitudes.Enabled = True
        lstrPaseInicio = "no"
        Call sub_sge_Llena_Grilla
        Screen.MousePointer = vbNormal
    Case 11
        mdiPrincipal.picComandos.Enabled = True
        Unload Me
        
End Select
    Exit Sub
ERROR:
        MsgBox Err.Description, vbExclamation, "Error"
End Sub
'Procedimiento para redimensionar los controles y objetos del form
Sub redimensionaForm()
    If objSSGG.verifica_resolucion < 1792 Then
        Me.Width = 7200
        grdHistSolicitudes.Width = 7150
        lblFondo2.Width = 7150
        With Toolbar1
            .Buttons(1).Caption = "&Nueva"
            .Buttons(3).Caption = "&Anular"
            .Buttons(5).Caption = "&Visualizar"
            .Buttons(6).Caption = "&Recep."
            .Buttons(9).Caption = "&Historial"
            .Buttons(11).Caption = "&Salir"
            
        End With
    Else
        Me.Width = 10450
        grdHistSolicitudes.Width = 10350
        lblFondo2.Width = 10350
        With Toolbar1
            .Buttons(1).Caption = "&Nueva"
            .Buttons(3).Caption = "&Anular"
            .Buttons(5).Caption = "&Visualizar"
            .Buttons(6).Caption = "&Recep."
            .Buttons(9).Caption = "&Historial"
            .Buttons(11).Caption = "&Salir"
            
        End With
    End If
End Sub

