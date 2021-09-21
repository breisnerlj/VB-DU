VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_ADM_Solicitud 
   BorderStyle     =   0  'None
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8295
   DrawStyle       =   5  'Transparent
   Icon            =   "frm_ADM_Solicitud.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Begin TrueDBGrid70.TDBGrid grdProductos 
      Height          =   4515
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7964
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   60
      TabIndex        =   2
      Top             =   420
      Width           =   10215
      Begin VB.Image imgLogo 
         Height          =   720
         Left            =   8520
         Picture         =   "frm_ADM_Solicitud.frx":0442
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1605
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
         Left            =   4260
         TabIndex        =   10
         Top             =   480
         Width           =   1695
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
         Left            =   1860
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
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
         Left            =   3840
         TabIndex        =   8
         Top             =   570
         Width           =   330
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
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
         Left            =   1200
         TabIndex        =   7
         Top             =   570
         Width           =   540
      End
      Begin VB.Label lblPeriodo 
         AutoSize        =   -1  'True
         Caption         =   "ULTIMO PERIODO:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1680
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
         Left            =   1860
         TabIndex        =   5
         Top             =   1020
         Width           =   5250
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
         Left            =   1080
         TabIndex        =   4
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AREA:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1110
         Width           =   570
      End
   End
   Begin MSComctlLib.ImageList IlsImagen 
      Left            =   7560
      Top             =   5040
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
            Picture         =   "frm_ADM_Solicitud.frx":2588
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_Solicitud.frx":2B22
            Key             =   "buscar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_Solicitud.frx":30BC
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_Solicitud.frx":3656
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_Solicitud.frx":3BF0
            Key             =   "Chek"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_Solicitud.frx":418A
            Key             =   "Bien"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_Solicitud.frx":4724
            Key             =   "Agregar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_Solicitud.frx":4CBE
            Key             =   "Atras"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_Solicitud.frx":5258
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_Solicitud.frx":57F2
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_Solicitud.frx":5D8C
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_Solicitud.frx":6326
            Key             =   "Hora"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   11
      Top             =   6570
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1111
      ButtonWidth     =   1508
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "IlsImagen"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Modificar"
            Object.ToolTipText     =   "Modificar Cantidad"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Agregar"
            Object.ToolTipText     =   "Agregar Producto"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Key             =   "Salir"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   0
      Picture         =   "frm_ADM_Solicitud.frx":68C0
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Servicios Generales-Solicitud de Productos"
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
      TabIndex        =   12
      Top             =   120
      Width           =   4665
   End
   Begin VB.Label lblMensaje 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000012&
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
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frm_ADM_Solicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lodynConsulta As oraDynaset
Dim lstrSolicitud As String
Dim lstrCodPeriodo As String
Dim lstrHabilitaCantidad As String
Dim lxdbProductos As New XArrayDB
Private lblnVer As Boolean
Dim objSSGG As New clsSSGG
Public blnNuevaHistorial As Boolean


Private Sub sub_sge_llena_grilla_detalle()
    On Error GoTo ERROR
    Dim StrSql As String
    Screen.MousePointer = vbHourglass
    'StrSql = " SELECT D.COD_PRODUCTO_GEN,P.DES_PRODUCTO,D.CTD_PRODUCTO " & _
             " FROM SSGG.DET_SOLICITUD_SGE D," & _
             " SSGG.MAE_PRODUCTO_GENERICO P" & _
             " WHERE D.COD_PRODUCTO_GEN=P.COD_PRODUCTO_GEN " & _
             " AND D.NUM_SOLICITUD='" & lstrSolicitud & "' " & _
             " ORDER BY P.DES_PRODUCTO"
    'Set lodynConsulta = godbVentas.CreateDynaset(StrSql, 0&)
    Set lodynConsulta = objSSGG.listaProductoSolicitud(lstrSolicitud)
    
    If Not lodynConsulta.EOF Then
        lxdbProductos.Clear   ' don't forget
        lxdbProductos.ReDim 0, lodynConsulta.RecordCount - 1, 0, 3
        Dim i As Integer
        i = 0
        lodynConsulta.MoveFirst
        While Not lodynConsulta.EOF
'            lxdbProductos(i, 0) = lodynConsulta("ITEM").Value
            lxdbProductos(i, 1) = lodynConsulta("COD_PRODUCTO_GEN").Value
            lxdbProductos(i, 2) = lodynConsulta("DES_PRODUCTO").Value
            lxdbProductos(i, 3) = lodynConsulta("CTD_PRODUCTO").Value
            lodynConsulta.MoveNext
            i = i + 1
        Wend
         sub_sge_num_item
         grdProductos.Rebind
         grdProductos.MoveFirst
    End If
    If lodynConsulta.RecordCount = 0 Then
        Screen.MousePointer = vbNormal
        MsgBox "Solicitud Sin Productos", vbInformation, "Aviso"
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
    End If
    Screen.MousePointer = vbNormal
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub

Private Sub sub_sge_num_item()
    On Error GoTo ERROR
    Dim i As Integer
    For i = lxdbProductos.LowerBound(1) To lxdbProductos.UpperBound(1)
        lxdbProductos(i, 0) = i + 1
    Next i
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub Form_Activate()
    On Error GoTo Control
    If lblnVer = False Then
        If Trim(lstrSolicitud) <> "" Then
            lblFechaInicial.Caption = "--------"
            lblFechaFinal.Caption = "--------"
        Else
            sub_sge_llena_fechas
        End If
    End If
    Exit Sub
Control:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        mdiPrincipal.picComandos.Enabled = True
        Unload Me
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiPrincipal.picComandos.Visible = True
    lblnVer = False
End Sub

Private Sub GRDPRODUCTOS_Click()
    grdProductos_GotFocus
End Sub

Private Sub GRDPRODUCTOS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then sub_EditarRegistro
End Sub



'Agrega un nuevo producto a la grilla con la cantidad solicitada
Private Sub sub_AgregarProducto()
    On Error GoTo ERROR
    Dim varNewRegistro As Variant
    Dim lvarRegistro As Variant
    If lstrHabilitaCantidad = "no" Then
        Exit Sub
    End If
    frm_ADM_BusProdSolicitud.blnNuevoForm = True
    Call frm_ADM_BusProdSolicitud.sub_sge_Agregar_Producto(varNewRegistro, lxdbProductos)
    Screen.MousePointer = vbHourglass
    If UBound(varNewRegistro) <> -1 Then
        lvarRegistro = Array("", varNewRegistro(0), _
                             varNewRegistro(1), varNewRegistro(2))
        sub_Agrega_Item lvarRegistro
        sub_sge_num_item
        grdProductos.Rebind
        sub_sge_deshabilita
        grdProductos.MoveLast
    End If
    'Toolbar1.Buttons(3)..SetFocus
    Screen.MousePointer = vbNormal
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub



Sub sub_Agrega_Item(ByVal vvarRegistro As Variant)
    On Error GoTo ERROR
    Dim lintIndex As Integer
    lxdbProductos.AppendRows
    lintIndex = lxdbProductos.UpperBound(1)
    Dim i As Integer
    i = 0
    While i <= UBound(vvarRegistro)
        lxdbProductos(lintIndex, i) = vvarRegistro(i)
        i = i + 1
    Wend
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub grdProductos_DblClick()
    sub_EditarRegistro
End Sub

Private Sub cmdModificar_Click()
    sub_EditarRegistro
End Sub

'Permite editar la cantidad registrada
Private Sub sub_EditarRegistro()
    On Error GoTo ERROR
    If lxdbProductos.Count(1) = 0 Or lstrHabilitaCantidad = "no" Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Dim ldblCantidad As Double
    ldblCantidad = 0
    If IsNumeric(grdProductos.Columns(3).Text) Then ldblCantidad = CDbl(grdProductos.Columns(3).Text)
    Call frm_ADM_CantProdSolicitud.sub_sge_Cantidad(ldblCantidad, lxdbProductos(grdProductos.Bookmark, 2))
    If ldblCantidad > 0 Then
        lxdbProductos(grdProductos.Bookmark, 3) = ldblCantidad
        grdProductos.Rebind
    End If
    grdProductos.SetFocus
    Screen.MousePointer = vbNormal
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub




Private Sub sub_sge_deshabilita()
    On Error GoTo ERROR
    If lxdbProductos.UpperBound(1) = -1 Or lxdbProductos.Count(1) = 0 Then
        Toolbar1.Buttons(5).Enabled = False
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        'Toolbar1.Buttons(3)..SetFocus
    Else
        Toolbar1.Buttons(5).Enabled = True
        Toolbar1.Buttons(1).Enabled = True
        Toolbar1.Buttons(7).Enabled = True
        'Toolbar1.Buttons(3)..SetFocus
    End If
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub cmdGrabar_Click()
    On Error GoTo ERROR
    
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub

Private Sub sub_sge_llena_fechas()
    On Error GoTo ERROR
    Dim lodynConsultaFechas As oraDynaset
    Dim strSqlFechas As String
    
'    strSqlFechas = " SELECT COD_PERIODO,FCH_INICIO,NVL(FCH_FIN,'01/01/2000') AS FECFIN" & _
                   " FROM SSGG.AUX_PERIODO_PEDIDO" & _
                   " WHERE COD_PERIODO = (SELECT MAX(COD_PERIODO) FROM SSGG.AUX_PERIODO_PEDIDO )"
'    Set lodynConsultaFechas = godbVentas.CreateDynaset(strSqlFechas, 0&)
    Set lodynConsultaFechas = objSSGG.llenaFechas

    If lodynConsultaFechas.RecordCount <> 0 Then
        lblFechaInicial.Caption = lodynConsultaFechas("FCH_INICIO").Value
        If Trim(lodynConsultaFechas("FCH_INICIO").Value) <> "" And Trim(lodynConsultaFechas("FECFIN").Value) = "01/01/2000" Then
            lblFechaFinal.Caption = "  /  /    "
            lblMensaje.Caption = "PERIODO ACTIVO PARA EMITIR SOLICITUDES"
            lstrCodPeriodo = Trim(lodynConsultaFechas("COD_PERIODO").Value)
            Toolbar1.Buttons(3).Enabled = True
            'Toolbar1.Buttons(3).SetFocus
        Else
            lblFechaFinal.Caption = lodynConsultaFechas("FECFIN").Value
            Toolbar1.Buttons(3).Enabled = False
            MsgBox "Periodo de Solicitudes culminado", vbInformation, "Aviso"
            mdiPrincipal.picComandos.Enabled = True
            Unload Me
        End If
    Else
        MsgBox "Aún no se ha habilitado periodo alguno", vbInformation, "Aviso"
        mdiPrincipal.picComandos.Enabled = True
        Unload Me
    End If
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo ERROR
    If KeyAscii = vbKeyEscape Then
        If lstrHabilitaCantidad = "si" Then
            If MsgBox("Desea cerrar la ventana?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
                mdiPrincipal.picComandos.Enabled = True
                Unload Me
            End If
        Else
            mdiPrincipal.picComandos.Enabled = True
            Unload Me
        End If
    End If
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub Form_Load()
    On Error GoTo ERROR
    Screen.MousePointer = vbHourglass
'    Call sub_sge_Centra_Ventana(Me)
    Me.Left = 0
    Me.Top = 0
    redimensionaForm
    Call spSetGrdDetalle(grdProductos)
    
    grdProductos.Array = lxdbProductos
    sub_sge_Area
    
    If Trim(lstrSolicitud) <> "" Then
        lblMensaje.Caption = "VISUALIZANDO SOLICITUD"
        grdProductos.Caption = "PRODUCTOS SOLICITADOS - SOLICITUD Nº " & lstrSolicitud
        Call sub_sge_llena_grilla_detalle    ' VISUALIZAR
    Else
        lblMensaje.Caption = ""
        grdProductos.Caption = "PRODUCTOS SOLICITADOS"
        lxdbProductos.ReDim 0, -1, 0, 3      ' CREAR (NUEVO)
    End If
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

Public Sub sub_sge_Generar_Solicitud()
    lstrSolicitud = ""
    lstrHabilitaCantidad = "si"
    Me.Show
    '---------------------
End Sub

Public Sub sub_sge_Visualizar_Solicitud(ByVal strNumSolicitud As String)
    Dim StrSql As String
    Dim odyn As oraDynaset
    lstrSolicitud = strNumSolicitud
    lblnVer = True
    lblPeriodo.Caption = "PERIODO GENERADO"
    
    'StrSql = "SELECT P.FCH_INICIO,P.FCH_FIN FROM SSGG.AUX_PERIODO_PEDIDO P,SSGG.CAB_SOLICITUD_SGE  S " & _
             "WHERE P.COD_PERIODO=S.COD_PERIODO AND S.NUM_SOLICITUD='" & strNumSolicitud & "'"
    
    'Set odyn = godbVentas.CreateDynaset(StrSql, ORADYN_READONLY)
    Set odyn = objSSGG.periodoGenerado(strNumSolicitud)
    If Not odyn.EOF Then
        lblFechaInicial.Caption = odyn("FCH_INICIO").Value
        If Not IsNull(odyn("FCH_FIN").Value) Then
            lblFechaFinal.Caption = odyn("FCH_FIN").Value
        End If
    End If
    lstrHabilitaCantidad = "no"
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(3).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    Me.Show
    '---------------------
End Sub

Private Sub spSetGrdDetalle(ByRef rgrd As TDBGrid)
    Dim pvarAncho As Variant
    Dim pvarTitulo As Variant
    Dim pvarAlinea As Variant
    Dim pvarCampoDato As Variant
 
    pvarAncho = Array(600, 1000, 7200, 1000) '6230
    pvarTitulo = Array("ITEM", "COD.PROD", "PRODUCTO", "CANT.")
    pvarAlinea = Array(2, 2, 0, 2)
    objSSGG.spGrilla_Carga rgrd, pvarTitulo, pvarAncho, pvarAlinea
    
    rgrd.MarqueeStyle = dbgHighlightRow
   rgrd.HeadBackColor = &H8000000F '&H80000007
   rgrd.HeadForeColor = &H80000012 '&HFFFF&
    rgrd.RowHeight = 0
    'rgrd.RowHeight = 1 * rgrd.RowHeight
    rgrd.RowHeight = 1 * 280
    rgrd.HeadLines = 1
    rgrd.Font.Size = 8
    rgrd.Styles(5).Font.Size = 8
    rgrd.Columns(1).Visible = False
    rgrd.AllowUpdate = True
    Dim i As Integer
    For i = 0 To rgrd.Columns.Count - 1
        rgrd.Columns(i).AllowFocus = False
    Next i
    rgrd.Columns("CANT.").Style.Font.Bold = True
End Sub


Private Sub grdProductos_GotFocus()
        grdProductos.Styles(5).BackColor = &H8000000D
        grdProductos.Styles(5).Font.Bold = True
End Sub

Private Sub grdProductos_LostFocus()
        grdProductos.Styles(5).BackColor = &H8000000C
        grdProductos.Styles(5).Font.Bold = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ERROR
Select Case Button.Index
    Case 1
        sub_EditarRegistro
    Case 3
        sub_AgregarProducto
        
        'Toolbar1.Buttons(3)..SetFocus
    Case 5
        Screen.MousePointer = vbHourglass
        grdProductos.SetFocus
        grdProductos.Delete
        Call sub_sge_num_item
        grdProductos.Rebind
        Call sub_sge_deshabilita
        Screen.MousePointer = vbNormal
    Case 7
        If lxdbProductos.UpperBound(1) <> (lxdbProductos.LowerBound(1) - 1) Then
            Dim i As Integer
            For i = lxdbProductos.LowerBound(1) To lxdbProductos.UpperBound(1)
                If Val(lxdbProductos(i, 3)) = 0 Then
                    MsgBox "Existe un Producto sin Cantidad" & Chr(13) & "Item: " & (i + 1), vbInformation, "Error"
                    grdProductos.SetFocus
                    Exit Sub
                End If
                If IsNumeric(lxdbProductos(i, 3)) = False Then
                    MsgBox "Cantidad no permitida" & Chr(13) & "Item: " & (i + 1), vbInformation, "Error"
                    grdProductos.SetFocus
                    Exit Sub
                End If
            Next i
        Else
            MsgBox "No existen Productos solicitados", vbExclamation, "Aviso"
            Exit Sub
        End If
    
        
        'CONSTRUYENDO CADENAS CON PALOTES
        Dim lstrCodigos As String
        Dim lstrCantUnidades As String
        Dim lstrCantFracciones As String, arrValores As Variant, arrDireccion As Variant
        
        Dim j As Integer
        lstrCodigos = ""
        lstrCantUnidades = ""
        lstrCantFracciones = ""
        Screen.MousePointer = vbHourglass
        
        For j = 0 To lxdbProductos.Count(1) - 1
            If Trim(lxdbProductos.Value(j, 1)) <> "" Then
                lstrCodigos = lxdbProductos.Value(j, 1) & "|" & lstrCodigos
                lstrCantUnidades = lxdbProductos.Value(j, 3) & "|" & lstrCantUnidades
                lstrCantFracciones = "0" & "|" & lstrCantFracciones
            End If
        Next j
        If Trim(lstrCodigos) = "" Then
            MsgBox "No existen productos seleccionados", vbInformation, "Aviso"
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
       
        If MsgBox("Conforme con los productos?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
           
        Dim strRetError As String, strRetMensaje As String
        Dim strMensaje As String
        'LimpiarParametros
        Screen.MousePointer = vbHourglass
         'Add RET_ERROR as an Output parameter and set its initial value.
        'godbVentas.Parameters.Add "RET_ERROR", "0", ORAPARM_OUTPUT
        'godbVentas.Parameters("RET_ERROR").serverType = ORATYPE_VARCHAR2
        
         'Add RET_ACCION as an Output parameter and set its initial value.
        'godbVentas.Parameters.Add "RET_MENSAJE", "0", ORAPARM_OUTPUT
        'godbVentas.Parameters("RET_MENSAJE").serverType = ORATYPE_VARCHAR2
        
        'Add CODIGO as an Input/Output parameter and set its initial value.
        'godbVentas.Parameters.Add "CODPERIODO", Trim(lstrCodPeriodo), ORAPARM_INPUT
        'godbVentas.Parameters("CODPERIODO").serverType = ORATYPE_VARCHAR2
        'Add CODIGO as an Input/Output parameter and set its initial value.
        'godbVentas.Parameters.Add "CODAREA", gstrCodAreaUsuario, ORAPARM_INPUT
        'godbVentas.Parameters("CODAREA").serverType = ORATYPE_VARCHAR2
        'Add DESCRIPCION as an Input/Output parameter and set its initial value.
        'godbVentas.Parameters.Add "USUARIO", gstrCodUsuario, ORAPARM_INPUT
        'godbVentas.Parameters("USUARIO").serverType = ORATYPE_VARCHAR2
        'Add DESCRIPCION as an Input/Output parameter and set its initial value.
        'godbVentas.Parameters.Add "CODPRODUCTOS", Trim(lstrCodigos), ORAPARM_INPUT
        'godbVentas.Parameters("CODPRODUCTOS").serverType = ORATYPE_VARCHAR2
        'Add DESCRIPCION as an Input/Output parameter and set its initial value.
        'godbVentas.Parameters.Add "UNIDADES", Trim(lstrCantUnidades), ORAPARM_INPUT
        'godbVentas.Parameters("UNIDADES").serverType = ORATYPE_VARCHAR2
        'Add DESCRIPCION as an Input/Output parameter and set its initial value.
        'godbVentas.Parameters.Add "FRACCIONES", Trim(lstrCantFracciones), ORAPARM_INPUT
        'godbVentas.Parameters("FRACCIONES").serverType = ORATYPE_VARCHAR2
        ''Execute the Stored Procedure.
        'godbVentas.ExecuteSQL ("Begin SSGG.SP_SGE_GEN_SOLICITUD (:RET_ERROR, :RET_MENSAJE, :CODPERIODO, :CODAREA, :USUARIO, :CODPRODUCTOS, :UNIDADES, :FRACCIONES); end;")
        'lstrError = godbVentas.Parameters("RET_ERROR").Value
        'lstrMensaje = IIf(IsNull(godbVentas.Parameters("RET_MENSAJE").Value), "", godbVentas.Parameters("RET_MENSAJE").Value)
        arrValores = Array("", "", Trim(lstrCodPeriodo), gstrCodAreaUsuario, gstrCodUsuario, Trim(lstrCodigos), _
                    Trim(lstrCantUnidades), Trim(lstrCantFracciones))
        arrDireccion = Array(Salida, Salida, entrada, entrada, entrada, entrada, _
                        entrada, entrada)
        strMensaje = gclsOracle.SP("SSGG.SP_SGE_GEN_SOLICITUD", arrValores, arrDireccion)
        If strMensaje = "" Then
            MsgBox "Solicitud Generada", vbInformation, "Aviso"
            Toolbar1.Buttons(1).Enabled = False
            Toolbar1.Buttons(5).Enabled = False
            Toolbar1.Buttons(3).Enabled = False
            Toolbar1.Buttons(7).Enabled = False
        Else
            Screen.MousePointer = vbNormal
            MsgBox strMensaje, vbCritical, App.ProductName
            Exit Sub
        End If
       
        'lstrHabilitaCantidad = "no"
        Screen.MousePointer = vbNormal
    Case 9
        mdiPrincipal.picComandos.Enabled = Not blnNuevaHistorial
        Unload Me
End Select
Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub

'Procedimiento para redimensionar los controles y objetos del form
Sub redimensionaForm()
    If objSSGG.verifica_resolucion < 1792 Then
        Me.Width = 7200
        grdProductos.Width = 7150
        Frame1.Width = 7150
        lblDesArea.Width = 5250
    Else
        Me.Width = 10400
        grdProductos.Width = 10350
        Frame1.Width = 10350
        lblDesArea.Width = 7750
    End If
End Sub

