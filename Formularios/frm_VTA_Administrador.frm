VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_VTA_Administrador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo Administrador "
   ClientHeight    =   6885
   ClientLeft      =   6195
   ClientTop       =   330
   ClientWidth     =   4845
   Icon            =   "frm_VTA_Administrador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin TrueDBGrid70.TDBGrid grdAdministrador 
      Bindings        =   "frm_VTA_Administrador.frx":030A
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   12091
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Codigo"
      Columns(0).DataField=   "COD_MENU"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripción"
      Columns(1).DataField=   "DES_MENU"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   16
      Columns(2)._MaxComboItems=   5
      Columns(2).ValueItems(0)._DefaultItem=   0
      Columns(2).ValueItems(0).Value=   "1"
      Columns(2).ValueItems(0).Value.vt=   8
      Columns(2).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(2).ValueItems(0).DisplayValue(0)=   "bHQAADYDAABCTTYDAAAAAAAANgAAACgAAAAQAAAAEAAAAAEAGAAAAAAAAAMAAAAAAAAAAAAAAAAA"
      Columns(2).ValueItems(0).DisplayValue(1)=   "AAAAAAD////////////////////////f39//////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(2)=   "//////////////+/v7+fn5/f39//////////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(3)=   "//+fn58/P/+fn5////////////////////////////////////////////////+/v79fX98AAP9/"
      Columns(2).ValueItems(0).DisplayValue(4)=   "f7////////////////////////////////////////////////+AgIAAAP8AAP9fX9+/v7//////"
      Columns(2).ValueItems(0).DisplayValue(5)=   "//////////////////////////////////////+fn98AAP9/f/8AAP+fn5//////////////////"
      Columns(2).ValueItems(0).DisplayValue(6)=   "//////////////////////////8/P/+/v/+/v/8AAP9fX9+fn5//////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(7)=   "//////////////+/v/////////8AAP8AAP9/f7/f39//////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(8)=   "//////////////9/f/8AAP8/P/+fn5/f39//////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(9)=   "//9/f/8AAP8AAP8/P/+/v7+fn5////////////////////////////////////////////8AAP8A"
      Columns(2).ValueItems(0).DisplayValue(10)=   "AP8AAP8AAP+fn9+/v7////////////////////////////////////////+/v/8AAP8AAP8AAP8A"
      Columns(2).ValueItems(0).DisplayValue(11)=   "AP+/v/////////////////////////////////////////////8/P/8AAP8/P///////////////"
      Columns(2).ValueItems(0).DisplayValue(12)=   "//////////////////////////////////////+/v/8AAP//////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(13)=   "//////////////////////////////+/v///////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(14)=   "//////////////////////////////////////////8="
      Columns(2).ValueItems(0).DisplayValue.vt=   9
      Columns(2).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems.Count=   1
      Columns(2).Caption=   "FLG_AUTORIZABLE"
      Columns(2).DataField=   "FLG_AUTORIZABLE"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=529"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=450"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4763"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4683"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2566"
      Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   0
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
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&HFFFFFF&,.fgcolor=&HFFFFE6&"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.bgcolor=&HFFFFFF&,.fgcolor=&H0&"
      _StyleDefs(41)  =   ":id=32,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(42)  =   ":id=32,.fontname=MS Sans Serif"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(50)  =   "Named:id=33:Normal"
      _StyleDefs(51)  =   ":id=33,.parent=0,.appearance=1"
      _StyleDefs(52)  =   "Named:id=34:Heading"
      _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(54)  =   ":id=34,.wraptext=-1"
      _StyleDefs(55)  =   "Named:id=35:Footing"
      _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(57)  =   "Named:id=36:Selected"
      _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=37:Caption"
      _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(61)  =   "Named:id=38:HighlightRow"
      _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&HFEEBDE&,.fgcolor=&H80000012&,.bold=-1,.fontsize=825"
      _StyleDefs(63)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(64)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(65)  =   "Named:id=39:EvenRow"
      _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(67)  =   "Named:id=40:OddRow"
      _StyleDefs(68)  =   ":id=40,.parent=33"
      _StyleDefs(69)  =   "Named:id=41:RecordSelector"
      _StyleDefs(70)  =   ":id=41,.parent=34"
      _StyleDefs(71)  =   "Named:id=42:FilterBar"
      _StyleDefs(72)  =   ":id=42,.parent=33"
   End
   Begin MSComctlLib.ImageList ilsImagenes 
      Left            =   2640
      Top             =   4680
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
            Picture         =   "frm_VTA_Administrador.frx":0321
            Key             =   "abajo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":0875
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":193F
            Key             =   "arriba"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":1E93
            Key             =   "sum"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":1FB3
            Key             =   "manito"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":22CF
            Key             =   "etiqueta"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":2561
            Key             =   "mensaje"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":2EB3
            Key             =   "nota"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":3805
            Key             =   "libro"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":4157
            Key             =   "notamensaje"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":53A9
            Key             =   "btl"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":5CFB
            Key             =   "lab"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsIconosChico 
      Left            =   1800
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":664D
            Key             =   "btl_rojo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":6971
            Key             =   "btl"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsCrono 
      Left            =   1080
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":7B5D
            Key             =   "btl"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":8D4D
            Key             =   "grupo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":91A5
            Key             =   "rayo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":95FD
            Key             =   "reloj"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Administrador.frx":9A55
            Key             =   "next"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1200
      Width           =   615
   End
   Begin ORADCLibCtl.ORADC oradcAdm 
      Height          =   255
      Left            =   1560
      Top             =   3960
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
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
End
Attribute VB_Name = "frm_VTA_Administrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Padre As String
Dim odynR1 As oraDynaset
Dim strAdm As String
Private lblnVentana As Boolean

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim objPermisos As New clsAutorizacion
   On Error GoTo Control

    lblnVentana = False
    'Set odynR1 = objUsuario.ListaPermisosAdm(objUsuario.Codigo, objUsuario.Aplicacion)
    Set odynR1 = objPermisos.ListaPermisos(objUsuario.Aplicacion, objUsuario.Codigo, Padre)
    Set oradcAdm.Recordset = odynR1
    SeteaGrilla

   Exit Sub

Control:

      MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Control
    psub_KeyDownAplicacion KeyCode, Shift
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
End Sub

Private Sub SeteaGrilla()
    
    Dim i%
    For i = 0 To grdAdministrador.Columns.Count - 1
        grdAdministrador.Columns(i).Visible = False
        grdAdministrador.Columns(i).WrapText = True
    Next i
    
    grdAdministrador.RowHeight = 2.2 * grdAdministrador.RowHeight
    
    grdAdministrador.Columns(2).FetchStyle = True
    
    grdAdministrador.Columns(0).Visible = False
    grdAdministrador.Columns(1).Visible = True
    grdAdministrador.Columns(2).Visible = True
    grdAdministrador.Style.VerticalAlignment = dbgVertCenter
    grdAdministrador.Columns(2).Alignment = dbgCenter
    grdAdministrador.Columns(2).ButtonText = True
    grdAdministrador.HeadLines = 0
    grdAdministrador.MarqueeStyle = dbgHighlightRow
    grdAdministrador.AllowUpdate = False

    psub_Grilla_Traslate grdAdministrador, "FLG_AUTORIZABLE", "1", ilsImagenes.ListImages("arriba").Picture
End Sub

Private Sub grdAdministrador_DblClick()

   On Error GoTo Control

    strAdm = Trim(grdAdministrador.Columns(0).Value)
    Select Case grdAdministrador.Columns(0).Value
        Case "011"
            Unload Me
            objVenta.CodigoDocumentoVenta = objUsuario.TipoDocNC
            objVenta.CodModalidadVenta = "001"
            objVenta.ptmModalidad = Venta_Regular
            ptmTipoPrecio = Regular
            objVenta.CodigoTipoVenta = Venta_Regular
            frm_VTA_NotaCredito.Show
            frm_VTA_Concepto_NotaCredito.Show vbModal
            
        Case "012"
            Unload Me
            frm_Lista_Caja_PreCerradas.Show
        
        Case "013"
            Unload Me
            frm_VTA_Ctrl_Depositos.Show
            
        Case "014"
            Unload Me
            frm_VTA_FormsSoat.Show
            
        Case "015"
            Unload Me
            frm_VTA_RepKardex.Show
            
        Case "016"
            Unload Me
            frm_cat_Busquedacodigo.Show
           
        Case "017"
            Unload Me
            frm_VTA_ConsultaDoc.Show
            
        Case "018"
            Unload Me
            
        Case "019"
            Unload Me
            frm_VTA_ConsCobServ.Show
            
        Case "020"
            Unload Me
            frm_VTA_Correccion.Show
            
        Case "021"
            Unload Me
            'frm_ADM_GuiasTransito.Show vbModal
            frm_ADM_AdmGuias.CodLocal = objUsuario.CodigoLocal
            frm_ADM_AdmGuias.Show 'vbModal
        
        '-------------------------------------------------
        'Agregado por DJARA para la transferencias a canje
        'PROY. LOGISTICA III - 15/10/08
        Case "022"
            Unload Me
            frm_ADM_ODevolucion.Show vbModal
        '-------------------------------------------------
        
        'Agregado por DJARA
        'REQUERIMIENTO INVENTARIO - 26/11/08
        Case "023"
            Unload Me
            frm_ADM_SustentoDiff.Show 'vbModal
        '-------------------------------------------------
        
        Case "047"
            Unload Me
            frm_ADM_Maquinas.Show
       Case "044"
            Unload Me
            frm_DLV_Pedido.Show
            
        Case "053"
            Unload Me
            mdiPrincipal.picComandos.Enabled = False
            frm_ADM_HistSolicitud.Show

        Case "054"
            Unload Me
            mdiPrincipal.picComandos.Enabled = False
            frm_ADM_Solicitud.Show

        Case "055"
            Unload Me
            mdiPrincipal.picComandos.Enabled = False
            frm_ADM_RecepcionSolicitudes.Show

        Case "056"
            Unload Me
            mdiPrincipal.picComandos.Enabled = False
            frm_ADM_HistGuiasLocal.Show

        Case "077"
            Unload Me
            frm_INV_CompraDirecta.Show

'COMENTADO POR PHERRERA 27/05/08 PARA ELIMINAR EL FORMULARIO QUE NO SE USA
'        Case "078"
'            Unload Me
'            frm_ADM_Doc_Anu_x_Dlv.Show
            
        Case "080"
            Unload Me
            frm_ADM_Producto_Petitorio.Show
            
        Case "081"
            Unload Me
            frm_VTA_ConsPetitorio.Show
        
        Case "082"
            Unload Me
            frm_ADM_SMM.CodLocal = objUsuario.CodigoLocal
            frm_ADM_SMM.Show vbModal

            'frm_VTA_HelpDesk.Show
                
        Case "083"
            Unload Me
            frm_ADM_ProductosSeleccionados.Show vbModal

        Case "091"
            Unload Me
            frm_ADM_Sincronizacion.Show
        
        Case "093"
            Unload Me
            frm_ADM_cnt_competencia.Show
        
        Case "094"
          'frm_VTA_Solicitud_Ajuste.Show vbModal
            Unload Me
            frm_ADM_SolicitudAjuste.Show
        
        Case "095"
            Unload Me
            frm_ADM_RptVentasProducto.Show
        
        '-------------------------------------------------
        'Agregado por CCIEZA para la generacion de pedidos mayoristas
        'PROY. BTL0 - 10/05/10
        Case "096"
            Unload Me
            frm_ADM_Pedido_Mayorista.Show
        '-------------------------------------------------
        Case "097"
            Unload Me
             frm_VTA_BeneficiarioNew.Show
        
        Case "098"
             Unload Me
             frm_ADM_MantMaquinasxLocal.Show
        
        Case "099"
             Unload Me
             Frm_VTA_Lista_Precios_Nuevo.Show
        Case "100"
             Unload Me
             frm_ADM_Entrega.Show
             
        Case "101"
             Unload Me
'            frm_ADM_SPVM.CodLocal = objUsuario.CodigoLocal
'            frm_ADM_SPVM.Show vbModal
             frm_ADM_PVM_Ingreso.Show vbModal
        
        Case "103"
             Unload Me
             frm_ADM_PedEspecial.CodLocal = objUsuario.CodigoLocal
             frm_ADM_PedEspecial.Show vbModal
             
        Case "105"
             Unload Me
             frm_ADM_CierreDiario.Show vbModal
             
        Case "106"
             Unload Me
             frm_VTA_PuntosConsulta.Show vbModal
        
        Case strAdm
            Unload Me

    End Select

   Exit Sub

Control:

      MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub grdAdministrador_KeyDown(KeyCode As Integer, Shift As Integer)
    If grdAdministrador.ApproxCount > 0 Then
        Select Case KeyCode
            Case vbKeyReturn
                grdAdministrador_DblClick
            ''Adicionado por JRAZURI para facilitar la navegacion con el teclado
            ''21/05/2008
            Case vbKeyUp And IsNull(grdAdministrador.GetBookmark(-1))
                grdAdministrador.MoveLast
                KeyCode = 0
            Case vbKeyDown And IsNull(grdAdministrador.GetBookmark(1))
                grdAdministrador.MoveFirst
                KeyCode = 0
            ''-----------------------------------------------------------------''
                
        End Select
    End If
End Sub
