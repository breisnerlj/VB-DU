VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_VTA_Mantenimientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimientos"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frm_VTA_Mantenimientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ilsImagenes 
      Left            =   3360
      Top             =   1920
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
            Picture         =   "frm_VTA_Mantenimientos.frx":000C
            Key             =   "abajo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Mantenimientos.frx":0560
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Mantenimientos.frx":162A
            Key             =   "arriba"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Mantenimientos.frx":1B7E
            Key             =   "sum"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Mantenimientos.frx":1C9E
            Key             =   "manito"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Mantenimientos.frx":1FBA
            Key             =   "etiqueta"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Mantenimientos.frx":224C
            Key             =   "mensaje"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Mantenimientos.frx":2B9E
            Key             =   "nota"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Mantenimientos.frx":34F0
            Key             =   "libro"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Mantenimientos.frx":3E42
            Key             =   "notamensaje"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Mantenimientos.frx":5094
            Key             =   "btl"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Mantenimientos.frx":59E6
            Key             =   "lab"
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid70.TDBGrid grdMantenimiento 
      Bindings        =   "frm_VTA_Mantenimientos.frx":6338
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   10186
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
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
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
      Appearance      =   0
      BorderStyle     =   0
      DefColWidth     =   0
      HeadLines       =   1
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
   Begin ORADCLibCtl.ORADC oradcMant 
      Height          =   255
      Left            =   1920
      Top             =   5280
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
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
Attribute VB_Name = "frm_VTA_Mantenimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Padre As String
Dim odynR1 As oraDynaset
Dim strMant As String
Private lblnVentana As Boolean

Private Sub Form_Load()
    lblnVentana = False
    Dim objPermisos As New clsAutorizacion
        'set odynR1 = objUsuario.ListaMantenimientos(objUsuario.Codigo, objUsuario.Aplicacion)
    Set odynR1 = objPermisos.ListaPermisos(objUsuario.Aplicacion, objUsuario.Codigo, Padre)
    Set oradcMant.Recordset = odynR1
    SeteaGrilla
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



Private Sub grdMantenimiento_DblClick()
    strMant = Trim(grdMantenimiento.Columns(0).Value)
    Select Case grdMantenimiento.Columns(0).Value
        Case "040"
            Unload Me
'COMENTADO POR PHERRERA 270508 PARA ELIMINAR FORMULARIO
'        Case "041"
'            Unload Me
'            frm_DLV_LstZona.Show
'            frm_DLV_LstZona.SetFocus
'COMENTADO POR PHERRERA 270508 PARA ELIMINAR FORMULARIO
'        Case "042"
'            Unload Me
'            frm_DLV_LstMotorizado.Show
'            frm_DLV_LstMotorizado.SetFocus
            
'COMENTADO POR PHERRERA 270508 PARA ELIMINAR FORMULARIO
'        Case "043"
'            Unload Me
'            frm_DLV_LstAsistMotorizado.Show
'            frm_DLV_LstAsistMotorizado.SetFocus
'
 
            
        Case "045"
            Unload Me
            frm_DLV_LstTipoDireccion.Show
            frm_DLV_LstTipoDireccion.SetFocus
            
        Case "046"
            Unload Me
            frm_DLV_Verificacion.Show
            frm_DLV_Verificacion.SetFocus
            
        Case "047"
            Unload Me
            frm_ADM_Maquinas.Show
            
        Case "048"
            Unload Me
            frm_ADM_LstMoneda.Show
            
        Case "049"
            Unload Me
            'frm_ADM_LstBanco.Show
        
        Case "050"
            Unload Me
            'frm_ADM_LstFormaPago.Show
        
        Case "051"
            Unload Me
            frm_DLV_Pedido_x_Estado.Show 'vbModal
            
        Case "052"
            Unload Me
            'frm_ADM_LstPromocion.Show
            
    End Select
End Sub

Private Sub grdMantenimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
                 grdMantenimiento_DblClick
        Case vbKeySpace
                 grdMantenimiento_DblClick
    End Select
End Sub

Private Sub SeteaGrilla()
    
    Dim i%
    For i = 0 To grdMantenimiento.Columns.Count - 1
        grdMantenimiento.Columns(i).Visible = False
        grdMantenimiento.Columns(i).WrapText = True
    Next i
    
    grdMantenimiento.RowHeight = 2.2 * grdMantenimiento.RowHeight
    
    grdMantenimiento.Columns(2).FetchStyle = True
    
    grdMantenimiento.Columns(0).Visible = False
    grdMantenimiento.Columns(1).Visible = True
    grdMantenimiento.Columns(2).Visible = True
    grdMantenimiento.Style.VerticalAlignment = dbgVertCenter
    grdMantenimiento.Columns(2).Alignment = dbgCenter
    grdMantenimiento.Columns(2).ButtonText = True
    grdMantenimiento.HeadLines = 0
    grdMantenimiento.MarqueeStyle = dbgHighlightRow
    grdMantenimiento.AllowUpdate = False
    
    psub_Grilla_Traslate grdMantenimiento, "FLG_AUTORIZABLE", "1", ilsImagenes.ListImages("btl").Picture
End Sub
