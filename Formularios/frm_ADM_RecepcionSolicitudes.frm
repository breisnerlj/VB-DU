VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_ADM_RecepcionSolicitudes 
   BorderStyle     =   0  'None
   Caption         =   "Recepción de Guias"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   Begin vbp_Ventas.ctlTextBox txtNumGuia 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
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
   Begin ORADCLibCtl.ORADC oradcGuia 
      Height          =   255
      Left            =   120
      Top             =   6480
      Visible         =   0   'False
      Width           =   3375
      _Version        =   65536
      _ExtentX        =   5953
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
   Begin TrueDBGrid70.TDBGrid grdGuia 
      Bindings        =   "frm_ADM_RecepcionSolicitudes.frx":0000
      Height          =   5205
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   9181
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
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin MSComctlLib.ImageList IlsImagen 
      Left            =   7320
      Top             =   0
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
            Picture         =   "frm_ADM_RecepcionSolicitudes.frx":0018
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_RecepcionSolicitudes.frx":05B2
            Key             =   "buscar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_RecepcionSolicitudes.frx":0B4C
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_RecepcionSolicitudes.frx":10E6
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_RecepcionSolicitudes.frx":1680
            Key             =   "Chek"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_RecepcionSolicitudes.frx":1C1A
            Key             =   "Bien"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_RecepcionSolicitudes.frx":21B4
            Key             =   "Agregar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_RecepcionSolicitudes.frx":274E
            Key             =   "Atras"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_RecepcionSolicitudes.frx":2CE8
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_RecepcionSolicitudes.frx":3282
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_RecepcionSolicitudes.frx":381C
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_ADM_RecepcionSolicitudes.frx":3DB6
            Key             =   "Hora"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   6
      Top             =   6570
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   1111
      ButtonWidth     =   1931
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "IlsImagen"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Mostrar"
            Object.ToolTipText     =   "Mostrar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Recepcionar"
            Object.ToolTipText     =   "Recepcionar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Key             =   "Salir"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   0
      Picture         =   "frm_ADM_RecepcionSolicitudes.frx":4350
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Servicios Generales-Recepcion de Solicitudes-Guias"
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
      TabIndex        =   7
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   1020
      Width           =   6975
   End
   Begin VB.Label lblDesArea 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000008&
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
      Left            =   6600
      TabIndex        =   4
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lblNumGuia 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "GUÍA RECIBIDA Nº"
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
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   630
      Width           =   1695
   End
   Begin VB.Image imgLogo 
      Height          =   750
      Left            =   8460
      Stretch         =   -1  'True
      Top             =   510
      Width           =   1695
   End
   Begin VB.Label lblFondo2 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Height          =   825
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   7200
   End
End
Attribute VB_Name = "frm_ADM_RecepcionSolicitudes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim lodynConsulta As oraDynaset

Dim lstrNumGuia As String
Dim objSSGG As New clsSSGG
Public blnNuevaForm As Boolean
Private Sub sub_sge_Llena_Grilla()
    On Error GoTo ERROR
    Dim StrSql As String
    Dim lodynConsultaCabecera As oraDynaset
    Dim lodynConsultaDetalle As oraDynaset
    
    'CABECERA
    '--------
    'StrSql = " SELECT G.NUM_GUIA,TO_CHAR(G.FCH_EMISION,'DD/MM/YYYY') FCH_EMISION,G.EST_GUIA," & _
             " G.COD_DESTINO," & _
             " (SELECT DES_CNT_COSTO FROM SSGG.MAE_CENTRO_COSTO WHERE COD_CNT_COSTO=G.COD_DESTINO) DESTINO " & _
             " FROM SSGG.CAB_GUIA_SGE G" & _
             " WHERE G.NUM_GUIA='" & Trim(txtNumGuia.Text) & "'" '& _
             '" AND G.COD_DESTINO = '" & CStr(Trim(gstrCodAreaUsuario)) & "'"
    Set lodynConsultaCabecera = objSSGG.cabeceraRecepcion(Trim(txtNumGuia.Text))
    'Set lodynConsultaCabecera = godbVentas.CreateDynaset(StrSql, 0&)
    
    If lodynConsultaCabecera.RecordCount = 0 Then
        MsgBox "No existe Guía Nº " & Trim(txtNumGuia.Text), vbCritical, "Error"
        Exit Sub
    End If
    
    If Trim(lodynConsultaCabecera("COD_DESTINO").Value) <> Trim(gstrCodAreaUsuario) Then
        MsgBox "GUIA NO ASIGNADA AL AREA O LOCAL", vbInformation, "Aviso"
        Exit Sub
    End If
    
    
    If Trim(lodynConsultaCabecera("EST_GUIA").Value) = "REC" Then
        MsgBox "La Guía ya ha sido RECEPCIONADA", vbExclamation, "Aviso"
        Exit Sub
    End If
    If Trim(lodynConsultaCabecera("EST_GUIA").Value) = "ANU" Then
        MsgBox "La Guía se encuentra ANULADA", vbExclamation, "Aviso"
        Exit Sub
    End If
    If Trim(lodynConsultaCabecera("EST_GUIA").Value) <> "EMI" Then
        MsgBox "La Guía NO se encuentra EMITIDA", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'DETALLE
    '--------
    'StrSql = " SELECT DG.COD_PRODUCTO,P.DES_PRODUCTO,P.CTD_FRACCIONAMIENTO, " & _
             " U.SIG_UNID_CONSUMO,DG.CTD_PRODUCTO,DG.CTD_PRODUCTO_FRAC " & _
             " FROM SSGG.DET_GUIA_SGE DG,SSGG.MAE_PRODUCTO_SGE " & _
             " P,SSGG.MAE_UNIDAD_CONSUMO  U " & _
             " WHERE DG.COD_PRODUCTO=P.COD_PRODUCTO " & _
             " AND P.COD_UNID_CONSUMO=U.COD_UNID_CONSUMO " & _
             " AND DG.NUM_GUIA='" & Trim(txtNumGuia.Text) & "' " & _
             " ORDER BY P.DES_PRODUCTO"
    'Set lodynConsultaDetalle = godbVentas.CreateDynaset(StrSql, 0&)
    '''Set lodynConsultaDetalle = objSSGG.cabeceraRecepcion(Trim(txtNumGuia.Text))
    Set lodynConsultaDetalle = objSSGG.detalleRecepcion(Trim(txtNumGuia.Text))
    
    If lodynConsultaDetalle.RecordCount = 0 Then
        MsgBox "No existe Detalle en la Guía Nº " & Trim(txtNumGuia.Text), vbCritical, "Error"
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    'TODO OK
    lstrNumGuia = Trim(lodynConsultaCabecera("NUM_GUIA").Value)
    Set oradcGuia.Recordset = lodynConsultaDetalle
    MsgBox "GUIA EMITIDA   para  " & UCase(Trim(lodynConsultaCabecera("DESTINO").Value)), vbInformation, "Aviso"
    
    
    grdGuia.Caption = "GUIA Nº:  " & lstrNumGuia
    
    lblDescripcion.Caption = "Area o Local:   " & UCase(Trim(lodynConsultaCabecera("DESTINO").Value)) & "       Fec.Emisión:   " & UCase(Trim(lodynConsultaCabecera("FCH_EMISION").Value)) & "      GUIA Nº: " & lstrNumGuia

    Toolbar1.Buttons(3).Enabled = True
    grdGuia.SetFocus
    Screen.MousePointer = vbNormal
    Exit Sub
ERROR:
    MsgBox Err.Description & Chr(13) & "Comuníquese con el area de Sistemas", vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
        If Trim(txtNumGuia.Text) = "" Then
            MsgBox "Debe ingresar un número de Guía", vbInformation, "Aviso"
            txtNumGuia.SetFocus
        Else
            ''Call sub_sge_Llena_Grilla
        End If
    Case vbKeyEscape
        mdiPrincipal.picComandos.Enabled = True
        Unload Me
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ERROR
    Me.left = 0
    Me.top = 0
    redimensionaForm
    Screen.MousePointer = vbHourglass
'    Call sub_sge_Centra_Ventana(Me)
    lblDesArea.Caption = "Area/Local:  " & gstrDesAreaUsuario
    Call spSetGrdDetalle(grdGuia)
    Screen.MousePointer = vbNormal
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub

Private Sub spSetGrdDetalle(ByRef rgrd As TDBGrid)
    Dim pvarAncho As Variant
    Dim pvarTitulo As Variant
    Dim pvarAlinea As Variant
    Dim pvarCampoDato As Variant
 
   pvarAncho = Array(900, 4450, 700, 1100, 1100, 1100)
   pvarTitulo = Array("Cod. Producto", "Producto", "Cant. Fracc.", "Unidad de Consumo", "Unidades enviadas", "Fracciones enviadas")
   pvarAlinea = Array(2, 0, 1, 2, 2, 2)
   pvarCampoDato = Array("COD_PRODUCTO", "DES_PRODUCTO", "CTD_FRACCIONAMIENTO", "SIG_UNID_CONSUMO", "CTD_PRODUCTO", "CTD_PRODUCTO_FRAC")
    
   objSSGG.spGrilla_Carga rgrd, pvarTitulo, pvarAncho, pvarAlinea, pvarCampoDato
        
   rgrd.MarqueeStyle = dbgHighlightRow
   rgrd.HeadBackColor = &H8000000F '&H80000007
   rgrd.HeadForeColor = &H80000012 '&HFFFF&
   rgrd.RowHeight = 0
   rgrd.RowHeight = 420
   rgrd.HeadLines = 2
   rgrd.Font.Size = 8
   rgrd.Styles(5).Font.Size = 8
   'rgrd.Columns(0).Visible = False
   rgrd.Columns("CTD_PRODUCTO").Style.Font.Bold = True
   rgrd.Columns("CTD_PRODUCTO_FRAC").Style.Font.Bold = True
   rgrd.Columns("CTD_PRODUCTO").Style.BackColor = RGB(254, 244, 207)
   rgrd.Columns("CTD_PRODUCTO_FRAC").Style.BackColor = RGB(254, 244, 207)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiPrincipal.picComandos.Visible = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ERROR
Select Case Button.Index
    Case 1
        If Trim(txtNumGuia.Text) = "" Then
            MsgBox "Debe ingresar un número de Guía", vbInformation, "Aviso"
            txtNumGuia.SetFocus
            resaltaObjeto txtNumGuia, Me
        Else
            Call sub_sge_Llena_Grilla
        End If
    Case 3
        If Trim(lstrNumGuia) = "" Then
            MsgBox "No existe Guía seleccionada", vbInformation, "Aviso"
            Exit Sub
        End If
    
        If MsgBox("Recepcionar Guía Nº " & Trim(lstrNumGuia) & " ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
            Exit Sub
        End If
    
        Dim lstrError As String
        Dim lstrMensaje As String
        Screen.MousePointer = vbHourglass
         'Add RET_ERROR as an Output parameter and set its initial value.
        godbVentas.Parameters.Add "RET_ERROR", "0", ORAPARM_OUTPUT
        godbVentas.Parameters("RET_ERROR").serverType = ORATYPE_VARCHAR2
    
         'Add RET_ACCION as an Output parameter and set its initial value.
        godbVentas.Parameters.Add "RET_MENSAJE", Space(100), ORAPARM_OUTPUT
        godbVentas.Parameters("RET_MENSAJE").serverType = ORATYPE_VARCHAR2
    
        'Add NUM_SOLICITUD as an Input/Output parameter and set its initial value.
        godbVentas.Parameters.Add "NUM_GUIA", Trim(lstrNumGuia), ORAPARM_INPUT
        godbVentas.Parameters("NUM_GUIA").serverType = ORATYPE_VARCHAR2
    
        'Add NUM_SOLICITUD as an Input/Output parameter and set its initial value.
        godbVentas.Parameters.Add "COD_USUARIO", Trim(gstrCodUsuario), ORAPARM_INPUT
        godbVentas.Parameters("COD_USUARIO").serverType = ORATYPE_VARCHAR2
    
        'Execute the Stored Procedure.
        godbVentas.ExecuteSQL ("Begin SSGG.SP_SGE_RECEP_GUIA (:RET_ERROR, :RET_MENSAJE, :NUM_GUIA, :COD_USUARIO); end;")
    
        lstrError = godbVentas.Parameters("RET_ERROR").Value
        lstrMensaje = IIf(IsNull(godbVentas.Parameters("RET_MENSAJE").Value), "", godbVentas.Parameters("RET_MENSAJE").Value)
    
        godbVentas.Parameters.Remove "RET_ERROR"
        godbVentas.Parameters.Remove "RET_MENSAJE"
        godbVentas.Parameters.Remove "NUM_GUIA"
        godbVentas.Parameters.Remove "COD_USUARIO"
    
        If Trim(lstrError) = "0" Then
            MsgBox "Guía Recepcionada", vbInformation, "Aviso"
        Else
            MsgBox "ERROR: " & Trim(lstrMensaje), vbCritical, "Error"
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        Toolbar1.Buttons(3).Enabled = False
        txtNumGuia.SetFocus
        Screen.MousePointer = vbNormal
    Case 5
        mdiPrincipal.picComandos.Enabled = Not blnNuevaForm
        Unload Me
End Select
    Exit Sub
ERROR:
    MsgBox Err.Description, vbExclamation, "Error"
    Screen.MousePointer = vbNormal
End Sub

Private Sub txtNumGuia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(txtNumGuia.Text) = "" Then
            MsgBox "Debe ingresar un número de Guía", vbInformation, "Aviso"
            txtNumGuia.SetFocus
        Else
            Call sub_sge_Llena_Grilla
        End If
        txtNumGuia.SetFocus
    End If
End Sub
'Procedimiento para redimensionar los controles y objetos del form
Sub redimensionaForm()
    If objSSGG.verifica_resolucion < 1792 Then
        Me.Width = 7200
        grdGuia.Width = 7150
        lblFondo2.Width = 7150
    Else
        Me.Width = 10450
        grdGuia.Width = 10350
        lblFondo2.Width = 10350
    End If
End Sub
