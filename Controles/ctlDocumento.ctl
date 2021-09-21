VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.UserControl ctlDocumento 
   BackColor       =   &H80000014&
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   ScaleHeight     =   5295
   ScaleWidth      =   7320
   Begin ORADCLibCtl.ORADC ORADC1 
      Height          =   375
      Left            =   2160
      Top             =   2160
      Visible         =   0   'False
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   207
      BackColor       =   -2147483629
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
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
   Begin TrueDBGrid70.TDBGrid grdDatos 
      Bindings        =   "ctlDocumento.ctx":0000
      Height          =   2895
      Left            =   0
      TabIndex        =   16
      Top             =   1080
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5106
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
   Begin VB.Label lblDeCajero 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cajero :"
      Height          =   195
      Left            =   0
      TabIndex        =   15
      Top             =   4500
      Width           =   1035
   End
   Begin VB.Label lblDeDependiente 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dependiente :"
      Height          =   195
      Left            =   0
      TabIndex        =   14
      Top             =   4680
      Width           =   1035
   End
   Begin VB.Label lblCajero 
      BackColor       =   &H00FFFFFF&
      Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
      Height          =   195
      Left            =   1080
      TabIndex        =   13
      Top             =   4500
      Width           =   4635
   End
   Begin VB.Label lblDependiente 
      BackColor       =   &H00FFFFFF&
      Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
      Height          =   195
      Left            =   1080
      TabIndex        =   12
      Top             =   4680
      Width           =   4635
   End
   Begin VB.Label lblDeDonacion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Donación :"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   0
      TabIndex        =   11
      Top             =   4980
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblInstitucion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   900
      TabIndex        =   10
      Top             =   4980
      Visible         =   0   'False
      Width           =   3915
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label lblAnulada 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ANULADA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblAdministrador 
      BackColor       =   &H00FFFFFF&
      Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1260
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label lblNumero 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nº XXX-XXXXXXX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3900
      TabIndex        =   6
      Top             =   315
      Width           =   3195
   End
   Begin VB.Label lblDeBoleta 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BOLETA DE VENTA"
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
      Height          =   225
      Left            =   5280
      TabIndex        =   5
      Top             =   75
      Width           =   1755
   End
   Begin VB.Image imaLogo 
      Height          =   705
      Left            =   0
      Picture         =   "ctlDocumento.ctx":0015
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2235
   End
   Begin VB.Label lblDeSenor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Señores :"
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   780
      Width           =   795
   End
   Begin VB.Label lblSenor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
      Height          =   195
      Left            =   900
      TabIndex        =   3
      Top             =   780
      Width           =   6135
   End
   Begin VB.Label lbldeFecha 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fecha :"
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   300
      Width           =   615
   End
   Begin VB.Label lblLogo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Botica Torres de Limatambo S.A.C."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2280
      TabIndex        =   1
      Top             =   60
      Width           =   2655
   End
   Begin VB.Label lblFecha 
      BackColor       =   &H00FFFFFF&
      Caption         =   "dd/MM/yyyy HH:mm"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2280
      TabIndex        =   0
      Top             =   540
      Width           =   1575
   End
End
Attribute VB_Name = "ctlDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private strCia As String
Private strCodTipoDocumento As String
Private strNumDocumento As String
Private strCodLocal As String
Private strCodigo As String
Dim objDocumento As New clsDocumento
Dim rsCabDocumento As oraDynaset
Dim rsDetDocumento As oraDynaset
Public xCodigoMagistral As String

Public Property Get Cia() As String
    Cia = strCia
End Property

Public Property Let Cia(ByVal vstrCia As String)
    strCia = vstrCia
End Property

Public Property Get CodTipoDocumento() As String
    CodTipoDocumento = strCodTipoDocumento
End Property

Public Property Let CodTipoDocumento(ByVal vstrCodTipoDocumento As String)
    strCodTipoDocumento = vstrCodTipoDocumento
End Property

Public Property Get NumDocumento() As String
    NumDocumento = strNumDocumento
End Property

Public Property Let NumDocumento(ByVal vstrNumDocumento As String)
    strNumDocumento = vstrNumDocumento
End Property

Public Property Get CodLocal() As String
    CodLocal = strCodLocal
End Property

Public Property Let CodLocal(ByVal vstrCodLocal As String)
    strCodLocal = vstrCodLocal
End Property

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

Public Property Let Codigo(ByVal vstrCodigo As String)
    strCodigo = vstrCodigo
End Property

Public Sub Mostrar()
    On Error GoTo CtrlErr

    lblDeBoleta.Caption = devTipDoc(CodTipoDocumento)
    lblNumero.Caption = "Nº " & NumDocumento
    Set rsCabDocumento = objDocumento.ListaCabecera(Cia, Replace(NumDocumento, "-", ""), CodTipoDocumento, CodLocal)
    lblFecha.Caption = IIf(IsNull(rsCabDocumento("FCH_REGISTRA").Value), "", rsCabDocumento("FCH_REGISTRA").Value)
    lblSenor.Caption = IIf(IsNull(rsCabDocumento("DES_RAZON_SOCIAL").Value), IIf(IsNull(rsCabDocumento("DES_AUX_CLI_NOMBRE").Value), "", rsCabDocumento("DES_AUX_CLI_NOMBRE").Value), rsCabDocumento("DES_RAZON_SOCIAL").Value)
    lblTotal.Caption = Format(IIf(IsNull(rsCabDocumento("MTO_TOTAL").Value), 0, rsCabDocumento("MTO_TOTAL").Value), "##0.#0")
    'LblCajero.Caption = IIf(devNomUsuario(rsCabDocumento("COD_USUARIO_DEPENDIENTE").Value), "", devNomUsuario(rsCabDocumento("COD_USUARIO_DEPENDIENTE").Value))
    lblCajero.Caption = devNomUsuario(IIf(IsNull(rsCabDocumento("COD_USUARIO_DEPENDIENTE").Value), "", rsCabDocumento("COD_USUARIO_DEPENDIENTE").Value))
    'lblDependiente.Caption = devNomUsuario(rsCabDocumento("COD_USUARIO").Value)
    If rsCabDocumento("COD_ESTADO").Value = objUsuario.EstadoAnulado Then
        lblAnulada.Visible = True
        lblAdministrador.Visible = True
        lblAdministrador.Caption = devNomUsuario(rsCabDocumento("COD_USUARIO_ACTUALIZA").Value)
    End If
 
    'Set rsDetDocumento = objDocumento.ListaDetalle(Cia, NumDocumento, CodTipoDocumento, CodLocal)
    Set rsDetDocumento = objDocumento.LstDetDocumentoxTipo(Cia, CodTipoDocumento, NumDocumento)
    
    Call SetgrdDatos
    Set ORADC1.Recordset = rsDetDocumento
    
    rsDetDocumento.MoveFirst
    While Not rsDetDocumento.EOF
        If (rsDetDocumento("COD_PRODUCTO").Value = "86520") Or (rsDetDocumento("COD_PRODUCTO").Value = "90287") Then
            xCodigoMagistral = rsDetDocumento("COD_PRODUCTO").Value
        End If
        rsDetDocumento.MoveNext
    Wend
    
    Exit Sub
CtrlErr:
    Err.Raise Err.Number, "ctlDocumento.Mostrar", Err.Description
End Sub

'Public Property Get xCodigoMagistral() As String
'    xCodigoMagistral = rsDetDocumento("COD_PRODUCTO").Value
'End Property

Private Function devTipDoc(ByVal TipoDoc As String) As String
Dim rsTipoDoc As oraDynaset

    On Error GoTo CtrlErr
    Set rsTipoDoc = objDocumento.ListaTipoDocumento(TipoDoc)
    If rsTipoDoc.EOF Then
        devTipDoc = "*"
    Else
        devTipDoc = rsTipoDoc("DES_TIPODOC").Value
    End If
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "ctlDocumento.devTipDoc", Err.Description
End Function

Private Function devNomUsuario(ByVal CodUsuario As String) As String
Dim rsUsuario As oraDynaset
    ''arturo escate es porque cuando no hay codigo devuelve al primero de la lista
    If CodUsuario = "" Then devNomUsuario = "*": Exit Function
    On Error GoTo CtrlErr

    Set rsUsuario = objUsuario.Lista(CodUsuario)
    If rsUsuario.EOF Then
        devNomUsuario = "*"
    Else
    
        devNomUsuario = CodUsuario & " - " & rsUsuario("DES_NOMBRE").Value & ", " & rsUsuario("APE_PAT_USUARIO").Value & " " & rsUsuario("APE_MAT_USUARIO").Value
    End If

    Exit Function
CtrlErr:
    Err.Raise Err.Number, "ctlDocumento.devNomUsuario", Err.Description


End Function

Public Sub FormatoGrilla(ByVal arrayDataField As Variant, ByVal arrayCaption As Variant, _
                         ByVal arrarWidth As Variant, ByVal arrayAlignment As Variant)
    Dim i As Byte
    Dim Columna As TrueDBGrid70.Column

    grdDatos.Columns.Clear 'Quita todas las columnas
    For i = 0 To UBound(arrayDataField)
            Set Columna = grdDatos.Columns.Add(i)
        
        If Not IsMissing(arrarWidth) Then Columna.Width = arrarWidth(i)
        If Not IsMissing(arrayAlignment) Then Columna.Alignment = arrayAlignment(i)
        If Not IsMissing(arrayDataField) Then Columna.DataField = arrayDataField(i)
        Columna.Caption = arrayCaption(i)
        Columna.AllowSizing = True
        Columna.WrapText = True
        Columna.Visible = True
    Next i
    grdDatos.AllowAddNew = False
    grdDatos.AllowUpdate = False
    grdDatos.HoldFields
    
    grdDatos.Splits(0).AllowColSelect = False
    grdDatos.Splits(0).AllowRowSelect = True
    grdDatos.Splits(0).AllowRowSizing = False
    grdDatos.Splits(0).AllowSizing = False
    grdDatos.Splits(0).Style.VerticalAlignment = dbgVertCenter
    grdDatos.Appearance = dbgFlat
    grdDatos.HeadBackColor = &HFF0000
    grdDatos.HeadForeColor = &HFFFFFF
    grdDatos.RecordSelectors = False
    grdDatos.DeadAreaBackColor = &HFFFFFF
    If CodTipoDocumento = objUsuario.TipoDocBol Then
        grdDatos.BackColor = &HFFFFC0
    Else
        grdDatos.BackColor = &HC0C0FF
    End If
    grdDatos.RowDividerStyle = dbgNoDividers
    grdDatos.Columns(2).NumberFormat = "##"
    grdDatos.Columns(3).NumberFormat = "##"
    grdDatos.Columns(4).NumberFormat = "###0.00"
    grdDatos.Columns(5).NumberFormat = "###0.00"
    grdDatos.Columns(6).NumberFormat = "###0.00"
    grdDatos.FetchRowStyle = True
End Sub

Private Sub SetgrdDatos()
Dim arrCampos As Variant
Dim arrCaption As Variant
Dim arrAncho As Variant
Dim arrAlineacion As Variant
    arrCampos = Array("COD_PRODUCTO", "DES_PRODUCTO", _
                      "CTD_PRODUCTO", "CTD_PRODUCTO_FRAC", _
                      "PRC_UNIT_VTA", "PCT_DSCT_UNIT", _
                      "MTO_SUBTOTAL")
                      
    arrCaption = Array("Código", "Descripción", _
                       "Unid", "Fracc.", _
                       "Sub-Total", "Dscto", _
                       "Total")
                       
    arrAncho = Array(900, 2500, _
                     500, 500, _
                     800, 800, _
                     800)
                     
    arrAlineacion = Array(dbgLeft, dbgLeft, _
                          dbgCenter, dbgCenter, _
                          dbgRight, dbgRight, _
                          dbgRight)
                          
    FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
End Sub

Private Sub grdDatos_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid70.StyleDisp)
Dim CodProd As String
On Error GoTo handle
    CodProd = Val(grdDatos.Columns(0).CellText(Bookmark))
    If CodProd = Codigo Then
        RowStyle.ForeColor = &H0&
        RowStyle.BackColor = &HC0E0FF
        RowStyle.Font.Bold = True
    End If
Exit Sub
handle:
    Err.Raise Err.Description, vbCritical, App.ProductName
End Sub

