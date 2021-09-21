VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frm_ADM_Sincronizacion 
   BorderStyle     =   0  'None
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5685
      Picture         =   "frm_OFF_Sincronizacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Procesar"
      Height          =   615
      Left            =   4380
      Picture         =   "frm_OFF_Sincronizacion.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   1095
   End
   Begin TrueDBGrid70.TDBGrid TDBGUsers 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9128
      _LayoutType     =   4
      _RowHeight      =   19
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nombre"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Código Liquidación"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1402"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1323"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=5318"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=5239"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=238"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=159"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=4710"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=4630"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=344"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=265"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      TabAction       =   2
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
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
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
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
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Named:id=33:Normal"
      _StyleDefs(57)  =   ":id=33,.parent=0"
      _StyleDefs(58)  =   "Named:id=34:Heading"
      _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(60)  =   ":id=34,.wraptext=-1"
      _StyleDefs(61)  =   "Named:id=35:Footing"
      _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   "Named:id=36:Selected"
      _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=37:Caption"
      _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(67)  =   "Named:id=38:HighlightRow"
      _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=39:EvenRow"
      _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(71)  =   "Named:id=40:OddRow"
      _StyleDefs(72)  =   ":id=40,.parent=33"
      _StyleDefs(73)  =   "Named:id=41:RecordSelector"
      _StyleDefs(74)  =   ":id=41,.parent=34"
      _StyleDefs(75)  =   "Named:id=42:FilterBar"
      _StyleDefs(76)  =   ":id=42,.parent=33"
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
      Left            =   6030
      TabIndex        =   4
      Top             =   6180
      Width           =   390
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
      Left            =   4320
      TabIndex        =   3
      Top             =   6180
      Width           =   1215
   End
End
Attribute VB_Name = "frm_ADM_Sincronizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oSync As cls_OFF_Sincronizacion
Private xUsers As XArrayDB

Private Sub cmdAceptar_Click()
    Dim r As Integer, c As Integer
    Dim isOk As Boolean
    
    On Error GoTo ErrorHandler
    
    If TDBGUsers.ApproxCount <= 0 Then
        MsgBox "No existen datos para procesar.", vbCritical, App.ProductName
        Exit Sub
    End If
    
    isOk = False
    
    TDBGUsers.MoveNext
    TDBGUsers.MoveFirst
    
    'For r = 0 To TDBGUsers.ApproxCount
    While Not TDBGUsers.EOF
        If Len(Trim(TDBGUsers.Columns(3).Value)) > 0 Then
            isOk = oSync.ProcesarLiquidacion(Trim(TDBGUsers.Columns(0).Value), _
                                             Trim(TDBGUsers.Columns(3).Value), _
                                             Trim(TDBGUsers.Columns(2).Value))
        End If
        TDBGUsers.MoveNext
    Wend
    
    Call EnviarArchivos
    
    If isOk = True Then _
        MsgBox "Proceso concluido satisfactoriamente.", vbInformation, App.ProductName
    
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbCritical, App.ProductName
    Call ObtenerListaUsuarios
    Call FormatGrid
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler
    
    setteaFormulario Me

    Set oSync = New cls_OFF_Sincronizacion
    
    Call EnviarArchivos
    Call FormatGrid
    
final:
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbCritical, App.ProductName
    GoTo final
End Sub

Private Sub EnviarArchivos()
    On Error GoTo ErrorHandler
    
    'Enviar los archivos de contingencia a la central
    'y obtener los usuarios a liquidar
    Call oSync.EnviarArchivos
    Call ObtenerListaUsuarios
        
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbCritical, App.ProductName
'''    GoTo final
End Sub

'''Private Sub EnviarArchivos()
'''    Dim rs As oraDynaset
'''    Dim row As Long, col As Byte

'''    On Error GoTo ErrorHandler
    
    'Enviar los archivos de contingencia a la central
    'y obtener los usuarios a liquidar
'''    Set rs = oSync.EnviarArchivos
    
'''    If rs Is Nothing Then Exit Sub
'''
'''    Set xUsers = New XArrayDB
'''    xUsers.LoadRows rs.GetRows
'''    xUsers.ReDim 0, -1, 0, rs.Fields.Count
'''
'''    row = 0
'''    While Not rs.EOF
'''        xUsers.AppendRows
'''
'''        For col = 0 To rs.Fields.Count - 1
'''            xUsers.Value(row, col) = rs.Fields(rs.FieldName(col)).Value
'''        Next
'''
'''        row = row + 1
'''        rs.MoveNext
'''    Wend
    
'''final:
'''    Set rs = Nothing
'''    Exit Sub
'''
'''ErrorHandler:
'''    MsgBox Err.Description, vbCritical, App.ProductName
'''    GoTo final
'''End Sub

Private Sub ObtenerListaUsuarios()
    Dim rs As oraDynaset

    On Error GoTo ErrorHandler
    
    Set xUsers = New XArrayDB
    Set rs = oSync.ObtenerListaUsuarios()
    
    If rs Is Nothing Then Exit Sub
    If rs.RecordCount <= 0 Then
        xUsers.ReDim 0, -1, 0, rs.Fields.Count
        Exit Sub
    End If
    
    xUsers.LoadRows rs.GetRows

    Set rs = Nothing
    Exit Sub

ErrorHandler:
    Set rs = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub FormatGrid()
    Dim c As TrueDBGrid70.Column
    Dim odynLiquidaciones As oraDynaset
    
    On Error GoTo ErrorHandler
    
    
    For Each c In TDBGUsers.Columns
        c.HeadAlignment = dbgCenter
        c.HeadingStyle.WrapText = True
        
        Select Case c.ColIndex
            Case 0
                c.Visible = True
                c.Locked = True
                c.AllowSizing = False
            Case 1
                c.Visible = True
                c.Locked = True
                c.AllowSizing = False
            Case 2
                c.Visible = False
                c.Locked = True
                c.AllowSizing = False
            Case 3
                c.Visible = True
                c.Locked = False
                c.AllowSizing = True
            Case 4
                c.Visible = False
                c.Locked = True
                c.AllowSizing = False
        End Select
    Next
    
    With TDBGUsers
        .HeadLines = 2
        .MarqueeStyle = dbgDottedCellBorder
        .HighlightRowStyle.BackColor = vbBlue
        .FetchRowStyle = True
        .Array = xUsers
        .Refresh
        .Rebind
    End With
    
    Set odynLiquidaciones = objLiquidacion.ListaUsurioxArqueo(objUsuario.CodigoEmpresa, _
                                               objUsuario.CodigoLocal, _
                                               Format(DateAdd("d", -7, Now), "dd/mm/yyyy"), _
                                               Format(Now, "dd/mm/yyyy"), 1)

    Call mdlGrilla.spGrilla_CboBox(TDBGUsers, 3, "COD_LIQUIDACION", odynLiquidaciones, "NOMB")

final:
    Set c = Nothing
    Set odynLiquidaciones = Nothing
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbCritical, App.ProductName
    GoTo final
End Sub

Private Sub Form_Resize()
    TDBGUsers.Move 120, 120, ScaleWidth - 120, cmdAceptar.top - 120
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo ErrorHandler
    
    Set oSync = Nothing
    Set xUsers = Nothing

final:
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbCritical, App.ProductName
    GoTo final
End Sub

Private Sub TDBGUsers_AfterColUpdate(ByVal ColIndex As Integer)
    TDBGUsers.MoveNext
    TDBGUsers.MovePrevious
End Sub

Private Sub TDBGUsers_DblClick()
If TDBGUsers.Columns(4).Value <> "" Then
    frm_ADM_rptContigencia.strSecError = Trim(TDBGUsers.Columns(4).Value)
    frm_ADM_rptContigencia.Show
End If
End Sub
