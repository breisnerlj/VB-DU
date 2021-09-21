VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.UserControl ctlGrillaArray 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   5250
   Begin TrueDBGrid70.TDBGrid GrdArray 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6165
      _LayoutType     =   0
      _RowHeight      =   16
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
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   0
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=-1,.fontsize=825,.italic=0"
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
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(30)  =   ":id=18,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(31)  =   ":id=18,.fontname=Arial"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Named:id=33:Normal"
      _StyleDefs(47)  =   ":id=33,.parent=0,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0"
      _StyleDefs(48)  =   ":id=33,.charset=0"
      _StyleDefs(49)  =   ":id=33,.fontname=Arial"
      _StyleDefs(50)  =   "Named:id=34:Heading"
      _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H80000015&,.fgcolor=&H80000014&"
      _StyleDefs(52)  =   ":id=34,.wraptext=-1"
      _StyleDefs(53)  =   "Named:id=35:Footing"
      _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(55)  =   "Named:id=36:Selected"
      _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=37:Caption"
      _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(59)  =   "Named:id=38:HighlightRow"
      _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&HFEEBDE&,.fgcolor=&H80000012&,.bold=0,.fontsize=825"
      _StyleDefs(61)  =   ":id=38,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(62)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(63)  =   "Named:id=39:EvenRow"
      _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(65)  =   "Named:id=40:OddRow"
      _StyleDefs(66)  =   ":id=40,.parent=33"
      _StyleDefs(67)  =   "Named:id=41:RecordSelector"
      _StyleDefs(68)  =   ":id=41,.parent=34"
      _StyleDefs(69)  =   "Named:id=42:FilterBar"
      _StyleDefs(70)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "ctlGrillaArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Click()
Public Event DblClick()
Public Event AfterInsert()
Public Event RegistroSeleccionado(ByVal DatoColumna0 As String)
Public Event DblClickRegistro(ByVal DatoColumna0 As String)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event FetchCellTips(ByVal SplitIndex As Integer, ByVal ColIndex As Integer, ByVal RowIndex As Long, CellTip As String, ByVal FullyDisplayed As Boolean, ByVal TipStyle As TrueDBGrid70.StyleDisp)
Public Event KeyPress(KeyAscii As Integer)
Public Event BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Public Event BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Public Event AfterColUpdate(ByVal ColIndex As Integer)
Public Event HeadClick(ByVal ColIndex As Integer)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event FirstRowChange(ByVal SplitIndex As Integer)

Public Event ButtonClick(ByVal ColIndex As Integer)

Public Event FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)

Private strClave As String
Private lngRegistros As Long
Private blnModoSort As Boolean
Private blnMenuPopUp As Boolean
Private blnResalte As Boolean

Private Sub GrdArray_AfterInsert()
    RaiseEvent AfterInsert
End Sub

Private Sub GrdArray_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    RaiseEvent BeforeColEdit(ColIndex, KeyAscii, Cancel)
End Sub

Private Sub GrdArray_ButtonClick(ByVal ColIndex As Integer)
    RaiseEvent ButtonClick(ColIndex)
End Sub

Private Sub GrdArray_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid70.StyleDisp)
    RaiseEvent FetchCellStyle(Condition, Split, Bookmark, Col, CellStyle)
End Sub

Private Sub GrdArray_FetchCellTips(ByVal SplitIndex As Integer, ByVal ColIndex As Integer, ByVal RowIndex As Long, CellTip As String, ByVal FullyDisplayed As Boolean, ByVal TipStyle As TrueDBGrid70.StyleDisp)
    RaiseEvent FetchCellTips(SplitIndex, ColIndex, RowIndex, CellTip, FullyDisplayed, TipStyle)
End Sub

Private Sub GrdArray_FirstRowChange(ByVal SplitIndex As Integer)
    RaiseEvent FirstRowChange(SplitIndex)
End Sub

'Public Enum TipoControlCelda
'    Normal = 1
'    Mayusculas = 2
'    Entero = 3
'    Real = 4
'    Telefono = 5
'    Porcentaje = 6
'    Documento = 7
'    AlfaNumerico = 8
'End Enum

'Private pTipo As TipoControlCelda

Private Sub GrdArray_GotFocus()
    If blnResalte = False Then
    End If
End Sub

Private Sub GrdArray_HeadClick(ByVal ColIndex As Integer)
    RaiseEvent HeadClick(ColIndex)
End Sub

Private Sub GrdArray_LostFocus()
    If blnResalte = False Then
    End If
End Sub

Private Sub GrdArray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'En esta seccion se va disparar un evento cada vez que se seleccione una
'nueva fila y retorna el dato de la columna0
Private Sub GrdArray_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    strClave = GrdArray.Columns(0).Text
    RaiseEvent RegistroSeleccionado(strClave)
End Sub

Public Sub ExportToFile(ByVal Path As String, Append As Boolean)
    GrdArray.ExportToFile Path, Append
End Sub

Public Function IsSelected(ByVal Bookmark As Variant) As Integer
    IsSelected = GrdArray.IsSelected(Bookmark)
End Function

Public Sub FormatoGrilla(ByVal arrayDataField As Variant, ByVal arrayCaption As Variant, _
                         ByVal arrarWidth As Variant, ByVal arrayAlignment As Variant, _
                         Optional ByVal arrayFoco As Variant)
    Dim i As Byte
    Dim Columna As TrueDBGrid70.Column
    
    GrdArray.Columns.Clear
    For i = 0 To UBound(arrayDataField)
        Set Columna = GrdArray.Columns.Add(i)
        If Not IsMissing(arrarWidth) Then Columna.Width = arrarWidth(i)
        If Not IsMissing(arrayAlignment) Then Columna.Alignment = arrayAlignment(i)
        Columna.Caption = arrayCaption(i)
        Columna.AllowSizing = True
        Columna.WrapText = True
        Columna.Visible = True
        Columna.AllowFocus = True
        If Not IsMissing(arrayFoco) Then Columna.AllowFocus = arrayFoco(i)
    Next i
    GrdArray.AllowAddNew = False
    GrdArray.AllowUpdate = False
    GrdArray.HoldFields
    GrdArray.Splits(0).AllowColSelect = False
    GrdArray.Splits(0).AllowRowSelect = True
    GrdArray.Splits(0).AllowRowSizing = False
    GrdArray.Splits(0).AllowSizing = False
    GrdArray.Splits(0).Style.VerticalAlignment = dbgVertCenter
    GrdArray.MarqueeStyle = dbgHighlightRowRaiseCell
    
End Sub

Public Function Limpiar()
    GrdArray.Close True
End Function

Public Property Get Columns() As TrueDBGrid70.Columns
   Set Columns = GrdArray.Columns
End Property

Public Property Let Columns(ByVal vNewValue As TrueDBGrid70.Columns)
    GrdArray.Columns.vNewValue
End Property

'------------ Eventos clasicos -------------------
Private Sub GrdArray_DblClick()
    RaiseEvent DblClick
    RaiseEvent DblClickRegistro(strClave)
End Sub
Private Sub GrdArray_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub GrdArray_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub GrdArray_Click()
    RaiseEvent Click
End Sub

Private Sub GrdArray_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    RaiseEvent BeforeColUpdate(ColIndex, OldValue, Cancel)
End Sub
    
Private Sub GrdArray_AfterColUpdate(ByVal ColIndex As Integer)
    RaiseEvent AfterColUpdate(ColIndex)
End Sub

'End Sub

'Private Sub BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer, Optional Tipo As TipoControlCelda = 1)
'    RaiseEvent BeforeColUpdate(ColIndex, OldValue, Cancel)
'        Select Case Tipo
'              Case 4
'                  Real
'                  If Not IsNumeric(OldValue) And OldValue <> "" Then
'                     'MsgBox "El valor ingreso no es válido", vbExclamation, "Error"
'                     Cancel = True
'                  ElseIf (Trim(GrdDeposito.Text) <> "") And (Trim(GrdDeposito.Text) > 0) Then
'                     GrdDeposito.Text = Format(Trim(GrdDeposito.Text), "#0.00")
'                  End If
'        End Select
'End Sub


Public Sub Refresh()
    GrdArray.Refresh
End Sub

Public Sub Rebind()
    GrdArray.Rebind
End Sub

Public Sub MoveFirst()
    GrdArray.MoveFirst
End Sub

Public Sub MoveLast()
    GrdArray.MoveLast
End Sub

Public Sub MoveNext()
    GrdArray.MoveNext
End Sub

Public Sub MovePrevious()
    GrdArray.MovePrevious
End Sub

Public Sub Delete()
    GrdArray.Delete
End Sub

'----------- Propiedades -------------------------------------
Public Property Get Clave() As String
    Clave = strClave
End Property

Public Property Get Bookmark() As Variant
    Bookmark = GrdArray.Bookmark
End Property

Public Property Let Bookmark(ByVal vNewValue As Variant)
    GrdArray.Bookmark = vNewValue
End Property

Public Property Get row() As Integer
    row = GrdArray.row
End Property
Public Property Let row(ByVal vNewValue As Integer)
    GrdArray.row = vNewValue
End Property

'--------------------------------------------------------------'
'--- cambios por crueda 11/10/2006 ---'
Public Property Get RowHeight()
    RowHeight = GrdArray.RowHeight
End Property

Public Property Let RowHeight(ByVal vNewValue As Variant)
    GrdArray.RowHeight = vNewValue
End Property

'--------------------------------------------------------------'
Public Property Get MultiSelect() As TrueDBGrid70.MultiSelectConstants
    MultiSelect = GrdArray.MultiSelect
End Property

Public Property Let MultiSelect(ByVal vNewValue As TrueDBGrid70.MultiSelectConstants)
    GrdArray.MultiSelect = vNewValue
End Property

Public Property Get ApproxCount() As Long
    ApproxCount = GrdArray.ApproxCount
End Property

Public Property Let ApproxCount(ByVal vNewValue As Long)
    GrdArray.ApproxCount = vNewValue
End Property

Public Property Get Enabled() As Boolean
    Enabled = GrdArray.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    GrdArray.Enabled = vNewValue
End Property

Public Property Get MenuPopUp() As Boolean
    MenuPopUp = blnMenuPopUp
End Property

Public Property Let MenuPopUp(ByVal vNewValue As Boolean)
    blnMenuPopUp = vNewValue
End Property

Public Property Get Resalte() As Boolean
    Resalte = blnResalte
End Property

Public Property Let Resalte(ByVal vNewValue As Boolean)
    blnResalte = vNewValue
End Property

Public Property Get ColumnHeaders() As Boolean
    ColumnHeaders = GrdArray.ColumnHeaders
End Property

Public Property Let ColumnHeaders(ByVal vNewValue As Boolean)
    GrdArray.ColumnHeaders = vNewValue
End Property

Public Property Get ColumnFooter() As Boolean
    ColumnFooter = GrdArray.ColumnFooters
End Property

Public Property Let ColumnFooter(ByVal vNewValue As Boolean)
    GrdArray.ColumnFooters = vNewValue
End Property

'---------------------------------------------------------------
Private Sub GrdArray_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And blnMenuPopUp Then
        'PopupMenu mnuMenu
    End If
End Sub

'Private Sub mnuExcel_Click()
'    On Error GoTo Control
'    Dim Ret As Variant
'    dlgArchivo.CancelError = True
'    dlgArchivo.DialogTitle = "Ingrese el nombre del archivo."
'    dlgArchivo.ShowSave
'    Ret = dlgArchivo.FileName
'    If Trim(Ret) = "" Then MsgBox "No se encontro el archivo.", vbOKOnly + vbExclamation, "Error": Exit Sub
'    If Dir(Ret) <> "" Then
'        If MsgBox("Ya existe un archivo con el nombre especificado." & Chr(13) & "Desea continuar ?", vbYesNo + vbQuestion, "Confirme") = vbNo Then Exit Sub
'    End If
'        GrdArray.ExportToFile Ret, False
'    MsgBox "Se exportó el archivo correctamente.", vbOKOnly + vbInformation, "Mensaje"
'    Exit Sub
'Control:
'    Select Case Err.Number
'    Case 32755
'    Case Else
'        MsgBox Err.Description, vbOKOnly + vbExclamation, "Error : " & Err.Number
'    End Select
'End Sub
Private Sub mnuImprimir_Click()
    GrdArray.PrintInfo.PrintPreview
End Sub

Public Property Let PrintInfo(ByVal mvar As TrueDBGrid70.PrintInfo)
    GrdArray.PrintInfo.mvar
End Property
Public Property Get PrintInfo() As TrueDBGrid70.PrintInfo
    Set PrintInfo = GrdArray.PrintInfo
End Property



'Private Sub mnuEmail_Click()
'    On Error Resume Next
'    With MAPISession1
'       .SignOff
'       .LogonUI = True
'       .DownLoadMail = False
'       .SignOn
'       .NewSession = True
'        MAPIMessages1.SessionID = .SessionID
'     MAPIMessages1.SessionID = .SessionID
'    End With
'
'    With MAPIMessages1
'        GrdArray.ExportToFile "c:\temp.xls", False
'        .Compose
'        '.RecipAddress = ""
'        .AddressResolveUI = True
'        .AttachmentPathName = "c:\temp.xls"
'        .ResolveName
'        '.MsgSubject = ""
'        .MsgNoteText = " "
'        .Send True
'        Kill "c:\temp.xls" 'Borra archivo temporal
'    End With
'End Sub

'Public Sub MostrarExcel()
'    mnuExcel_Click
'End Sub
'Public Sub MostrarEmail()
'    mnuEmail_Click
'End Sub
'Public Sub MostrarImprimir()
'    mnuImprimir_Click
'End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", GrdArray.Enabled, True
    PropBag.WriteProperty "MenuPopUp", blnMenuPopUp, True
    PropBag.WriteProperty "Resalte", blnResalte, True
    PropBag.WriteProperty "ColumnHeaders", GrdArray.ColumnHeaders, True
    PropBag.WriteProperty "MultiSelect", GrdArray.MultiSelect, 1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    GrdArray.Enabled = PropBag.ReadProperty("Enabled", True)
    blnMenuPopUp = PropBag.ReadProperty("MenuPopUp", True)
    blnResalte = PropBag.ReadProperty("Resalte", True)
    GrdArray.ColumnHeaders = PropBag.ReadProperty("ColumnHeaders", True)
    GrdArray.MultiSelect = PropBag.ReadProperty("MultiSelect", 1)
End Sub

Private Sub UserControl_Resize()
    GrdArray.Width = UserControl.Width
    GrdArray.Height = UserControl.Height
    GrdArray.Width = UserControl.Width
    GrdArray.Height = UserControl.Height
End Sub

Public Property Let Array1(ByVal Xarray As XArrayDB)
    GrdArray.Array = Xarray
    GrdArray.Rebind
End Property

Public Property Get Array1() As XArrayDB
    Set Array1 = GrdArray.Array
End Property

Public Property Get Col() As Variant
    Col = GrdArray.Col
End Property

Public Property Let Col(ByVal vNewValue As Variant)
    GrdArray.Col = vNewValue
End Property

Public Property Get MarqueeStyle() As Variant
    MarqueeStyle = GrdArray.MarqueeStyle
End Property

Public Property Let MarqueeStyle(ByVal vNewValue As Variant)
    GrdArray.MarqueeStyle = vNewValue
End Property

Public Property Get AllowUpdate() As Boolean
    AllowUpdate = GrdArray.AllowUpdate
End Property

Public Property Let AllowUpdate(ByVal vNewValue As Boolean)
    GrdArray.AllowUpdate = vNewValue
End Property

Public Property Get HeadLines() As Integer
    HeadLines = GrdArray.HeadLines
End Property

Public Property Let HeadLines(ByVal vNewValue As Integer)
    GrdArray.HeadLines = vNewValue
End Property

Public Property Get CellTips() As Integer
    CellTips = GrdArray.CellTips
End Property

Public Property Let CellTips(ByVal vNewValue As Integer)
    GrdArray.CellTips = vNewValue
End Property

Public Property Get CellTipsWidth() As Single
    CellTipsWidth = GrdArray.CellTipsWidth
End Property

Public Property Let CellTipsWidth(ByVal vNewValue As Single)
    GrdArray.CellTipsWidth = vNewValue
End Property

Public Sub Update()
    GrdArray.Update
End Sub

Public Function RowBookmark(ByVal RowNum As Long) As Variant
     RowBookmark = GrdArray.RowBookmark(RowNum)
End Function

'---- Creado por CRUEDA 20/07/2007
'----------------------------------------------------------------
Public Property Get StylesFondo() As Single
     StylesFondo = GrdArray.Styles(5).Font.Bold = True
End Property

Public Property Let StylesFondo(ByVal vNewValue As Single)
    GrdArray.Styles(5).Font.Bold = True = vNewValue
End Property

Public Property Get StylesSize() As Variant
    GrdArray.Styles(5).Font.Size = 12
End Property

Public Property Let StylesSize(ByVal vNewValue As Variant)
    GrdArray.Styles(5).Font.Size = vNewValue
End Property

Public Function CambiaSeleccionadoBackColor(Fondo As Variant)
    GrdArray.Styles(5).BackColor = Fondo
End Function

Public Function CambiaSeleccionadoForeColor(Fondo As Variant)
    GrdArray.Styles(5).ForeColor = Fondo
End Function

'---- Creado por JRAZURI 24/05/2008
'----------------------------------------------------------------
Public Property Get EditorStyle() As Style
    Set EditorStyle = GrdArray.EditorStyle
End Property

Public Property Set EditorStyle(ByVal vNewValue As Style)
    Set GrdArray.EditorStyle = vNewValue
End Property

Public Property Get Splits() As Splits
    Set Splits = GrdArray.Splits
End Property

Public Property Set Splits(ByVal vNewValue As Splits)
    GrdArray.Splits.vNewValue
End Property

Public Property Get EOF() As Boolean
    EOF = GrdArray.EOF
End Property

Public Property Get BOF() As Boolean
    BOF = GrdArray.BOF
End Property

Public Property Get Styles() As Styles
    Set Styles = GrdArray.Styles
End Property

Public Property Set Styles(ByVal vNewValue As Styles)
    GrdArray.Styles.vNewValue
End Property

Public Property Get EditActive() As Boolean
    EditActive = GrdArray.EditActive
End Property

Public Property Let EditActive(ByVal vNewValue As Boolean)
    GrdArray.EditActive = vNewValue
End Property

'---- Creado por CCIEZA 25/09/2009
'----------------------------------------------------------------
Public Property Get CheckedAll(ByVal varCol As Variant) As CheckBoxConstants
Dim varChecked  As CheckBoxConstants
Dim xdb As XArrayDB
Dim intIndex As Integer
Dim i As Integer
        
    If GrdArray.ApproxCount = 0 Then
        varChecked = vbUnchecked
    Else
            
        intIndex = GrdArray.Columns(varCol).ColIndex
        
        varChecked = vbGrayed
        
        Set xdb = GrdArray.Array
        For i = xdb.LowerBound(1) To xdb.UpperBound(1)
            If Abs(xdb(i, intIndex)) = 1 And varChecked = vbGrayed Then
                varChecked = vbChecked
            ElseIf Not (Abs(xdb(i, intIndex)) = 1) And varChecked = vbGrayed Then
                varChecked = vbUnchecked
            ElseIf Abs(xdb(i, intIndex)) = 1 And varChecked = vbChecked Then
                varChecked = vbChecked
            ElseIf Not (Abs(xdb(i, intIndex)) = 1) And varChecked = vbUnchecked Then
                varChecked = vbUnchecked
            Else
                varChecked = vbGrayed
                Exit For
            End If
        Next i
        
    End If
    CheckedAll = varChecked
End Property

Public Property Let CheckedAll(ByVal varCol As Variant, ByVal vChecked As CheckBoxConstants)
Dim xdb As XArrayDB
Dim intIndex As Integer
Dim i As Integer
    
    If vChecked = vbChecked Or vChecked = vbUnchecked Then
        intIndex = GrdArray.Columns(varCol).ColIndex
        
        'GrdArray.MoveNext
        'GrdArray.MovePrevious
        
        Set xdb = GrdArray.Array
        
        For i = xdb.LowerBound(1) To xdb.UpperBound(1)
            xdb(i, intIndex) = IIf(vChecked = vbChecked, "1", "0")
        Next i
        
        'GrdArray.Array = xdb
        
        GrdArray.Columns(varCol).Refetch
    End If
End Property

'---- Creado por CCIEZA 27/10/2008
'----------------------------------------------------------------
Public Property Get FetchRowStyle() As Boolean
    FetchRowStyle = GrdArray.FetchRowStyle
End Property

Public Property Let FetchRowStyle(ByVal blnFetchRowStyle As Boolean)
    GrdArray.FetchRowStyle = blnFetchRowStyle
End Property

Public Sub ReOpen()
    GrdArray.ReOpen
End Sub

Public Property Get SelBookmarks() As TrueDBGrid70.SelBookmarks
    Set SelBookmarks = GrdArray.SelBookmarks
End Property

Public Property Let SelBookmarks(ByVal vNew As TrueDBGrid70.SelBookmarks)
    GrdArray.SelBookmarks.vNew
End Property

