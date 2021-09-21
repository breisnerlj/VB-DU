VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlListCombo 
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   ForwardFocus    =   -1  'True
   ScaleHeight     =   3330
   ScaleWidth      =   3750
   Begin MSComctlLib.ListView Lista 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "ctlListCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Codigo As String
Private Descripcion As String
Private Marca As String
Private rs As OraDynaset
Public Event DblClick()
Public Event Click()
Public Event RegistroSeleccionado(Item As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Private Sub Lista_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Lista_Click()
    RaiseEvent Click
End Sub

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    RaiseEvent RegistroSeleccionado(Item.Index)
End Sub

Private Sub Lista_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Initialize()
    BoundColumn2 = "COD"
    ListField2 = "DES"
End Sub

Private Sub UserControl_Resize()
    Lista.Width = UserControl.Width
    Lista.Height = UserControl.Height
    Lista.ColumnHeaders.Item(1).Width = IIf(Lista.Width - 100 < 0, 0, Lista.Width - 100)
End Sub

Public Property Get RowSource() As OraDynaset
    Set RowSource = rs
End Property

Public Property Set RowSource(mvar As OraDynaset)
   Set rs = mvar
   Dim i As Integer, j As Integer
   Dim Item As String
   Dim LItem As ListItem
   i = rs.FieldIndex(Codigo)
   j = rs.FieldIndex(Descripcion)
    While Not rs.EOF
        Item = CStr(rs(i).Value) & "-" & CStr(rs(j).Value)
        
        Set LItem = Lista.ListItems.Add(, , Item)
        
        If Check And Marca <> "" Then
            Dim k As Integer
            k = rs.FieldIndex(Marca)
            LItem.Checked = ("1" = Trim(CStr(rs(k).Value)))
        End If
        rs.MoveNext
    Wend
   Set mvar = Nothing
End Property
'-------------------------------------
'-- JRAZURI 11/07/2007 - ADICIONE ESTA PROPIEDAD YA QUE LA ANTERIOR SE USA ASI COMO ESTA
'-------------------------------------
Public Property Get DataSource() As OraDynaset
    Set DataSource = rs
End Property

Public Property Set DataSource(mvar As OraDynaset)
   Set rs = mvar
   Dim i As Integer, j As Integer
   Dim Item As String
   Dim LItem As ListItem
    While Not rs.EOF
        Item = CStr(rs(Descripcion).Value)
        
        Set LItem = Lista.ListItems.Add(, "*" & rs(Codigo).Value, Item)
        
        If Check And Marca <> "" Then
            Dim k As Integer
            k = rs.FieldIndex(Marca)
            LItem.Checked = ("1" = Trim(CStr(rs(k).Value)))
        End If
        rs.MoveNext
    Wend
   Set mvar = Nothing
End Property

Public Property Get ListField2() As String
    ListField2 = Descripcion
End Property
Public Property Let ListField2(mvar As String)
    Descripcion = mvar
End Property

Public Property Get BoundColumn2() As String
    BoundColumn2 = Codigo
End Property
Public Property Let BoundColumn2(mvar As String)
    Codigo = mvar
End Property

'-------------------------------------
'-------------------------------------
Public Property Get ListField() As String
    ListField = Codigo
End Property
Public Property Let ListField(mvar As String)
    Codigo = mvar
End Property

Public Property Get BoundColumn() As String
    BoundColumn = Descripcion
End Property
Public Property Let BoundColumn(mvar As String)
    Descripcion = mvar
End Property

Public Property Get CheckColumn() As String
    CheckColumn = Marca
End Property
Public Property Let CheckColumn(mvar As String)
    Marca = mvar
End Property

Public Property Get ListCount() As Integer
    ListCount = Lista.ListItems.Count
End Property

Public Function Selected(Index) As Boolean
    Selected = Lista.ListItems(Index).Selected
End Function

Public Function List(Index) As String
    List = Lista.ListItems(Index).Text
End Function

Sub AddItem(Item As String, Optional Index As Integer = 0)
    If Index = 0 Then
        Lista.ListItems.Add , , Item
    Else
        Lista.ListItems.Add Index, , Item
    End If
End Sub

Sub RemoveItem(Index As Integer)
    Lista.ListItems.Remove Index
End Sub

Public Property Get Check() As Boolean
    Check = Lista.CheckBoxes
End Property

Public Property Let Check(mvar As Boolean)
    Lista.CheckBoxes = mvar
    Lista.Refresh
    PropertyChanged "Check"
End Property

''Public Property Get Pre_Selected() As String
''    Pre_Selected = List2.ItemData(List2.ListIndex)
''End Property
''Public Property Let Pre_Selected(mvar As String)
''   List2.ItemData(List2.ListIndex) = mvar
''End Property

Public Sub Clear()
    Lista.ListItems.Clear
End Sub

Public Property Get ListItem(Index) As ListItem
    Set ListItem = Lista.ListItems(Index)
End Property

'Public Property Let ListItem(Index, mvar As ListItem)
'    Set Lista.ListItems(Index) = mvar
'End Property

Public Property Get ListItems() As ListItems
    Set ListItems = Lista.ListItems
End Property

'Public Property Let ListItems(mvar As ListItems)
'    Set Lista.ListItems = mvar
'End Property

Public Property Get CheckedAll() As CheckBoxConstants
Dim Item As ListItem
Dim varChecked  As CheckBoxConstants
        
    varChecked = vbGrayed
        
    For Each Item In ListItems
        If Item.Checked And varChecked = vbGrayed Then
            varChecked = vbChecked
        ElseIf Not Item.Checked And varChecked = vbGrayed Then
            varChecked = vbUnchecked
        ElseIf Item.Checked And varChecked = vbChecked Then
            varChecked = vbChecked
        ElseIf Not Item.Checked And varChecked = vbUnchecked Then
            varChecked = vbUnchecked
        Else
            varChecked = vbGrayed
            Exit For
        End If
        
    Next Item
    
    
    If ListItems.Count = 0 Then
        varChecked = vbUnchecked
    End If
    CheckedAll = varChecked
End Property

Public Property Let CheckedAll(ByVal vChecked As CheckBoxConstants)
Dim Item As ListItem
    If vChecked = vbChecked Or vChecked = vbUnchecked Then
        For Each Item In ListItems
            Item.Checked = (vChecked = vbChecked)
        Next Item
    End If
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Check", Lista.CheckBoxes, False
    PropBag.WriteProperty "BoundColumn2", Codigo, "COD"
    PropBag.WriteProperty "ListField2", Descripcion, "DES"
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Lista.CheckBoxes = PropBag.ReadProperty("Check", False)
    Codigo = PropBag.ReadProperty("BoundColumn2", "COD")
    Descripcion = PropBag.ReadProperty("ListField2", "DES")
End Sub
