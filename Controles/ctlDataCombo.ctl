VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.UserControl ctlDataCombo 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2640
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   330
   ScaleWidth      =   2640
   Begin MSDBCtls.DBList DBList1 
      Bindings        =   "ctlDataCombo.ctx":0000
      Height          =   1425
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2223
      _Version        =   393216
   End
   Begin ORADCLibCtl.ORADC ORADC1 
      Height          =   375
      Left            =   1200
      Top             =   480
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   661
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
   Begin MSDBCtls.DBCombo dcDato 
      Bindings        =   "ctlDataCombo.ctx":0015
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
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
End
Attribute VB_Name = "ctlDataCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pColorF As OLE_COLOR, pColorD As OLE_COLOR
Private pblnFoco As Boolean, pblnTabAuto As Boolean
Public Event Change()
Public Event Click(Area As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Private bytFoco As Byte, lblnKey As Boolean

Public Property Let ColorFoco(ByVal vColor As OLE_COLOR)
    pColorF = vColor
End Property
Public Property Get ColorFoco() As OLE_COLOR
    ColorFoco = pColorF
End Property

Public Property Let ColorDefault(ByVal vColor As OLE_COLOR)
    pColorD = vColor
    dcDato.BackColor = vColor
End Property
Public Property Get ColorDefault() As OLE_COLOR
    ColorDefault = dcDato.BackColor
End Property

Private Sub dcDato_Change()
    PropertyChanged "BoundText"
    PropertyChanged "Text"
    RaiseEvent Change
    'If bytFoco = 1 Then dcDato.Font.Bold = True
    bytFoco = 0
End Sub

Private Sub dcDato_Click(Area As Integer)
    Dim temp As Integer
    temp = Area
    If lblnKey Then temp = 2
    RaiseEvent Click(temp)
End Sub

Private Sub dcDato_GotFocus()
    If pblnFoco = True Then
        bytFoco = 1
        dcDato.BackColor = pColorF
        dcDato.Refresh
    End If
End Sub

Private Sub dcDato_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    lblnKey = False
    If KeyCode = 40 Or KeyCode = 38 Then lblnKey = True
    If KeyCode = 13 And pblnTabAuto = True Then SendKeys "{TAB}"
End Sub

Private Sub dcDato_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub



Private Sub dcDato_LostFocus()
    If pblnFoco = True Then
        dcDato.BackColor = pColorD
        dcDato.Font.Bold = False
        dcDato.Refresh
    End If
End Sub

Public Property Let Text(ByVal vText As String)
    dcDato.Text = vText
End Property
Public Property Get Text() As String
    Text = dcDato.Text
End Property

Public Property Let BoundText(ByVal vText As String)
    dcDato.BoundText = vText
End Property
Public Property Get BoundText() As String
    BoundText = dcDato.BoundText
End Property

Public Property Get ListField() As String
    ListField = dcDato.ListField
End Property
Public Property Let ListField(ByVal mvar As String)
    dcDato.ListField = mvar
End Property

Public Property Get BoundColumn() As String
    BoundColumn = dcDato.BoundColumn
End Property
Public Property Let BoundColumn(ByVal mvar As String)
    dcDato.BoundColumn = mvar
    DBList1.BoundColumn = mvar
End Property

Public Property Let ListField2(ByVal mvar As String)
    DBList1.ListField = mvar
End Property
Public Property Get BoundText2() As String
    DBList1.BoundText = dcDato.BoundText
    BoundText2 = DBList1.Text
End Property

Public Property Get RowSource() As OracleInProcServer.oraDynaset
    Set RowSource = dcDato.RowSource
End Property
Public Property Set RowSource(mvar As OracleInProcServer.oraDynaset)
    Set ORADC1.Recordset = mvar
    'Set dcDato.RowSource = mvar
    Set mvar = Nothing
End Property

Public Property Get MatchEntry() As MSDBCtls.MatchEntryConstants
    MatchEntry = dcDato.MatchEntry
End Property
Public Property Let MatchEntry(mvar As MSDBCtls.MatchEntryConstants)
    dcDato.MatchEntry = mvar
End Property

Public Property Let EnabledFoco(ByVal vModo As Boolean)
    pblnFoco = vModo
End Property
Public Property Get EnabledFoco() As Boolean
    EnabledFoco = pblnFoco
End Property

Public Property Let Enabled(ByVal vModo As Boolean)
    dcDato.Enabled = vModo
    UserControl.Enabled = vModo
End Property
Public Property Get Enabled() As Boolean
    Enabled = dcDato.Enabled
End Property

Public Property Let TABAuto(ByVal vModo As Boolean)
    pblnTabAuto = vModo
End Property
Public Property Get TABAuto() As Boolean
    TABAuto = pblnTabAuto
End Property

 
 
Private Sub UserControl_Initialize()
    ColorFoco = &HC0C0FF
    ColorDefault = vbWhite
    pblnFoco = True
    pblnTabAuto = True
    lblnKey = False
End Sub

Private Sub UserControl_Resize()
    dcDato.Width = UserControl.Width
    UserControl.Height = dcDato.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ColorFoco", pColorF, &HC0C0FF
    PropBag.WriteProperty "ColorDefault", pColorD, vbWhite
    PropBag.WriteProperty "ColorDefault", dcDato.BackColor, vbWhite

    PropBag.WriteProperty "ListField", dcDato.ListField, ""
    PropBag.WriteProperty "BoundColumn", dcDato.BoundColumn, ""
    PropBag.WriteProperty "MatchEntry", dcDato.MatchEntry, 0
    PropBag.WriteProperty "EnabledFoco", pblnFoco, True
    PropBag.WriteProperty "Enabled", dcDato.Enabled, True
    PropBag.WriteProperty "TABAuto", pblnTabAuto, True
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    pColorF = PropBag.ReadProperty("ColorFoco", &HC0C0FF)
    pColorD = PropBag.ReadProperty("ColorDefault", vbWhite)
    dcDato.BackColor = PropBag.ReadProperty("ColorDefault", vbWhite)

    dcDato.ListField = PropBag.ReadProperty("ListField", "")
    dcDato.BoundColumn = PropBag.ReadProperty("BoundColumn", "")
    dcDato.MatchEntry = PropBag.ReadProperty("MatchEntry", 0)
    pblnFoco = PropBag.ReadProperty("EnabledFoco", True)
    Enabled = PropBag.ReadProperty("Enabled", True)
    pblnTabAuto = PropBag.ReadProperty("TABAuto", True)
End Sub


Public Property Let Bloqueado(mvar As Boolean)
    If mvar = True Then
        UserControl.Enabled = False
        dcDato.BackColor = &H8000000F
    Else
        UserControl.Enabled = True
        dcDato.BackColor = pColorD
    End If
End Property

