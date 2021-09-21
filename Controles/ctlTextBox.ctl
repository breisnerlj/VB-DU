VERSION 5.00
Begin VB.UserControl ctlTextBox 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4035
   ScaleHeight     =   315
   ScaleWidth      =   4035
   Begin VB.TextBox txtDato 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4035
   End
End
Attribute VB_Name = "ctlTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pColorF As OLE_COLOR, pColorD As OLE_COLOR
Private pBloqueado As Boolean

Public Enum TipoControlTextBox
    Normal = 1
    Mayusculas = 2
    Entero = 3
    Real = 4
    Telefono = 5
    Porcentaje = 6
    Documento = 7
    AlfaNumerico = 8
End Enum

Public Enum TipoAlignment
    tLeft = 0
    tRight = 1
    tCenter = 2
End Enum

Private pTipo As TipoControlTextBox 'Define el tipo de datos que acepta el textbox
Private pblnFoco As Boolean, pblnTabAuto As Boolean, pblnTipoSQL As Boolean

Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event Change()
Public Event Click()




'Entorno al recibir el foco
Private Sub txtDato_GotFocus()
    If pblnFoco = True Then
        txtDato.BackColor = pColorF
        txtDato.FontBold = True
        txtDato.SelStart = 0
        txtDato.SelLength = Len(txtDato.Text)
    End If
End Sub


'Entorno al dejar el control
Private Sub txtDato_LostFocus()
    If pblnFoco = True Then
        txtDato.BackColor = pColorD
        txtDato.FontBold = False
    End If
End Sub


'Comportamiento del control
Private Sub txtDato_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If KeyAscii = 8 Then Exit Sub 'BackScape
    If KeyAscii = 39 Then
        If InStr(1, txtDato.Text, "'") <> 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    Select Case pTipo
        Case 2 'Mayusculas
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case 3 'Entero
            If Not (IsNumeric(Chr(KeyAscii))) Then KeyAscii = 0
        Case 4 'Real
            If KeyAscii = 46 And InStr(1, txtDato.Text, ".") < 1 Then Exit Sub 'Un solo Punto Decimal
            If Not (IsNumeric(Chr(KeyAscii))) Then KeyAscii = 0
        Case 5 'Telefono
            If (KeyAscii = 32 Or KeyAscii = 40 Or KeyAscii = 41 Or KeyAscii = 42 Or KeyAscii = 45) Then Exit Sub
            If Not (IsNumeric(Chr(KeyAscii))) Then KeyAscii = 0
        Case 6 'Porcentaje
            If KeyAscii = 46 And InStr(1, txtDato.Text, ".") < 1 Then Exit Sub 'Un solo Punto Decimal
            If Not (IsNumeric(Chr(KeyAscii))) Then KeyAscii = 0
        Case 7 'Porcentaje
            If KeyAscii = 45 And InStr(1, txtDato.Text, "-") < 1 Then Exit Sub 'Un solo Punto Decimal
            If Not (IsNumeric(Chr(KeyAscii))) Then KeyAscii = 0
        Case 8
            Select Case KeyAscii
                Case "33", "34", "36", _
                     "39", "40", _
                     "41", "42", "43", "44", _
                      "46", "47", "58", _
                     "59", "60", "61", "62", _
                     "63", "64", "91", "92", _
                     "93", "94", "95", "96", _
                     "123", "124", "125", "126", _
                     "145", "146", "147", "148", _
                     "161", "166", "168", "173", _
                     "176", "178", "180", "183", _
                     "186", "191", "231"
                    KeyAscii = 0
                    Exit Sub
                Case Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End Select
     End Select
End Sub




'------------------------------------------------------------------------------
'--------------------- Propiedades del Control --------------------------------
Public Property Let Locked(ByVal vValor As Boolean)
    txtDato.Locked = vValor
End Property
Public Property Get Locked() As Boolean
    Locked = txtDato.Locked
End Property

Public Property Let ColorFoco(ByVal vColor As OLE_COLOR)
    pColorF = vColor
    txtDato.BackColor = vColor
End Property
Public Property Get ColorFoco() As OLE_COLOR
    ColorFoco = pColorF
End Property

Public Property Let ColorDefault(ByVal vColor As OLE_COLOR)
    pColorD = vColor
    txtDato.BackColor = vColor
End Property
Public Property Get ColorDefault() As OLE_COLOR
    ColorDefault = txtDato.BackColor
End Property

Public Property Let Tipo(ByVal vTipo As TipoControlTextBox)
    pTipo = vTipo
End Property
Public Property Get Tipo() As TipoControlTextBox
    Tipo = pTipo
End Property


Public Property Set Font(ByVal vFont As Font)
    Set txtDato.Font = vFont
    PropertyChanged "Font"
End Property
Public Property Get Font() As Font
    Set Font = txtDato.Font
End Property




Public Property Let Text(ByVal vText As String)
    txtDato.Text = vText
End Property
Public Property Get Text() As String
    Dim i As Integer, temp As String
    temp = txtDato.Text
    If pblnTipoSQL = False Then
        Text = temp
    Else
        i = InStr(1, txtDato.Text, "'")
        If i <> 0 Then
            Text = Mid(temp, 1, i) & "'" & Mid(temp, i + 1)
        Else
            Text = temp
        End If
    End If
End Property

Public Property Let Enabled(ByVal vModo As Boolean)
    txtDato.Enabled = vModo
    UserControl.Enabled = vModo
End Property
Public Property Get Enabled() As Boolean
    Enabled = txtDato.Enabled
End Property

Public Property Let Alignment(ByVal vModo As TipoAlignment)
    txtDato.Alignment = vModo
End Property
Public Property Get Alignment() As TipoAlignment
    Alignment = txtDato.Alignment
End Property

Public Property Let MaxLength(ByVal vModo As Integer)
    txtDato.MaxLength = vModo
End Property
Public Property Get MaxLength() As Integer
    MaxLength = txtDato.MaxLength
End Property

Public Property Let PasswordChar(ByVal vModo As String)
    txtDato.PasswordChar = vModo
End Property
Public Property Get PasswordChar() As String
    PasswordChar = txtDato.PasswordChar
End Property

Public Property Let EnabledFoco(ByVal vModo As Boolean)
    pblnFoco = vModo
End Property
Public Property Get EnabledFoco() As Boolean
    EnabledFoco = pblnFoco
End Property

Public Property Let TABAuto(ByVal vModo As Boolean)
    pblnTabAuto = vModo
End Property
Public Property Get TABAuto() As Boolean
    TABAuto = pblnTabAuto
End Property

'Colocara el doble apostrofe ('') para evitar errores
Public Property Let TipoSQL(ByVal vModo As Boolean)
    pblnTipoSQL = vModo
End Property
Public Property Get TipoSQL() As Boolean
    TipoSQL = pblnTipoSQL
End Property
Public Property Let BackColor(ByVal vBackColor As OLE_COLOR)
    txtDato.BackColor = vBackColor
End Property
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------

Public Function Clear()
    txtDato.Text = ""
End Function
Public Sub selection()
    txtDato.SelStart = 0
    txtDato.SelLength = Len(txtDato.Text)
    txtDato.SetFocus
End Sub

'---------------------------------------
'Eventos clasicos
Private Sub txtDato_Change()
    RaiseEvent Change
End Sub
Private Sub txtDato_KeyDown(KeyCode As Integer, Shift As Integer)
Dim bolFormatoOK As Boolean
Dim i As Integer
    RaiseEvent KeyDown(KeyCode, Shift)
    
    bolFormatoOK = True
    Select Case pTipo
            Case 6
                If KeyCode = vbKeyReturn Then
                    If Val(txtDato.Text) > 0 And Val(txtDato.Text) <= 100 Then
                        bolFormatoOK = True
                    Else
                        txtDato.Text = ""
                        bolFormatoOK = False
                    End If
                End If
            Case 7
                If KeyCode = vbKeyReturn Then
               i = InStr(txtDato.Text, "-")
                If i = 0 Then
                    txtDato.Text = right("000" + left(txtDato.Text, 3), 3) + "-" + Mid(txtDato.Text, 4)
                    i = 4
                End If
               If Len(txtDato.Text) = 11 And Mid(txtDato.Text, 4, 1) = "-" Then
                  If Val(left(txtDato.Text, 3)) = 0 Or Val(right(txtDato.Text, 7)) = 0 Then
                     bolFormatoOK = False
                  End If
               Else
                  txtDato.Text = right("000" + left(txtDato.Text, i - 1), 3) + "-" + right("0000000" + Mid(txtDato.Text, i + 1), 7)
               End If
               End If
                
                
    End Select
    
    If KeyCode = 13 And pblnTabAuto = True And bolFormatoOK = True Then SendKeys "{TAB}"
    
End Sub
Private Sub txtDato_Click()
    RaiseEvent Click
End Sub
'---------------------------------------

Public Sub Focus()
    txtDato.SetFocus
    txtDato.SelStart = 0
    txtDato.SelLength = Len(txtDato.Text)
End Sub





Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ColorFoco", pColorF, &HC0C0FF
    PropBag.WriteProperty "ColorDefault", pColorD, vbWhite
    PropBag.WriteProperty "ColorDefault", txtDato.BackColor, vbWhite
    PropBag.WriteProperty "Tipo", pTipo, 1
    PropBag.WriteProperty "Alignment", txtDato.Alignment, 0
    PropBag.WriteProperty "Enabled", txtDato.Enabled, True
    PropBag.WriteProperty "MaxLength", txtDato.MaxLength, 0
    PropBag.WriteProperty "PasswordChar", txtDato.PasswordChar, ""
    PropBag.WriteProperty "EnabledFoco", pblnFoco, True
    PropBag.WriteProperty "TABAuto", pblnTabAuto, True
    PropBag.WriteProperty "TipoSQL", pblnTipoSQL, False
    PropBag.WriteProperty "Bloqueado", pBloqueado, False
    PropBag.WriteProperty "Font", txtDato.Font, Ambient.Font
    PropBag.WriteProperty "Locked", txtDato.Locked, False
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    pColorF = PropBag.ReadProperty("ColorFoco", &HC0C0FF)
    pColorD = PropBag.ReadProperty("ColorDefault", vbWhite)
    txtDato.BackColor = PropBag.ReadProperty("ColorDefault", vbWhite)
    pTipo = PropBag.ReadProperty("Tipo", 1)
    txtDato.Alignment = PropBag.ReadProperty("Alignment", 0)
    txtDato.Enabled = PropBag.ReadProperty("Enabled", True)
    txtDato.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    txtDato.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    pblnFoco = PropBag.ReadProperty("EnabledFoco", True)
    pblnTabAuto = PropBag.ReadProperty("TABAuto", True)
    pblnTipoSQL = PropBag.ReadProperty("TipoSQL", False)
    pBloqueado = PropBag.ReadProperty("Bloqueado", False)
    txtDato.Locked = PropBag.ReadProperty("Locked", False)
    Set txtDato.Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub
 



Private Sub UserControl_Initialize()
    txtDato.Height = 315
    pTipo = Normal
    ColorFoco = &HC0C0FF
    ColorDefault = vbWhite
    pblnFoco = True
    pblnTabAuto = True
    pblnTipoSQL = False
End Sub

Private Sub UserControl_Resize()
    txtDato.Width = UserControl.Width
    txtDato.Height = UserControl.Height
End Sub



Public Property Let Bloqueado(mvar As Boolean)
    pBloqueado = mvar
    If mvar = True Then
        UserControl.Enabled = False
        txtDato.BackColor = &H8000000F
    Else
        UserControl.Enabled = True
        txtDato.BackColor = pColorD
    End If
End Property

Public Property Get Bloqueado() As Boolean
    Bloqueado = pBloqueado
End Property
