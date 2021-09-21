VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_OFF_Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim strTipo As String
Dim strRuc As String
Dim strDireccion As String
Dim strNombre As String

Public Property Get Tipo() As String
    Tipo = strTipo

End Property

Public Property Let Tipo(ByVal lstrTipo As String)
strTipo = lstrTipo

End Property

Public Property Get Ruc() As String
    Ruc = strRuc
End Property

Public Property Let Ruc(ByVal lstrRuc As String)
    strRuc = lstrRuc
End Property



Public Property Get Direccion() As String
    Direccion = strDireccion
End Property

Public Property Let Direccion(ByVal lstrDireccion As String)
    strDireccion = lstrDireccion
End Property

Public Property Get Nombre() As String
    Nombre = strNombre
End Property

Public Property Let Nombre(ByVal lstrNombre As String)
    strNombre = lstrNombre
End Property


Public Function Limpia()
    strTipo = "0"
    strRuc = ""
    strDireccion = ""
    strNombre = ""
End Function

Public Function ValidaRuc(ByVal strRuc As String) As Integer
Dim intDig1 As Integer
Dim intDig2 As Integer
Dim intDig3 As Integer
Dim intDig4 As Integer
Dim intDig5 As Integer
Dim intDig6 As Integer
Dim intDig7 As Integer
Dim intDig8 As Integer
Dim intDig9 As Integer
Dim intDig0 As Integer
Dim intDigX As Integer
'---------------------
Dim intValorA As Integer
Dim intValorB As Integer
Dim intValorC As Integer
Dim intValorD As Integer
Dim intDigV As Integer


On Error GoTo CtrlErr
    '------ paso 0
    intDig1 = Val(Mid(strRuc, 1, 1))
    intDig2 = Val(Mid(strRuc, 2, 1))
    intDig3 = Val(Mid(strRuc, 3, 1))
    intDig4 = Val(Mid(strRuc, 4, 1))
    intDig5 = Val(Mid(strRuc, 5, 1))
    intDig6 = Val(Mid(strRuc, 6, 1))
    intDig7 = Val(Mid(strRuc, 7, 1))
    intDig8 = Val(Mid(strRuc, 8, 1))
    intDig9 = Val(Mid(strRuc, 9, 1))
    intDig0 = Val(Mid(strRuc, 10, 1))
    intDigX = Val(Mid(strRuc, 11, 1))
    '------ paso 1
    intDig1 = intDig1 * 5
    intDig2 = intDig2 * 4
    intDig3 = intDig3 * 3
    intDig4 = intDig4 * 2
    intDig5 = intDig5 * 7
    intDig6 = intDig6 * 6
    intDig7 = intDig7 * 5
    intDig8 = intDig8 * 4
    intDig9 = intDig9 * 3
    intDig0 = intDig0 * 2
    '------ paso 2
    intValorA = intDig1 + intDig2 + intDig3 + intDig4 + intDig5 + intDig6
    intValorA = intValorA + intDig7 + intDig8 + intDig9 + intDig0
    '------ paso 3
    intValorB = Int(intValorA / 11)
    '------ paso 4
    intValorC = intValorA - (intValorB * 11)
    '------ paso 5
    intValorD = 11 - intValorC
    '------ paso 6
    If intValorD = 10 Then
        intDigV = 0
    ElseIf intValorD = 11 Then
        intDigV = 1
    Else
        intDigV = intValorD
    End If
    '@@@@@@@@@@@@@@@@@@@@@@@@@@
    If Trim(strRuc) = "10000000001" Or Trim(strRuc) = "10000000002" Then
        ValidaRuc = 0
    End If
    If intDigX = intDigV Then
        If Mid(strRuc, 1, 4) = "9999" Then
            ValidaRuc = 1
        Else
            ValidaRuc = 0
        End If
    Else
        ValidaRuc = -1
    End If
    
Exit Function

CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Cliente.ValidaRuc", Err.Description


End Function

