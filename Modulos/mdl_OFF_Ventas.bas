Attribute VB_Name = "mdl_OFF_Ventas"
Option Explicit




'Autor : Juan Arturo Escate Espichan
'Fecha : 08/05/2008
'Proposito : Pone varios arrays en un Xarray
'---------------------------------------------------------------------------------------------------------
Function ArrayaXarray(ByVal Array1 As Variant, _
                      Optional ByVal Array2 As Variant = "", _
                      Optional ByVal Array3 As Variant = "", _
                      Optional ByVal Array4 As Variant = "", _
                      Optional ByVal Array5 As Variant = "", _
                      Optional ByVal Array6 As Variant = "", _
                      Optional ByVal Array7 As Variant = "", _
                      Optional ByVal Array8 As Variant = "", _
                      Optional ByVal Array9 As Variant = "", _
                      Optional ByVal Array10 As Variant = "", _
                      Optional ByVal Array11 As Variant = "", _
                      Optional ByVal Array12 As Variant = "", _
                      Optional ByVal Array13 As Variant = "", _
                      Optional ByVal Array14 As Variant = "", _
                      Optional ByVal Array15 As Variant = "" _
                      ) As XArrayDB
                      
Dim columnas As Integer
Dim i As Integer
Dim xTmp As New XArrayDB
'Verifico cual el numero de columnas
If IsArray(Array1) Then columnas = 1
If IsArray(Array2) Then columnas = 2
If IsArray(Array3) Then columnas = 3
If IsArray(Array4) Then columnas = 4
If IsArray(Array5) Then columnas = 5
If IsArray(Array6) Then columnas = 6
If IsArray(Array7) Then columnas = 7
If IsArray(Array8) Then columnas = 8
If IsArray(Array9) Then columnas = 9
If IsArray(Array10) Then columnas = 10
If IsArray(Array11) Then columnas = 11
If IsArray(Array12) Then columnas = 12
If IsArray(Array13) Then columnas = 13
If IsArray(Array14) Then columnas = 14
If IsArray(Array15) Then columnas = 15
Dim y As Integer
i = 0
   xTmp.ReDim 0, -1, 0, columnas - 1
    While i <= UBound(Array1)
        y = xTmp.Count(1)
        xTmp.AppendRows
        If IsArray(Array1) Then xTmp(y, 0) = Array1(i)
        If IsArray(Array2) Then xTmp(y, 1) = Array2(i)
        If IsArray(Array3) Then xTmp(y, 2) = Array3(i)
        If IsArray(Array4) Then xTmp(y, 3) = Array4(i)
        If IsArray(Array5) Then xTmp(y, 4) = Array5(i)
        If IsArray(Array6) Then xTmp(y, 5) = Array6(i)
        If IsArray(Array7) Then xTmp(y, 6) = Array7(i)
        If IsArray(Array8) Then xTmp(y, 7) = Array8(i)
        If IsArray(Array9) Then xTmp(y, 8) = Array9(i)
        If IsArray(Array10) Then xTmp(y, 9) = Array10(i)
        If IsArray(Array11) Then xTmp(y, 10) = Array11(i)
        If IsArray(Array12) Then xTmp(y, 11) = Array12(i)
        If IsArray(Array13) Then xTmp(y, 12) = Array13(i)
        If IsArray(Array14) Then xTmp(y, 13) = Array14(i)
        If IsArray(Array15) Then xTmp(y, 14) = Array15(i)
        
        i = i + 1
    Wend
Set ArrayaXarray = xTmp
End Function

Public Function Letra(ByVal strnum As String, Optional vLo) As String
Dim lngA As Long
Dim Negativo As Boolean
Dim ln As Long
Dim Una As Boolean
Dim Millon As Boolean
Dim Millones As Boolean
Dim vez As Long
Dim MaxVez As Long
Dim strQ As String
Dim strB As String
Dim strU As String
Dim strD As String
Dim strC As String
Dim iA As Long
Dim k As Long
    
Dim strN() As String
Dim lo As Long

'Si no se especifica el ancho...
If IsMissing(vLo) Then
  lo = 255
Else
  lo = vLo
End If
        
Dim unidad(0 To 9) As String
Dim decena(0 To 9) As String
Dim centena(0 To 9) As String
Dim deci(0 To 9) As String
Dim otros(0 To 15) As String
'Asignar los valores
unidad(1) = "uno"          ' Dejar en minusculas
unidad(2) = "dos"
unidad(3) = "tres"
unidad(4) = "cuatro"
unidad(5) = "cinco"
unidad(6) = "seis"
unidad(7) = "siete"
unidad(8) = "ocho"
unidad(9) = "nueve"

decena(1) = "diez"
decena(2) = "veinte"
decena(3) = "treinta"
decena(4) = "cuarenta"
decena(5) = "cincuenta"
decena(6) = "sesenta"
decena(7) = "setenta"
decena(8) = "ochenta"
decena(9) = "noventa"
   
centena(1) = "ciento"
centena(2) = "doscientos"
centena(3) = "trescientos"
centena(4) = "cuatrocientos"
centena(5) = "quinientos"
centena(6) = "seiscientos"
centena(7) = "setecientos"
centena(8) = "ochocientos"
centena(9) = "novecientos"
    
deci(1) = "dieci"
deci(2) = "veinti"
deci(3) = "treinti"
deci(4) = "cuarenti"
deci(5) = "cincuenti"
deci(6) = "sesenti"
deci(7) = "setenti"
deci(8) = "ochenti"
deci(9) = "noventi"
    
otros(1) = "1"
otros(2) = "2"
otros(3) = "3"
otros(4) = "4"
otros(5) = "5"
otros(6) = "6"
otros(7) = "7"
otros(8) = "8"
otros(9) = "9"
otros(10) = "10"
otros(11) = "once"
otros(12) = "doce"
otros(13) = "trece"
otros(14) = "catorce"
otros(15) = "quince"

On Error GoTo 0
    
lngA = Abs(Val(strnum))
Negativo = (lngA <> Val(strnum))
strnum = CStr(lngA)
ln = Len(strnum)
    
If lngA = 0 Then
  strnum = left$("cero" & Space$(lo), lo)
  Exit Function
End If
    '
Una = True
Millon = False
Millones = False
If ln < 4 Then Una = False
  If lngA > 999999 Then Millon = True
  If lngA > 1999999 Then Millones = True
  strB = ""
  strQ = strnum
  vez = 0
    
  ReDim strN(1 To 4)
  strQ = right$(String$(12, "0") & strnum, 12)
  For k = Len(strQ) To 1 Step -3
    vez = vez + 1
    strN(vez) = Mid$(strQ, k - 2, 3)
  Next
  MaxVez = 4
  For k = 4 To 1 Step -1
    If strN(k) = "000" Then
      MaxVez = MaxVez - 1
    Else
      Exit For
    End If
  Next
  For vez = 1 To MaxVez
    strU = "": strD = "": strC = ""
    strnum = strN(vez)
    ln = Len(strnum)
    k = Val(right$(strnum, 2))
    If right$(strnum, 1) = "0" Then
      k = k \ 10
      strD = decena(k)
      ElseIf k > 10 And k < 16 Then
        k = Val(Mid$(strnum, ln - 1, 2))
        strD = otros(k)
      Else
        strU = unidad(Val(right$(strnum, 1)))
        If ln - 1 > 0 Then
          k = Val(Mid$(strnum, ln - 1, 1))
          strD = deci(k)
        End If
      End If
      If ln - 2 > 0 Then
        k = Val(Mid$(strnum, ln - 2, 1))
        strC = centena(k) & " "
      End If
      If strU = "uno" And left$(strB, 4) = " mil" Then strU = ""
        strB = strC & strD & strU & " " & strB
        ''If (vez = 1 Or vez = 3) And strN(vez + 1) <> "000" Then strB = " mil " & strB
        '-------------------------------------
      If (vez = 1 Or vez = 3) And strN(vez + 1) <> "000" And Len(CStr(lngA)) = 4 And Val(Mid(CStr(lngA), 1, 1)) = 1 Then
        strB = " un mil " & strB
      Else
        If (vez = 1 Or vez = 3) And strN(vez + 1) <> "000" Then strB = " mil " & strB
        End If
      '-------------------------------------
      If vez = 2 And Millon Then
        If Millones Then
          strB = " millones " & strB
        Else
          strB = "un millón " & strB
        End If
      End If
  Next
  strB = CStr(strB)
  If right$(strB, 3) = "uno" Then strB = left$(strB, Len(strB) - 1) & "o"
  Do                              'Quitar los espacios que haya por medio
   iA = InStr(strB, "  ")
   If iA = 0 Then Exit Do
   strB = left$(strB, iA - 1) & Mid$(strB, iA + 1)
  Loop
  If left$(strB, 6) = "uno un" Then strB = Mid$(strB, 5)
  If left$(strB, 7) = "uno mil" Then strB = Mid$(strB, 5)
  If right$(strB, 16) <> "millones mil uno" Then
    iA = InStr(strB, "millones mil uno")
    If iA Then strB = left$(strB, iA + 8) & Mid$(strB, iA + 13)
  End If
  If right$(strB, 6) = "ciento" Then strB = left$(strB, Len(strB) - 2)
  If Negativo Then strB = "menos " & strB
  
  strC = Space$(lo)
  LSet strC = strB
  Letra = strC
End Function

Public Function CadenaDirecc_Part1(ByVal vstrCadDirecc As String)
    Dim i As Integer
    Dim C1 As String
    C1 = ""
    For i = 1 To Len(vstrCadDirecc)
       If i <= 31 Then
           C1 = C1 + Mid(vstrCadDirecc, i, 1)
       End If
    Next i
    CadenaDirecc_Part1 = C1
End Function

Public Function CadenaDirecc_Part2(ByVal vstrCadDirecc As String)
    Dim i As Integer
    Dim C2 As String
    C2 = ""
    For i = 1 To Len(vstrCadDirecc)
       If i > 31 Then
           C2 = C2 + Mid(vstrCadDirecc, i, 1)
       End If
    Next i
    CadenaDirecc_Part2 = C2
End Function


Public Function fmod(ByVal strNumTarjeta As String) As Integer
Dim j As Integer
Dim i As Integer
Dim X As Integer
Dim k As Long
Dim lngResult As Long
Dim lngTotal As Long

X = 1
If InStr(strNumTarjeta, "00000000") = 0 Then
    For i = Len(strNumTarjeta) To 1 Step -1
        lngResult = Val(Mid(strNumTarjeta, i, 1)) * X
        If lngResult >= 10 Then
            lngTotal = lngTotal + Val(Mid(Trim(Str(lngResult)), 1, 1)) + Val(Mid(Trim(Str(lngResult)), 2, 1))
        Else
            lngTotal = lngTotal + lngResult
        End If
        If X = 2 Then
            X = 1
        Else
            X = 2
        End If
    Next i
    
    k = lngTotal Mod 10
    
    If k = 0 Then
        j = 1
    Else
        j = 0
    End If
    
    
Else
    j = 1
End If



fmod = j

End Function



