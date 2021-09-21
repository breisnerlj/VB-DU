Attribute VB_Name = "mdlGrilla"
Option Explicit
'******* Configuracion de Grilla**********
Sub spGrilla_Carga(ByRef rgrd As TDBGrid, _
                   ByVal vvarCaption As Variant, _
                   Optional ByVal vvarWidth As Variant, _
                   Optional ByVal vvarAlignment As Variant, _
                   Optional ByVal vvarDataSource As Variant, _
                   Optional ByVal vblnAllowSizing As Boolean)
Dim Columna As TrueDBGrid70.Column
Dim i%
    rgrd.Columns.Clear
    For i = 0 To UBound(vvarCaption)
        Set Columna = rgrd.Columns.Add(i)
        If Not IsMissing(vvarWidth) Then rgrd.Columns(i).Width = vvarWidth(i)
        If Not IsMissing(vvarAlignment) Then rgrd.Columns(i).Alignment = vvarAlignment(i)
        If Not IsMissing(vvarDataSource) Then rgrd.Columns(i).DataField = vvarDataSource(i)
        rgrd.Columns(i).AllowSizing = vblnAllowSizing
        rgrd.Columns(i).Visible = True
        rgrd.Columns(i).Caption = vvarCaption(i)
        rgrd.Columns(i).WrapText = True
        DoEvents
    Next i
    rgrd.AllowAddNew = False
    rgrd.Splits(0).AllowColSelect = False
    rgrd.Splits(0).AllowRowSelect = True
    rgrd.Splits(0).AllowRowSizing = False
    rgrd.Splits(0).AllowSizing = False
    rgrd.Splits(0).Style.VerticalAlignment = dbgVertCenter
    rgrd.AllowUpdate = False
    rgrd.HoldFields
'    rgrd.Styles(5).BackColor = &H80FF80
'    rgrd.Styles(5).Font.Bold = True
'    rgrd.Styles(5).ForeColor = vbBlack
    Set Columna = Nothing
End Sub

Sub spGrilla_Titulos(ByRef rgrd As TDBGrid, ByVal vvarTitulo As Variant)
Dim i%
    For i = 0 To UBound(vvarTitulo)
        rgrd.Columns(i).Caption = vvarTitulo(i)
        DoEvents
    Next i
End Sub

Sub spGrilla_Alinea(ByRef rgrd As TDBGrid, ByVal vvarAlinea As Variant)
Dim i%
    For i = 0 To UBound(vvarAlinea)
        rgrd.Columns(i).Alignment = vvarAlinea(i)
        DoEvents
    Next i
End Sub

Sub spGrilla_Ancho(ByRef rgrd As TDBGrid, ByVal vvarAncho As Variant)
Dim i%
    For i = 0 To UBound(vvarAncho)
        rgrd.Columns(i).Width = vvarAncho(i)
        DoEvents
    Next i
End Sub

Sub spGrilla_DatoCampo(ByRef rgrd As TDBGrid, ByVal vvarDatoCampo As Variant)
Dim i%
    For i = 0 To rgrd.Columns.Count - 1
        If i > UBound(vvarDatoCampo) Then rgrd.Columns(i).Visible = False Else rgrd.Columns(i).DataField = vvarDatoCampo(i)
        DoEvents
    Next i
End Sub

Sub spGrilla_CheckBox(ByRef rgrd As Object, ByVal vvarColumn As Variant)
    rgrd.Columns(vvarColumn).ValueItems.Presentation = dbgCheckBox
End Sub

Sub spGrilla_CboBox(ByRef rgrd As Object, _
                    ByVal vvarColumn As Variant, _
                    ByVal vvarValue As Variant, _
                    ByVal vrst As Object, _
                    ByVal vvarDisplay As Variant)
                    
Dim ValueItem As New TrueDBGrid70.ValueItem
Dim i%
    rgrd.Columns(vvarColumn).ValueItems.Clear
    vrst.MoveFirst
    While Not vrst.EOF
        ValueItem.Value = vrst(vvarValue).Value
        If Not IsMissing(vvarDisplay) Then
           ValueItem.DisplayValue = vrst(vvarDisplay).Value
        End If
        rgrd.Columns(vvarColumn).ValueItems.Add ValueItem
        vrst.MoveNext
    Wend
    If vrst.RecordCount = 0 Then Exit Sub
    rgrd.Columns(vvarColumn).ValueItems.Presentation = dbgComboBox
    rgrd.Columns(vvarColumn).ValueItems.MaxComboItems = vrst.RecordCount
    rgrd.Columns(vvarColumn).DropDownList = True
    
    If Not IsMissing(vvarDisplay) Then
        rgrd.Columns(vvarColumn).ValueItems.Translate = True
    End If
    
    Set ValueItem = Nothing
End Sub

Sub spGrilla_Traslate(ByRef rgrd As Object, ByVal vvarColumn As Variant, ByVal vvarValue As Variant, ByVal vvarDisplayValue As Variant)
Dim ValueItem As New TrueDBGrid70.ValueItem
On Error GoTo Error
    rgrd.Columns(vvarColumn).ValueItems.Translate = True
    ValueItem.DisplayValue = vvarDisplayValue
    ValueItem.Value = vvarValue
    rgrd.Columns(vvarColumn).ValueItems.Add ValueItem
    Set ValueItem = Nothing
    GoTo OK
Error:
    MsgBox Err.Description, vbExclamation, "Error"
OK:
    On Error GoTo 0
End Sub

Sub spGrilla_Locked_Columns(ByRef rgrd As TDBGrid, Optional ByVal vintColumna%)
Dim i%
    rgrd.AllowUpdate = False
    With rgrd
        For i = 0 To .Columns.Count - 1
            .Columns(i).Locked = True
            .Columns(i).AllowFocus = False
            If Not IsMissing(vintColumna) Then If (i = vintColumna) Then .Columns(i).Locked = False
            DoEvents
        Next i
        DoEvents
    End With
End Sub

Sub spGrilla_UnLocked_Column(ByRef rgrd As TDBGrid, ByVal vvarColumna As Variant)
    DoEvents
    rgrd.AllowUpdate = True
    With rgrd
            .Columns(vvarColumna).Locked = False
            .Columns(vvarColumna).AllowFocus = True
    End With
End Sub
'*************************************************************************************


Function pfstr_Valida_Fecha(ByRef rstrFecha$, _
                           Optional ByVal vblnVacio As Boolean = False, _
                           Optional ByVal vstrFormato$ = "DD/MM/YYYY", _
                           Optional ByVal vstrMaxFecha$ = "", _
                           Optional ByVal vstrMinFecha$ = "", _
                           Optional ByVal blnMenorHoy As Boolean = False) As String
    
    pfstr_Valida_Fecha = ""
    On Error GoTo Error
    
    Dim j%
    
    If vblnVacio And Trim(rstrFecha) = "" Then Exit Function
    
    rstrFecha = CDate(rstrFecha)
    GoTo OK:
Error:
    pfstr_Valida_Fecha = "Fecha No Valida"
OK:
        
    '-- Parchesito para que valide la cantidad de caracteres de la fecha --'
    j = 0
    For j = 0 To Len(rstrFecha) - 1
               
    Next j
    
    If j <> Len(Format(objUsuario.sysdate, "dd/mm/yyyy")) Then
        pfstr_Valida_Fecha = "Valor de Fecha No Válida"
        Exit Function
    End If
    '-----------------------------------------------------------------------'
        
        
    If Not IsDate(Format(rstrFecha, vstrFormato)) Then
         pfstr_Valida_Fecha = "Valor de Fecha No Válida"
         Exit Function
        
    End If
    
    If Trim(vstrMaxFecha) <> "" Then
        If CDate(Format(rstrFecha, vstrFormato)) > CDate(Format(vstrMaxFecha, vstrFormato)) Then
            pfstr_Valida_Fecha = "Fecha Máxima << " & Format(vstrMaxFecha, vstrFormato) & " >>"
            Exit Function
        End If
    End If
    If Trim(vstrMinFecha) <> "" Then
        If CDate(Format(rstrFecha, vstrFormato)) < CDate(Format(vstrMinFecha, vstrFormato)) Then
            pfstr_Valida_Fecha = "Fecha Mínima << " & Format(vstrMinFecha, vstrFormato) & " >>"
            Exit Function
        End If
    End If
    
    If blnMenorHoy Then
        If CDate(Format(rstrFecha, vstrFormato)) < CDate(Format(objUsuario.sysdate, vstrFormato)) Then
            pfstr_Valida_Fecha = "Fecha No puede ser menor a la de Hoy"
            Exit Function
        End If
    End If
End Function

