VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim loses As OraSession
Dim lodb As OraDatabase
Dim strIntento As Integer
Enum TipoParam
    entrada = ORAPARM_INPUT
    Salida = ORAPARM_OUTPUT
    Entrada_Salida = ORAPARM_BOTH
End Enum
Property Let Num_Intentos(Intento As Integer)
    strIntento = Intento
End Property
Property Get OSession() As OraSession
    Set OSession = loses
End Property

Property Get ODataBase() As OraDatabase
    Set ODataBase = lodb
End Property

Function Conexion(ByVal vstrTNSNAME$, ByVal vstrUsuario$, ByVal vstrPASSWORD$) As String
    Set loses = Nothing
    Set lodb = Nothing
    On Error GoTo Error
    Set loses = CreateObject("OracleInProcServer.XOraSession")
    Set lodb = loses.OpenDatabase(vstrTNSNAME, vstrUsuario & "/" & vstrPASSWORD, 0&)
    'lodb.AutoCommit = False
    Conexion = ""
    Exit Function
Error:
    MsgBox "Error en la Conexión a la BD" & Chr(13) & Err.Description, vbCritical, App.ProductName
    Conexion = Err.Description
End Function

Sub Cerrar()

'    Dim i As Integer
'    For i = lodb.Parameters.Count - 1 To 0 Step -1
'        lodb.Parameters.Remove i
'    Next
    LimpiaParametros
    lodb.Close
    Set lodb = Nothing
    Set loses = Nothing
End Sub

Sub ParamAdd(ByVal vvarParametros As Variant, _
             ByVal vvarValores As Variant)
Dim i%
    For i = 0 To UBound(vvarParametros)
        lodb.Parameters.Add vvarParametros(i), _
                            vvarValores(i), _
                            ORAPARM_INPUT, _
                            ORATYPE_CHAR
    Next i
End Sub

Sub ParamRemove(ByVal vvarParametros As Variant)
Dim i%
    For i = 0 To UBound(vvarParametros)
        lodb.Parameters.Remove vvarParametros(i)
    Next i
End Sub

Sub Execute(ByVal vstrSqlStmt$)
    lodb.ExecuteSQL vstrSqlStmt
End Sub

'''Function Query(ByVal vstrSql As String, _
'''                    Optional ByVal vvarParametros As Variant, _
'''                    Optional ByVal vvarValores As Variant, _
'''                    Optional ByVal vlngFetchLimit As Long = 0, _
'''                    Optional ByRef roradc As Object) As OraDynaset
'''
'''    Dim intIntentos%, odynSql As OraDynaset
'''    If Not IsMissing(vvarParametros) Then Call ParamAdd(vvarParametros, vvarValores)
'''
'''    intIntentos = 1
'''    Screen.MousePointer = vbHourglass
'''    On Local Error GoTo Error
'''
'''    If Not roradc Is Nothing Then
'''        If vlngFetchLimit = 0 Then
'''            Set roradc.Recordset = lodb.CreateDynaset(vstrSql, 0&)
'''        Else
'''            Set roradc.Recordset = lodb.CreateCustomDynaset(vstrSql, 0&, 256, 16, 20, vlngFetchLimit, 4096)
'''        End If
'''        Set odynSql = roradc.Recordset
'''    Else
'''        If vlngFetchLimit = 0 Then
'''            Set odynSql = lodb.CreateDynaset(vstrSql, 0&)
'''        Else
'''            Set odynSql = lodb.CreateCustomDynaset(vstrSql, 0&, 256, 16, 20, vlngFetchLimit, 4096)
'''        End If
'''    End If
'''    GoTo OK
'''Error:
'''
'''    If intIntentos = 1 Then
'''        intIntentos = intIntentos + 1
'''        Resume
'''    ElseIf intIntentos >= 3 Then
'''        MsgBox "COMUNIQUESE CON SISTEMAS" & Chr(13) & Chr(13) & Chr(13) & Err.Description, vbCritical, "Error Crítico en el Intento " & (intIntentos)
'''    Else
'''        MsgBox "PRESIONE ACEPTAR PARA INTENTAR NUEVAMENTE" & Chr(13) & Chr(13) & Chr(13) & Err.Description, vbExclamation, "Ejecutando Consulta en el intento " & intIntentos
'''    End If
'''    intIntentos = intIntentos + 1
'''    If intIntentos <= 5 Then Resume Else GoTo OK
'''OK:
'''    If Not IsMissing(vvarValores) Then Call ParamRemove(vvarParametros)
'''    Set Query = odynSql
'''    Screen.MousePointer = vbDefault
'''End Function
'''''
'/****** PARA EJECUTAR LOS STORE PROCEDURE DEL ORACLE ******/
Function SP(ByVal vstrSP_Nombre$, _
            ByRef rvarValores As Variant, _
            ByVal vvarInOut_Put As Variant) As String
Dim i%, strExec$, varParam() As Variant
Dim X As OraMDAttribute
    
    Screen.MousePointer = vbHourglass
    strExec = ""
    Dim strPar As String 'ECASTILLO 21.04.2021
    strPar = ""
    LimpiaParametros
           
    For i = 0 To UBound(rvarValores)
        ReDim Preserve varParam(i)
        varParam(i) = "P_" & i
        lodb.Parameters.Remove varParam(i)
    Next i
    
    'Genera la Cadena para Ejecutar el SP
    strExec = " BEGIN " & vstrSP_Nombre & " ( "
    If Not IsMissing(varParam) Then
        For i = 0 To UBound(varParam)
            lodb.Parameters.Add varParam(i), IIf(vvarInOut_Put(i) = Salida, Space(1500), Trim(rvarValores(i))), vvarInOut_Put(i), ORATYPE_VARCHAR2
            strPar = strPar & varParam(i) & " => " & rvarValores(i) & IIf(i = UBound(rvarValores), "", ", ")
            strExec = strExec & ":" & varParam(i) & IIf(i = UBound(varParam), "", ", ")
        Next i
    End If
    strExec = strExec & "); END; "
    
    On Error GoTo Error
    
    lodb.ExecuteSQL strExec
    On Error GoTo 0
    
    For i = 0 To UBound(varParam)
        If (vvarInOut_Put(i) = ORAPARM_BOTH) Or (vvarInOut_Put(i) = ORAPARM_OUTPUT) Then
            rvarValores(i) = Trim("" & lodb.Parameters(varParam(i)).Value)
        End If
    Next i
    'I.ECASTILLO 21.04.2021
    createLog "log", "dlv_unificado", strExec, strPar
    'F.ECASTILLO 21.04.2021
    GoTo Salida
Error:
    'I.ECASTILLO 21.04.2021
    createLog "log", "dlv_unificado", strExec, strPar, Err.Description, "1"
    'F.ECASTILLO 21.04.2021
    SP = Err.Description
    On Error GoTo 0
Salida:
    ParamRemove varParam
    Screen.MousePointer = vbDefault
End Function

'/**********************************************************************************/

Function FN_Cursor(ByVal vstrFN_Nombre As String, _
                   ByVal vlngFetchLimit As Long, _
                   ParamArray vvarValores() As Variant _
                   ) As OracleInProcServer.oraDynaset
    
    Dim i%, strExec$, intIntentos%, varParam() As Variant

On Error GoTo Error
    Dim strPar As String 'ECASTILLO 21.04.2021
    strPar = ""
    
    For i = 0 To UBound(vvarValores)
        ReDim Preserve varParam(i)
        varParam(i) = "A_" & i
        lodb.Parameters.Remove varParam(i)
    Next i
    
    
    For i = 0 To lodb.Parameters.Count - 1
        lodb.Parameters.Remove i
    Next
    
    'Screen.MousePointer = vbHourglass
    strExec = " BEGIN :SALIDA := " & vstrFN_Nombre & " ( "
        
    If Not IsMissing(vvarValores) Then
        For i = 0 To UBound(vvarValores)
            lodb.Parameters.Add varParam(i), vvarValores(i), ORAPARM_INPUT, ORATYPE_VARCHAR2
            strPar = strPar & varParam(i) & " => " & vvarValores(i) & IIf(i = UBound(vvarValores), "", ", ")
            strExec = strExec & ":" & varParam(i) & IIf(i = UBound(vvarValores), "", ", ")
        Next i
    End If
    
    strExec = strExec & "); END;"
    
    'On Error GoTo Error
    If val(vlngFetchLimit) = 0 Then
        Set FN_Cursor = lodb.CreatePlsqlDynaset(strExec, "SALIDA", ORADYN_READONLY)
    Else
        '''''''''MsgBox "QUEEEEE????"
        Set FN_Cursor = lodb.CreatePlsqlCustomDynaset(strExec, "SALIDA", ORADYN_READONLY, 256, 16, 20, vlngFetchLimit, 4096)
        '''''Set FN_Cursor = lodb.CreatePlsqlCustomDynaset(strExec, "SALIDA", ORADYN_NOCACHE, 256, 16, 20, vlngFetchLimit, 4096)
    End If
    'GoTo Salida
    'I.ECASTILLO 21.04.2021
    createLog "log", "dlv_unificado", strExec, strPar
    'F.ECASTILLO 21.04.2021
    Exit Function
Error:
    
    'I.ECASTILLO 21.04.2021
    createLog "log", "dlv_unificado", strExec, strPar, Err.Description, "1"
    'F.ECASTILLO 21.04.2021
    Err.Raise vbObjectError + lodb.LastServerErr, "clsOracle", Err.Description
'    If Not strIntento = 0 Then
'        If strIntento = intIntentos Then GoTo Salida
'    End If
'    If intIntentos = 1 Then
'        intIntentos = intIntentos + 1
'        Resume
'    ElseIf intIntentos >= 3 Then
'        MsgBox "COMUNIQUESE CON SISTEMAS" & Chr(13) & Chr(13) & Chr(13) & Err.Description, vbCritical, "Error Crítico en el Intento " & (intIntentos)
'    Else
'        MsgBox "PRESIONE ACEPTAR PARA INTENTAR NUEVAMENTE" & Chr(13) & Chr(13) & Chr(13) & Err.Description, vbExclamation, "Ejecutando Consulta en el intento " & intIntentos
'    End If
'    intIntentos = intIntentos + 1
'    If intIntentos <= 5 Then Resume Else GoTo Salida
'Salida:
'On Error GoTo 0
'    'If Not IsMissing(vvarValores) Then Call ParamRemove(varParam)
'    Screen.MousePointer = vbDefault
End Function
'/**********************************************************************************/
Function FN_ORADC(ByVal vstrFN_Nombre As String, _
                  ByRef roradc As ORADC, _
                  Optional ByVal vlngFetchLimit As Long = 0, _
                  Optional ByVal vvarValores As Variant) As oraDynaset
    
    Dim i%, strExec$, intIntentos%, varParam() As Variant
    
    If Not IsMissing(vvarValores) Then
        For i = 0 To UBound(vvarValores)
            ReDim Preserve varParam(i)
            varParam(i) = "A_" & i
            lodb.Parameters.Remove varParam(i)
        Next i
    End If
    Screen.MousePointer = vbHourglass
    strExec = " BEGIN :SALIDA := " & vstrFN_Nombre
        
    If Not IsMissing(vvarValores) Then
        strExec = strExec & " ( "
        For i = 0 To UBound(vvarValores)
            lodb.Parameters.Add varParam(i), vvarValores(i), ORAPARM_INPUT, ORATYPE_VARCHAR2
            strExec = strExec & ":" & varParam(i) & IIf(i = UBound(vvarValores), "", ", ")
        Next i
        strExec = strExec & ") "
    End If
    
    strExec = strExec & "; END;"
    
    On Error GoTo Error
    If vlngFetchLimit = 0 Then
        Set roradc.Recordset = lodb.CreatePlsqlDynaset(strExec, "SALIDA", 0&)
    Else
        Set roradc.Recordset = lodb.CreatePlsqlCustomDynaset(strExec, "SALIDA", 0&, 256, 16, 20, vlngFetchLimit, 4096)
    End If

    Set FN_ORADC = roradc.Recordset
    GoTo Salida
Error:

    If intIntentos = 1 Then
        intIntentos = intIntentos + 1
        Resume
    ElseIf intIntentos >= 3 Then
        MsgBox "COMUNIQUESE CON SISTEMAS" & Chr(13) & Chr(13) & Chr(13) & Err.Description, vbCritical, "Error Crítico en el Intento " & (intIntentos)
    Else
        MsgBox "PRESIONE ACEPTAR PARA INTENTAR NUEVAMENTE" & Chr(13) & Chr(13) & Chr(13) & Err.Description, vbExclamation, "Ejecutando Consulta en el intento " & intIntentos
    End If
    intIntentos = intIntentos + 1
    If intIntentos <= 5 Then Resume Else GoTo Salida
Salida:
On Error GoTo 0
    If Not IsMissing(vvarValores) Then ParamRemove varParam
    Screen.MousePointer = vbDefault
End Function
'/**********************************************************************************/
Function FN_Valor(ByVal vstrFN_Nombre As String, _
                  ParamArray vvarValores()) As Variant
    
    Dim i%, strExec$, odynSql As oraDynaset, intIntentos%, varParam() As Variant
    
    Screen.MousePointer = vbHourglass
    
    lodb.Parameters.Remove "RET_SALIDA_SQL"
    For i = 0 To UBound(vvarValores)
        ReDim Preserve varParam(i)
        varParam(i) = "A_" & i
        lodb.Parameters.Remove varParam(i)
    Next i
    
    lodb.Parameters.Add "RET_SALIDA_SQL", Space(1000), ORAPARM_OUTPUT, ORATYPE_VARCHAR2
    strExec = " BEGIN :RET_SALIDA_SQL := " & vstrFN_Nombre
    If UBound(vvarValores) >= 0 Then
        strExec = strExec & " ( "
        For i = 0 To UBound(vvarValores)
            lodb.Parameters.Add varParam(i), vvarValores(i), ORAPARM_INPUT, ORATYPE_VARCHAR2
            strExec = strExec & ":" & varParam(i) & IIf(i = UBound(varParam), "", ", ")
        Next i
        strExec = strExec & ") "
    End If
    strExec = strExec & "; END;"
    
    On Error GoTo Error
            
    lodb.ExecuteSQL strExec
'comentado 13/08/09 pherrera
'Este reintento no es necesario, entraba en un bucle innecesario. si se cae, se debe mostrar el mensaje de error!!!!

'    GoTo Ok
'Error:
'    If Not strIntento = 0 Then
'        If strIntento = intIntentos Then GoTo Salida
'    End If
'    If intIntentos = 1 Then
'        intIntentos = intIntentos + 1
'        Resume
'    ElseIf intIntentos >= 3 Then
'        MsgBox "COMUNIQUESE CON SISTEMAS" & Chr(13) & Chr(13) & Chr(13) & Err.Description, vbCritical, "Error Crítico en el Intento " & (intIntentos)
'    Else
'        MsgBox "PRESIONE ACEPTAR PARA INTENTAR NUEVAMENTE" & Chr(13) & Chr(13) & Chr(13) & Err.Description, vbExclamation, "Ejecutando Consulta en el intento " & intIntentos
'        'Lo cambie 31/10/2006'
'        intIntentos = 1
'    End If
'    If strIntento <= 5 Then Resume Else GoTo Salida
'Ok:
    FN_Valor = lodb.Parameters("RET_SALIDA_SQL").Value
'Salida:
'On Error GoTo 0
'    If Not IsMissing(vvarValores) Then Call ParamRemove(varParam)
    lodb.Parameters.Remove "RET_SALIDA_SQL"
    Screen.MousePointer = vbDefault
    Exit Function
Error:
   If Not IsMissing(vvarValores) Then Call ParamRemove(varParam)
    lodb.Parameters.Remove "RET_SALIDA_SQL"
    Screen.MousePointer = vbDefault
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
  
'/**********************************************************************************/
Function Const_Val(ByVal vstrConstante_Nombre As String) As Variant
    
    Dim i%, strExec$, odynSql As oraDynaset, intIntentos%
    
    Screen.MousePointer = vbHourglass
    
    lodb.Parameters.Add "RET_SALIDA_SQL", Space(1000), ORAPARM_OUTPUT, ORATYPE_VARCHAR2
    strExec = " BEGIN :RET_SALIDA_SQL := " & vstrConstante_Nombre & "; END; "
    On Error GoTo Error
    lodb.ExecuteSQL strExec
    GoTo ok
Error:
    If intIntentos = 1 Then
        intIntentos = intIntentos + 1
        Resume
    ElseIf intIntentos >= 3 Then
        MsgBox "COMUNIQUESE CON SISTEMAS" & Chr(13) & Chr(13) & Chr(13) & Err.Description, vbCritical, "Error Crítico en el Intento " & (intIntentos)
    Else
        MsgBox "PRESIONE ACEPTAR PARA INTENTAR NUEVAMENTE" & Chr(13) & Chr(13) & Chr(13) & Err.Description, vbExclamation, "Ejecutando Consulta en el intento " & intIntentos
    End If
    intIntentos = intIntentos + 1
    If intIntentos <= 5 Then Resume Else GoTo Salida
ok:
    Const_Val = lodb.Parameters("RET_SALIDA_SQL").Value
    
Salida:
On Error GoTo 0
    lodb.Parameters.Remove "RET_SALIDA_SQL"
    Screen.MousePointer = vbDefault
End Function

'******** Obtiene los MaxLen de todas los TextBox ********
Function maxLen(ByVal vstrTabla$, _
                ByVal vstrColumna$, _
                Optional ByVal vstrOwner$, _
                Optional ByVal vstrDBlink$) As Integer
    Screen.MousePointer = vbHourglass
    Dim odynMaxLen As oraDynaset, strOwner$
    On Error GoTo selerror
    
    strOwner = IIf(Trim(vstrOwner) = "", lodb.Connect, Trim(vstrOwner))
        
    maxLen = FN_Valor("CMR.PKG_UTIL.FN_DATA_LENGTH", strOwner, Trim(vstrTabla), Trim(vstrColumna), vstrDBlink)
    
    On Error GoTo 0
    GoTo Final
selerror:
    maxLen = 0
    MsgBox "La tabla " & vstrTabla & vstrDBlink & " o la columna " & vstrColumna & " no pertenecen al esquema " & strOwner & "@" & lodb.DatabaseName & Chr(13) & _
         "Contáctese con el Departamento de Sistemas", vbExclamation, "Atención " & App.EXEName
    On Error GoTo 0
Final:
    Set odynMaxLen = Nothing
    Screen.MousePointer = vbDefault
End Function

Function AppVersion(ByVal vstrAplicacion$, ByVal vstrVersion$) As Boolean
Dim varValores As Variant
Dim varIO As Variant
    AppVersion = True
    varValores = Array(vstrAplicacion, vstrVersion)
    varIO = Array(entrada, entrada)
    On Error GoTo Control
    Call SP("DBMS_APPLICATION_INFO.SET_MODULE", varValores, varIO)
    Exit Function
Control:
    MsgBox "Error en registrar Aplicación" & Chr(13) & Err.Description, vbCritical, "Error)"
    AppVersion = False
End Function


''Sub Describir(ByVal vstrNomOraObj$)
''Dim xMD As OraMetaData
''Dim x$
''
''On Error GoTo Error_Des
''    x = vstrNomOraObj
''    Set xMD = lodb.Describe(x)
''    MsgBox x & " is of the type " & xMD.Type
''    RecursiveDescribe x, xMD
''    Exit Sub
'''''Dim OraDynaset As OraDynaset
'''''Dim OraMetaData As OraMetaData
'''''Dim OraMDAttribute As OraMDAttribute
'''''Dim attrList As OraMetaData
'''''
'''''Dim attr As OraMetaData
'''''
'''''Dim NumAttributes As Integer
'''''Dim NumMethods As Integer
'''''Dim i%
'''''
'''''On Error GoTo Error_Des
'''''    Set OraMetaData = lodb.Describe(vstrNomOraObj)
'''''    NumAttributes = OraMetaData!NumAttributes
'''''    NumMethods = OraMetaData!NumMethods
'''''
'''''    MsgBox "El Objeto «" & vstrNomOraObj & "» tiene " & NumAttributes & " Atributos"
'''''    MsgBox "El Objeto «" & vstrNomOraObj & "» tiene  " & NumMethods & "  Metodos "
'''''
'''''    'Retrieve the attribute list of this type object
'''''
'''''    Set attrList = OraMetaData!Attributes.Value
'''''    'Display the name and data type of each attribute
'''''
'''''
'''''    For i = 0 To attrList.Count - 1
'''''
'''''       Set attr = attrList(i).Value
'''''
'''''     ' each attr is actually an OraMetaData
'''''
'''''       MsgBox "Attribute Name: " & attr!name
'''''
'''''       MsgBox "Attribute Type: " & attr!TypeName
'''''
'''''    Next i
''Error_Des:
''    MsgBox Err.Description, vbCritical, "Error"
''    On Error GoTo 0
''End Sub
''
''
''
''Private Sub RecursiveDescribe(name$, xMD As OraMetaData)
''Dim xMDAttr As OraMDAttribute
''Dim i%
''    For i = 0 To xMD.Count - 1
''       Set xMDAttr = xMD.Attribute(i)
''       ' If an attribute can be described further, describe it,
''       ' otherwise display its attribute name & value
''       If (xMDAttr.IsMDObject) Then
''         RecursiveDescribe xMDAttr.name, xMDAttr.Value
''       Else
''         MsgBox name & "->" & xMDAttr.name & " = " & xMDAttr.Value
''       End If
''    Next i
''End Sub
''



Sub LimpiaParametros()
Dim objParam As Object
    
    While lodb.Parameters.Count > 0
        For Each objParam In lodb.Parameters
            lodb.Parameters.Remove objParam.name
        Next objParam
    Wend

'    While lodb.Parameters.Count > 0
'        lodb.Parameters.Remove lodb.Parameters(i).name
'    Wend
End Sub

Function Fecha_Servidor() As Date
    Fecha_Servidor = CDate(FN_Valor("TO_CHAR(SYSDATE, 'DD/MM/YYYY')"))
End Function

Sub createLog(ByVal strExtention As String, ByVal strNameFile As String, ByVal strSQL As String, Optional strParams As String, Optional strError As String, Optional isError As String)
    Dim strPath As String
    'Dim strExtention As String
    'Dim strNameFile As String
    Dim strDateTime As String
    Dim strDateNow As String
    Dim myStamp As Date
    Dim dateDiff As String
    
    If gstrFlagLogFile2 = "1" Then
    Else
        If gstrIndCreaLogError = "1" And isError = "1" Then
        Else
            Exit Sub
        End If
    End If
    strPath = gstrPathLog '"" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "PATHDLVLOG", "10") '"C:\Documents and Settings\aescate\Escritorio\ecastillo\demo\"
    'strExtention = "log"
    strDateNow = DateTime.Now
    If Len(Trim(strPath)) = 0 Then strPath = App.Path & "\log\"
    strDateTime = Fecha_Servidor
    strDateTime = Replace(strDateTime, "/", "-")
    strNameFile = strNameFile & "_" & strDateTime
    Dim File As String
    If right$(strPath, 1) <> "\" Then strPath = strPath & "\"
    If Trim$(strExtention) = "" Then
        strExtention = "*.*"
    ElseIf left$(strExtention, 2) <> "*." Then
        strExtention = "*." & strExtention
    End If
    File = Dir$(strPath & strExtention)
    Do While Len(File)
        myStamp = FileDateTime(strPath & File)
        dateDiff = DateTime.dateDiff("d", myStamp, strDateTime)
        If dateDiff > 2 Then
            Kill (strPath & File)
        End If
        File = Dir$
    Loop
    If left$(strExtention, 2) = "*." Then strExtention = "." & right$(strExtention, 3)
    
    Open strPath & strNameFile & strExtention For Append As #1
        'Print #1, "=============================================================================================="
        Print #1, "Date: " & strDateNow
        Print #1, "User: " & objUsuario.Codigo
        Print #1, "SQL: " & strSQL
        Print #1, "Params: " & strParams
        If Len(Trim(strError)) > 0 Then
            Print #1, "Error: " & strError
        End If
        Print #1, "=============================================================================================="
    Close #1
End Sub

