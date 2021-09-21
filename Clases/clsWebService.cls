VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWebService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Sys call to convert multiple byte chars to a char
Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long) As Long
    
Const CP_UTF8 As Long = 65001 ' UTF-8 Code Page
Private strResponseExample As String

Private Function PostRequest2(ByVal host As String, ByVal URL As String, ByVal data As Variant, Optional Auth As String, Optional fnc As String) As Collection
On Error GoTo Err
    Dim objXmlHttp As MSXML2.XMLHTTP60
    Set objXmlHttp = New MSXML2.XMLHTTP60
    'I.ECASTILLO 30.06.2021
    Dim params As New XArrayDB
    Dim values As New XArrayDB
    Dim errors As New XArrayDB
    Dim i As Integer
    'F.ECASTILLO 30.06.2021
    Dim strError As String
    data = Replace(data, "'", """")
    data = Replace(data, "|", "'")
'    Debug.Print vbNewLine & data
    Dim response As String
    Dim codeError As String
    objXmlHttp.Open "POST", URL, False
    objXmlHttp.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    objXmlHttp.setRequestHeader "Host", host
    objXmlHttp.setRequestHeader "Content-Length", "Len(data)"
    'Authorization:Basic cmVzZXJ2YTpTdDBjazIwMjA
    If Len(Auth) > 0 Then objXmlHttp.setRequestHeader "Authorization", "Basic " & Auth
    'I.ECASTILLO 30.06.2021
    params.ReDim 0, -1, 0, 9
    values.ReDim 0, -1, 0, 9
    errors.ReDim 0, -1, 0, 9
    i = 0
    params.AppendRows
    values.AppendRows
    
    params(i, 0) = "" & "Action"
    values(i, 0) = "" & fnc
    params(i, 1) = "" & "Method"
    values(i, 1) = "" & "POST"
    params(i, 2) = "" & "EndPoint"
    values(i, 2) = "" & URL
    params(i, 3) = "" & "Request"
    values(i, 3) = "" & "N"
    If IsMissing(data) = False Then values(i, 3) = "" & data
    params(i, 4) = "" & "Key"
    values(i, 4) = "" & "N"
    If Len(Auth) > 0 Then values(i, 4) = "" & Auth
    'F.ECASTILLO 30.06.2021
    If IsMissing(data) = False Then
        objXmlHttp.send data
    Else
        objXmlHttp.send
    End If

    codeError = objXmlHttp.Status
    response = objXmlHttp.responseText
'    codeError = "200"
'    response = FileText
    strError = response
    If codeError <> "200" Then GoTo Err
    Set objXmlHttp = Nothing
    'I.ECASTILLO 30.06.2021
    params(i, 5) = "" & "Status"
    values(i, 5) = "" & codeError
    params(i, 6) = "" & "Response"
    values(i, 6) = "" & response
    'F.ECASTILLO 30.06.2021
    If response <> "[]" And response <> "" Then
        response = Replace(response, "'", "|")
        response = "{'data': " & response & "}"
        response = Replace(response, """", "'")
        Debug.Print response
        Set PostRequest2 = JsonConverter.ParseJson(response)
    Else
        Set PostRequest2 = Nothing
    End If
    'Set PostRequest2 = responseParse
    'I.ECASTILLO 30.06.2021
    createLog gstrFlagLogFile1, gstrFlagLogBD1, "log", "dlv_unificado_ws", params, values, errors
    'F.ECASTILLO 30.06.2021
    Exit Function
Err:
    'I.ECASTILLO 30.06.2021
    params(i, 5) = "" & "Status"
    values(i, 5) = "" & IIf(codeError = "", objXmlHttp.Status, codeError)
    params(i, 6) = "" & "Error"
    values(i, 6) = "" & strError
    createLog gstrFlagLogFile1, gstrFlagLogBD1, "log", "dlv_unificado_ws", params, values, errors, "1"
    'F.ECASTILLO 30.06.2021
    If codeError <> "200" Then MsgBox "Error: " & IIf(codeError = "", objXmlHttp.Status, codeError) & vbNewLine & "Existe un inconveniente con clsWebServices." & fnc & ": " & vbNewLine & strError, vbOKOnly + vbCritical, App.ProductName
    Exit Function
'    Err.Raise Err.Number, "clsWebService.PostRequest2", Err.Description
    MsgBox "Error: Existe un inconveniente con WebService.", vbOKOnly + vbCritical, App.ProductName
End Function

'I.ECASTILLO 27.10.2020
Private Function SendRequest(ByVal method As String, ByVal host As String, ByVal URL As String, Optional data As Variant, Optional Auth As String, Optional fnc As String) As Dictionary
On Error GoTo Err
    Dim objXmlHttp As MSXML2.XMLHTTP60
    Set objXmlHttp = New MSXML2.XMLHTTP60
    'I.ECASTILLO 30.06.2021
    Dim params As New XArrayDB
    Dim values As New XArrayDB
    Dim errors As New XArrayDB
    Dim i As Integer
    'F.ECASTILLO 30.06.2021
    Dim strError As String
    If IsMissing(data) = False Then
        data = Replace(data, "'", """")
    End If
'    Debug.Print vbNewLine & data
    Dim response As String
    Dim codeError As String
    Select Case method
        Case "POST", "PUT", "PATCH"
            objXmlHttp.Open method, URL, False
            objXmlHttp.setRequestHeader "Content-Type", "application/json; charset=utf-8"
            objXmlHttp.setRequestHeader "Host", host
            objXmlHttp.setRequestHeader "Content-Length", "Len(data)"
            If Len(Auth) > 0 Then objXmlHttp.setRequestHeader "Authorization", "Basic " & Auth
        Case "GET"
            objXmlHttp.Open "GET", URL, False
            'objXmlHttp.setRequestHeader "Content-Type", "application/json"
            objXmlHttp.setRequestHeader "Host", host
        'Case "PUT"
        'Case "PATCH"
    End Select
    'I.ECASTILLO 30.06.2021
    params.ReDim 0, -1, 0, 9
    values.ReDim 0, -1, 0, 9
    errors.ReDim 0, -1, 0, 9
    i = 0
    params.AppendRows
    values.AppendRows
    
    params(i, 0) = "" & "Action"
    values(i, 0) = "" & fnc
    params(i, 1) = "" & "Method"
    values(i, 1) = "" & method
    params(i, 2) = "" & "EndPoint"
    values(i, 2) = "" & URL
    params(i, 3) = "" & "Request"
    values(i, 3) = "" & "N"
    If IsMissing(data) = False Then values(i, 3) = "" & data
    params(i, 4) = "" & "Key"
    values(i, 4) = "" & "N"
    If Len(Auth) > 0 Then values(i, 4) = "" & Auth
    'F.ECASTILLO 30.06.2021
    If IsMissing(data) = False Then
        objXmlHttp.send data
    Else
        objXmlHttp.send
    End If
    
    codeError = objXmlHttp.Status
    response = objXmlHttp.responseText
    strError = response
    If codeError <> "200" Then GoTo Err
    Set objXmlHttp = Nothing
    'I.ECASTILLO 30.06.2021
    params(i, 5) = "" & "Status"
    values(i, 5) = "" & codeError
    params(i, 6) = "" & "Response"
    values(i, 6) = "" & response
    'F.ECASTILLO 30.06.2021
    If Len(Trim(response)) > 2 Then
        response = Replace(response, "'", "|")
        response = "{'data': " & response & "}"
        response = Replace(response, """", "'")
        Debug.Print response
        Set SendRequest = JsonConverter.ParseJson(response)
    Else
        Set SendRequest = Nothing
    End If
    'Set PostRequest2 = responseParse
    'I.ECASTILLO 30.06.2021
    createLog gstrFlagLogFile1, gstrFlagLogBD1, "log", "dlv_unificado_ws", params, values, errors
    'F.ECASTILLO 30.06.2021
    Exit Function
Err:
    'I.ECASTILLO 30.06.2021
    params(i, 5) = "" & "Status"
    values(i, 5) = "" & IIf(codeError = "", objXmlHttp.Status, codeError)
    params(i, 6) = "" & "Error"
    values(i, 6) = "" & strError
    createLog gstrFlagLogFile1, gstrFlagLogBD1, "log", "dlv_unificado_ws", params, values, errors, "1"
    'F.ECASTILLO 30.06.2021
    If codeError <> "200" Then MsgBox "Error: " & IIf(codeError = "", objXmlHttp.Status, codeError) & vbNewLine & "Existe un inconveniente con clsWebServices." & fnc & ": " & vbNewLine & strError, vbOKOnly + vbCritical, App.ProductName
    Exit Function
'    Err.Raise Err.Number, "clsWebService.PostRequest2", Err.Description
    MsgBox "Error: Existe un inconveniente con WebService.", vbOKOnly + vbCritical, App.ProductName
End Function
'F.ECASTILLO 27.10.2020

Public Function GetStockTotal(ByVal Cia As String, ByVal Tipo As String, ByVal CodigoLocal As String, ByVal Codigo As Variant, ByVal nCia As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim data As Dictionary
    Set data = New Dictionary
    Dim strCia As String
    Dim sendString, detail, COMMA As String
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS101") '"http://mapimfqaomni.backend.cindibyinkafarma.com/inkafarma-backend/drugstores/drugstoreStockAll"
    host = GetDomainHost(URL)
    
    Dim x As Integer
    
'    If CIA = "94" Then
'        strCia = "INKA"
'    Else
'        strCia = "MIFA"
'    End If
    'I.ECASTILLO 17.09.2021 PARAMETRIZAR MARCA
    strCia = nCia
    If Len(Trim(strCia)) = 0 Then
        Select Case Cia
            Case "94", "93", "92", "1DLV"
                strCia = "INKA"
            Case Else
                strCia = "MIFA"
        End Select
    End If
    'F.ECASTILLO 17.09.2021
    
    '{'cia':'INKA/MIFA','typeSearch':'all/department/province','shoppingCart':[{'productId':'009465'},{'productId':'009466'}]}
    For x = 0 To UBound(Codigo)
        If x <> UBound(Codigo) Then COMMA = ","
        detail = detail & "{'productId': '" & Codigo(x) & "'}" & COMMA
        'detail = detail & "{'productId': '001807'}" & comma
    Next x
'    data.RemoveAll
'    data.item("cia") = strCia
'    data.item("typeSearch") = Tipo
'    data.item("shoppingCart") = arr
    
    sendString = "{" & _
                    "'cia': '" & strCia & "'," & _
                    "'typeSearch': '" & Tipo & "'," & _
                    "'drugstoreId': '" & CodigoLocal & "'," & _
                    "'shoppingCart': [" & detail & "]}"
                    
    Set objResponse = SendRequest("POST", host, URL, sendString, , "GetStockTotal")
    Set GetStockTotal = objResponse
'    Dim example As String
    '{'data':[{'drugstoreId':'aw0','drugstore':'Botica aws','listProducts':[{'productId':'001944','product':'Glicenex SR 750 Mg Prolongada\n(1 Comprimido)','Price':12.0'isFractional':1,'requiereReceta':0,'fractionType':[{'fractionatedText':'PART_MODE','fractionatedDesc':'Dscripcion','stock':60.0},{'fractionatedText':'PACK_MODE','fractionatedDesc':'Dscripcion','stock':2.0}]}]},
    ' {'drugstoreId':'BTL','drugstore':'Botica BTL','listProducts':[{'productId':'001944','product':'Glicenex SR 750 Mg Prolongada\n(1 Comprimido)','Price':12.0'isFractional':1,'requiereReceta':0,'fractionType':[{'fractionatedText':'PART_MODE','fractionatedDesc':'Dscripcion','stock':30.0},{'fractionatedText':'PACK_MODE','fractionatedDesc':'Dscripcion','stock':2.0}]}]}]}
    'example = "{'data':[{'id':1,'name':'28DEJULIO','stock':'CAMPONUMERICOPORCONFIRMAR','stockFractionated':'CAMPONUMERICOPORCONFIRMAR'},{'id':2,'name':'CAVENECIA','stock':'CAMPONUMERICOPORCONFIRMAR','stockFractionated':'CAMPO NUMERICO POR CONFIRMAR'}]}"
'    example = "{'data':"
'    example = example & FileText
'    example = example & "}"
'
'    Debug.Print example
'    Set GetStockTotal = JsonConverter.ParseJson(example)
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebService.GetStockTotal", Err.Description
End Function


Public Function GetStockRealWS(ByVal Cia As String, ByVal CodigoLocal As String, ByVal Codigo As Variant, ByVal nCia As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim data As Dictionary
    'Dim arr As XArrayDB
    Set data = New Dictionary
    Dim strCia As String
    Dim sendString, detail, COMMA As String
    Dim x As Integer
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS102") '"http://mapimfqaomni.backend.cindibyinkafarma.com/inkafarma-backend/drugstores/drugstoresStock"
    host = GetDomainHost(URL)
    
'    If CIA = "1DLV" Then
'        strCia = "INKA"
'    ElseIf CIA = "0DLV" Then
'        strCia = "MIFA"
'    End If
    'I.ECASTILLO 17.09.2021 PARAMETRIZAR MARCA
    strCia = nCia
    If Len(Trim(strCia)) = 0 Then
        Select Case Cia
            Case "94", "93", "92", "1DLV"
                strCia = "INKA"
            Case Else
                strCia = "MIFA"
        End Select
    End If
    'F.ECASTILLO 17.09.2021
    
    '{"drugstore":123,"shoppingCart":[{"productId":"009465","fractionalMode":false}{"productId":"009466","fractionalMode":false}]}
    For x = 0 To UBound(Codigo)
        COMMA = ""
        If x <> UBound(Codigo) Then COMMA = ","
        detail = detail & "{'productId': '" & Codigo(x) & "'}" & COMMA
        'detail = detail & "{'productId': '001807'}" & comma
        Debug.Print vbNewLine & JsonConverter.ConvertToJson(detail)
    Next x
    
'    data.RemoveAll
'    data.item("cia") = strCia
'    data.item("drugstore") = "506" 'CodigoLocal
'    data.item("shoppingCart") = arr
    sendString = "{" & _
                    "'cia': '" & strCia & "'," & _
                    "'drugstore': '" & CodigoLocal & "'," & _
                    "'shoppingCart': [" & detail & "]}"
    
    Set objResponse = SendRequest("POST", host, URL, sendString, , "GetStockRealWS")
    Set GetStockRealWS = objResponse
''    Dim example As String
''    'example = "{'data':[{'productId':'009465','fractionType':[{'fractionatedText':'PART_MODE','stock':15}]}]}"
''    example = "{'data':[{'productId': '527722','product': null,'isFractional': 1,'fractionType': [{'fractionatedText': 'PACK_MODE','fractionatedDesc': 'BLISTER/10','stock': 0,'unitQuantity': 2},{'fractionatedText': 'PART_MODE','fractionatedDesc': 'UNIDAD','stock': 1,'unitQuantity': 1}],'fecModSym': '2021-08-28 10:09:50','fecSym': '2021-08-28 10:10:08','fechSysDate': '2021-08-28 10:46:55','fecSymLocal': '2021-08-28 10:45:54','fecHeartbeat': '2021-08-28 10:45:54'}]}"
''    Debug.Print example
''    Set GetStockRealWS = JsonConverter.ParseJson(example)
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebService.GetStockRealWS", Err.Description
End Function

Public Function ReservaStock(ByVal codGrupoCia As String, ByVal Proforma As String, ByVal gloss As String, ByVal drugstoreId As String, ByVal Producto As XArrayDB, ByVal codUser As String, ByVal nCia As String) As Dictionary
'On Error GoTo Err
    Dim objResponse As Dictionary
    Dim data As Dictionary
    Set data = New Dictionary
    Dim strCia As String
    Dim sendString, detail, COMMA As String
    Dim x As Integer
    Dim URL As String
    Dim host As String
    Dim Key As String
    URL = GetUrlWS("URLWS103") '"http://reservastock.farmaciasperuanas.pe/api/v1/stockReserve"
    host = GetDomainHost(URL)
    Key = GetKeyUrlWS("KEYWS103") '"cmVzZXJ2YTpTdDBjazIwMjA"
    'I.ECASTILLO 17.09.2021 PARAMETRIZAR MARCA
    strCia = nCia
    If Len(Trim(strCia)) = 0 Then
        Select Case codGrupoCia
            Case "94", "93", "92", "1DLV"
                strCia = "INKA"
            Case Else
                strCia = "MIFA"
        End Select
    End If
    'F.ECASTILLO 17.09.2021
    For x = 0 To Producto.Count(1) - 1
        COMMA = ""
        If x < Producto.Count(1) - 1 Then COMMA = ","
        detail = detail & "{'productId': '" & Producto(x, 0) & "','quantity': " & Producto(x, 1) & "}" & COMMA
        'detail = detail & "{'productId': '001807'}" & comma
        Debug.Print vbNewLine & JsonConverter.ConvertToJson(detail)
    Next x

    sendString = "{" & _
                    "'deliveryService': 'POSU'," & _
                    "'cia': '" & strCia & "'," & _
                    "'localId': '" & drugstoreId & "'," & _
                    "'glosa': '" & gloss & "'," & _
                    "'voucher': '" & Proforma & "'," & _
                    "'userId': '" & codUser & "'," & _
                    "'reserveList': [" & detail & "]}"
    Set objResponse = SendRequest("POST", host, URL, sendString, Key, "ReservaStock")
    Set ReservaStock = objResponse
'    Dim example As String
'    example = "{'data':{'errorCode':null,'successCode':'SUCCESS','message':'Reserva exitosa'}}"
'    Set ReservaStock = JsonConverter.ParseJson(example)
    Exit Function
'Err:
'    Err.Raise Err.Number, "clsWebService.ReservaStock", Err.Description
End Function

Public Function AnulaReservaStock(ByVal codGrupoCia As String, ByVal Proforma As String, ByVal gloss As String, ByVal drugstoreId As String, ByVal Producto As XArrayDB, ByVal codUser As String, ByVal nCia As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim data As Dictionary
    Set data = New Dictionary
    Dim strCia As String
    Dim sendString, detail, COMMA As String
    Dim x As Integer
    Dim URL As String
    Dim host As String
    Dim Key As String
    URL = GetUrlWS("URLWS104") '"http://reservastock.farmaciasperuanas.pe/api/v1/stockRelease"
    host = GetDomainHost(URL)
    Key = GetKeyUrlWS("KEYWS104") '"cmVzZXJ2YTpTdDBjazIwMjA"
    
'    If codGrupoCia = "94" Then
'        strCia = "INKA"
'    Else
'        strCia = "MIFA"
'    End If
    'I.ECASTILLO 17.09.2021 PARAMETRIZAR MARCA
    strCia = nCia
    If Len(Trim(strCia)) = 0 Then
        Select Case codGrupoCia
            Case "94", "93", "92", "1DLV"
                strCia = "INKA"
            Case Else
                strCia = "MIFA"
        End Select
    End If
    'F.ECASTILLO 17.09.2021
    'data.item("codGrupoCia") = strCia
    'data.item("proforma") = proforma
    'data.item("gloss") = gloss
    'data.item("drugstoreId") = drugstoreId
    'data.item("codProd") = CodProd
    'data.item("newQuantity") = newQuantity
    'data.item("typeDoc") = typeDoc
    'data.item("codUser") = codUser
    For x = 0 To Producto.Count(1) - 1
        COMMA = ""
        If x <> Producto.Count(1) - 1 Then COMMA = ","
        detail = detail & "{'productId': '" & Producto(x, 0) & "','quantity': '" & Producto(x, 1) & "'}" & COMMA
        'detail = detail & "{'productId': '001807'}" & comma
        Debug.Print vbNewLine & JsonConverter.ConvertToJson(detail)
    Next x
    
    sendString = "{" & _
                    "'deliveryService': 'POSU'," & _
                    "'cia': '" & strCia & "'," & _
                    "'localId': '" & drugstoreId & "'," & _
                    "'glosa': '" & gloss & "'," & _
                    "'voucher': '" & Proforma & "'," & _
                    "'userId': '" & codUser & "'," & _
                    "'reserveList': [" & detail & "]}"
    
    Set objResponse = SendRequest("POST", host, URL, sendString, Key, "AnulaReservaStock")
    Set AnulaReservaStock = objResponse
'    Dim example As String
'    example = "{'data':{'errorCode':null,'successCode':'SUCCESS','message':'Anulación exitosa'}}"
'    Set AnulaReservaStock = JsonConverter.ParseJson(example)
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebService.AnulaReservaStock", Err.Description
End Function

Public Function ReservadeStock(ByVal deliveryService As String, ByVal Cia As String, ByVal localId As String, ByVal Glosa As String, ByVal voucher As String, ByVal userId As String, ByVal reserveList As XArrayDB, ByVal nCia As String) As Dictionary
'On Error GoTo Err
    Dim objResponse As Dictionary
    Dim data As Dictionary
    Set data = New Dictionary
    Dim strCia As String
    Dim sendString, detail, COMMA As String
    Dim URL As String
    Dim host As String
    Dim Key As String
    URL = GetUrlWS("URLWS105") '"http://reservastock.farmaciasperuanas.pe/api/v1/stockReserve"
    host = GetDomainHost(URL)
    Key = GetKeyUrlWS("KEYWS105") '"cmVzZXJ2YTpTdDBjazIwMjA"
    
    Dim x As Integer
    'I.ECASTILLO 17.09.2021 PARAMETRIZAR MARCA
    strCia = nCia
    If Len(Trim(strCia)) = 0 Then
        Select Case Cia
            Case "94", "93", "92", "1DLV"
                strCia = "INKA"
            Case Else
                strCia = "MIFA"
        End Select
    End If
    'F.ECASTILLO 17.09.2021
    
    'Reserva de Stock (POS UNIFICADO)
    For x = 0 To reserveList.Count(1) - 1
        COMMA = ""
        If x <> reserveList.Count(1) - 1 Then COMMA = ","
         'detail = detail & "{'productId': '" & reserveList(X, 0) & "','quantity': '" & reserveList(X, 1) & "'}" & COMMA
        detail = detail & "{'productId': '" & reserveList(x, 0) & "','quantity': '" & reserveList(x, 1) & "','fractional': '" & IIf(reserveList(x, 2) = 1, "S", "N") & "','quantityFractional': " & reserveList(x, 3) & "}" & COMMA
        
        Debug.Print vbNewLine & JsonConverter.ConvertToJson(detail)
    Next x
    
    sendString = "{" & _
                    "'deliveryService': '" & deliveryService & "'," & _
                    "'cia': '" & strCia & "'," & _
                    "'localId': '" & localId & "'," & _
                    "'glosa': '" & Glosa & "'," & _
                    "'voucher': '" & voucher & "'," & _
                    "'userId': '" & userId & "'," & _
                    "'reserveList': [" & detail & "]}"
                    
    Set objResponse = SendRequest("POST", host, URL, sendString, Key, "ReservadeStock")
    Set ReservadeStock = objResponse

'    Dim example As String
    'Cuando el producto no existe:
    '{ 'code': 'SUCCESS', 'reserveList': [ { 'productId': '51533512', 'quantity': 2, 'status': '1403', 'description': 'ORA-01403: no data found\nORA-06512: at \'PTOVENTA.DIGITAL_CONSUMO\', line 44\nORA-06512: at line 1\n' }, { 'productId': '19000212', 'quantity': 1, 'status': '1403', 'description': 'ORA-01403: no data found\nORA-06512: at \'PTOVENTA.DIGITAL_CONSUMO\', line 44\nORA-06512: at line 1\n' } ] }
    
     'Cuando el producto no existe:
    '{ 'code': 'ERROR_INSINK_NOT_FOUND_LOCAL_ID', 'message': 'Local ID not Found in ExternalDatabaseParameters Table', 'reserveList': null }

'    example = "{ 'code': 'SUCCESS', 'reserveList': [ { 'productId': '515335', 'quantity': 2, 'status': '0' }, { 'productId': '190002', 'quantity': 1, 'status': '0' } ] } "
'
'
'    Set GetReservaStock = JsonConverter.ParseJson(example)
    Exit Function
'Err:
'    Err.Raise Err.Number, "clsWebService.ReservadeStock", Err.Description
End Function

Public Function LiberaciondeStock(ByVal deliveryService As String, ByVal Cia As String, ByVal localId As String, ByVal Glosa As String, ByVal voucher As String, ByVal userId As String, ByVal reserveList As XArrayDB, ByVal nCia As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim data As Dictionary
    Set data = New Dictionary
    Dim strCia As String
    Dim sendString, detail, COMMA As String
    Dim URL As String
    Dim host As String
    Dim Key As String
    URL = GetUrlWS("URLWS106") '"http://reservastock.farmaciasperuanas.pe/api/v1/stockRelease"
    host = GetDomainHost(URL)
    Key = GetKeyUrlWS("KEYWS106") '"cmVzZXJ2YTpTdDBjazIwMjA"
    
    Dim x As Integer
    
'    If CIA = "94" Then
'        strCia = "INKA"
'    Else
'        strCia = "MIFA"
'    End If
    'I.ECASTILLO 17.09.2021 PARAMETRIZAR MARCA
    strCia = nCia
    If Len(Trim(strCia)) = 0 Then
        Select Case Cia
            Case "94", "93", "92", "1DLV"
                strCia = "INKA"
            Case Else
                strCia = "MIFA"
        End Select
    End If
    'F.ECASTILLO 17.09.2021
    'Reserva de Stock (POS UNIFICADO)
    For x = 0 To reserveList.Count(1) - 1
        COMMA = ""
        If x <> reserveList.Count(1) - 1 Then COMMA = ","
         detail = detail & "{'productId': '" & reserveList(x, 0) & "','quantity': '" & reserveList(x, 1) & "','fractional':'" & IIf(reserveList(x, 2) = 1, "S", "N") & "','quantityFractional':'" & reserveList(x, 3) & "'}" & COMMA
        'detail = detail & "{'productId': '001807'}" & comma
        Debug.Print vbNewLine & JsonConverter.ConvertToJson(detail)
    Next x
    
    sendString = "{" & _
                    "'deliveryService': '" & deliveryService & "'," & _
                    "'cia': '" & strCia & "'," & _
                    "'localId': '" & localId & "'," & _
                    "'glosa': '" & Glosa & "'," & _
                    "'voucher': '" & voucher & "'," & _
                    "'userId': '" & userId & "'," & _
                    "'reserveList': [" & detail & "]}"
                    
    'Set objResponse = PostRequest2(host, URL, sendString, Key, "LiberaciondeStock")
    Set objResponse = SendRequest("POST", host, URL, sendString, Key, "LiberaciondeStock")
    Set LiberaciondeStock = objResponse

'    Dim example As String
'     'Cuando el producto no existe:
'    '{ 'code': 'SUCCESS', 'reserveList': [ { 'productId': '51533512', 'quantity': 2, 'status': '1403', 'description': 'ORA-01403: no data found\nORA-06512: at \'PTOVENTA.DIGITAL_CONSUMO\', line 44\nORA-06512: at line 1\n' }, { 'productId': '19000212', 'quantity': 1, 'status': '1403', 'description': 'ORA-01403: no data found\nORA-06512: at \'PTOVENTA.DIGITAL_CONSUMO\', line 44\nORA-06512: at line 1\n' } ] }
'
'     'Cuando el producto no existe:
'    '{ 'code': 'ERROR_INSINK_NOT_FOUND_LOCAL_ID', 'message': 'Local ID not Found in ExternalDatabaseParameters Table', 'reserveList': null }
'
'
'    example = "{ 'code': 'SUCCESS', 'reserveList': [ { 'productId': '515335', 'quantity': 2, 'status': '0' }, { 'productId': '190002', 'quantity': 1, 'status': '0' } ] }"
'
'
'    Set GetLiberaciondeStock = JsonConverter.ParseJson(example)
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebService.GetReservaStock", Err.Description
End Function

'I.ECASTILLO 12.08.2021
Public Function GetStockRealWSDCSAP(ByVal Cia As String, ByVal CodigoLocal As String, ByVal Codigo As Variant, ByVal nCia As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim data As Dictionary
    Set data = New Dictionary
    Dim strCia As String
    Dim sendString, detail, COMMA As String
    Dim x As Integer
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS107") '"http://mapimfqaomni.backend.cindibyinkafarma.com/inkafarma-backend/drugstores/drugstoresStock"
    host = GetDomainHost(URL)
    
    'I.ECASTILLO 17.09.2021 PARAMETRIZAR MARCA
    strCia = nCia
    If Len(Trim(strCia)) = 0 Then
        Select Case Cia
            Case "94", "93", "92", "1DLV"
                strCia = "INKA"
            Case Else
                strCia = "MIFA"
        End Select
    End If
    'F.ECASTILLO 17.09.2021
    
    For x = 0 To UBound(Codigo)
        COMMA = ""
        If x <> UBound(Codigo) Then COMMA = ","
        detail = detail & "{'productId': '" & Codigo(x) & "'}" & COMMA
    Next x
    sendString = "{" & _
                    "'cia': '" & strCia & "'," & _
                    "'drugstore': '" & CodigoLocal & "'," & _
                    "'shoppingCart': [" & detail & "]}"
    
    Set objResponse = SendRequest("POST", host, URL, sendString, , "GetStockRealWSDCSAP")
    Set GetStockRealWSDCSAP = objResponse
'    Dim example As String
'    'example = "{'data':[{'productId':'206327','product':null,'isFractional':0,'fractionType':[{'fractionatedText':'PACK_MODE','fractionatedDesc':'UN','stock':151345,'unitQuantity':1}],'fecModSym':'2021-08-12 20:01:04','fecSym':'2021-08-12 20:04:55','fechSysDate':'2021-08-13 15:39:23','fecSymLocal':'2021-08-13 15:39:14','fecHeartbeat':'2021-08-13 15:39:14'}]}"
'    example = "{'data':[{'productId':'118193','product':null,'isFractional':1,'fractionType':[{'fractionatedText':'PACK_MODE','fractionatedDesc':'UN','stock':2557,'unitQuantity':1}],'fecModSym':'2021-09-09 12:08:56','fecSym':'2021-09-09 12:10:45','fechSysDate':'2021-09-09 13:00:19','fecSymLocal':'2021-09-09 12:35:45','fecHeartbeat':null}]}"
'    Set GetStockRealWSDCSAP = JsonConverter.ParseJson(example)
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebService.GetStockRealWSDCSAP", Err.Description
End Function

'F.ECASTILLO 12.08.2021

'I.ECASTILLO 27.10.2020
Public Function listaCapacidades(ByVal CodLocal As String, ByVal segmento As String, ByVal Tipo As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim sendString As String
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS110")
    host = GetDomainHost(URL)
    
    sendString = "{" & _
                 "  'fulfillmentCenterCode': '" & CodLocal & "'," & _
                 "  'channel': 'CALL'," & _
                 "  'segmentGap': " & segmento & "," & _
                 "  'serviceTypeCode': '" & Tipo & "'" & _
                 "}"
    Set objResponse = SendRequest("POST", host, URL, sendString, , "listaCapacidades")
    Set listaCapacidades = objResponse
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebService.listaCapacidades", Err.Description
End Function
Public Function validaCapacidades(ByVal CodLocal As String, ByVal dia As String, ByVal hora As String, ByVal Tipo As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS111")
    host = GetDomainHost(URL)
    URL = Replace(URL, ":dia", dia)
    URL = Replace(URL, ":hora", hora)
    URL = Replace(URL, ":canal", "CALL")
    URL = Replace(URL, ":tipo", Tipo)
    URL = Replace(URL, ":local", CodLocal)
    Set objResponse = SendRequest("GET", host, URL, , , "validaCapacidades")
    Set validaCapacidades = objResponse
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebService.validaCapacidades", Err.Description
End Function
Public Function reservaCapacidades(ByVal CodLocal As String, ByVal dia As String, ByVal hora As String, ByVal Tipo As String, ByVal strProforma As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim sendString As String
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS112")
    host = GetDomainHost(URL)
    URL = Replace(URL, ":dia", dia)
    URL = Replace(URL, ":hora", hora)
    sendString = "{" & _
                 "  'fulfillmentCenterCode': '" & CodLocal & "'," & _
                 "  'channel':'CALL'," & _
                 "  'serviceTypeCode': '" & Tipo & "'," & _
                 "  'ecommerceId': '" & strProforma & "'" & _
                 "}"
    Set objResponse = SendRequest("PUT", host, URL, sendString, , "reservaCapacidades")
    Set reservaCapacidades = objResponse
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebService.reservaCapacidades", Err.Description
End Function
Public Function retornaCapacidades(ByVal CodLocal As String, ByVal dia As String, ByVal hora As String, ByVal Tipo As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim sendString As String
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS113")
    host = GetDomainHost(URL)
    URL = Replace(URL, ":dia", dia)
    URL = Replace(URL, ":hora", hora)
    sendString = "{" & _
                 "  'fulfillmentCenterCode': '" & CodLocal & "'," & _
                 "  'channel':'CALL'," & _
                 "  'serviceTypeCode': '" & Tipo & "'" & _
                 "}"
    Set objResponse = SendRequest("PATCH", host, URL, sendString, , "retornaCapacidades")
    Set retornaCapacidades = objResponse
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebService.retornaCapacidades", Err.Description
End Function
Public Function asignaDigital(ByVal proformaNumber As String, ByVal localCode As String, ByVal ciaCode As String, ByVal localCodeRef As String, ByVal dateCreated As String, ByVal dateConfirmada As String, ByVal processStartConfirm As String, ByVal Source As String, ByVal companyCode As String, _
                              ByVal creditCardProviderId As String, ByVal currencyId As String, ByVal discountApplied As String, ByVal deliveryType As String, ByVal totalTax As String, ByVal amount As String, ByVal flgCappa As String, ByVal user_dni As String, ByVal user_isInkaClub As String, _
                              ByVal user_phone As String, ByVal user_email As String, ByVal user_name As String, ByVal user_lastName As String, ByVal user_isAnonymous As String, ByVal paymentMethod_id As String, ByVal paymentMethod_name As String, _
                              ByVal receipt_ruc As String, ByVal receipt_companyName As String, ByVal receipt_companyAddress As String, ByVal receipt_receiptType_name As String, ByVal argDelivery As XArrayDB, _
                              ByVal paymentAmountDto_productsTotalCostNoDiscount As String, ByVal paymentAmountDto_grossPrice As String, ByVal paymentAmountDto_discountApplied As String, ByVal paymentAmountDto_productsTotalCost As String, ByVal paymentAmountDto_deliveryCost As String, ByVal paymentAmountDto_paidAmountByCoin As String, _
                              ByVal productos As XArrayDB, ByVal vales As XArrayDB, _
                              ByVal agreementFlag As String, ByVal agreementCode As String, ByVal agreementCodCli As String, ByVal agreementBeneficiaryPct As String, ByVal agreementFields As XArrayDB, ByVal dataCli As String, ByVal nameOf As String, _
                              ByVal observation As String, ByVal observationMotorized As String, ByVal flagMultiBrand As String, ByVal deliveryTime As String, ByVal dateConfirmadaEnd As String, ByVal argAdicionales As XArrayDB) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim sendString As String
    Dim sendItems As String
    Dim sendVales As String
    Dim itemsConv As String
    Dim x, i, j As Integer
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS114")
    host = GetDomainHost(URL)
    
    Dim flg_2e_reserva
    flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
    
    
    For x = 0 To vales.Count(1) - 1
        sendVales = sendVales & "{" & _
                                "   'paymentMethod': '" & vales(x, 0) & "'," & _
                                "   'voucherNumber': '" & vales(x, 1) & "'" & _
                                "},"
    Next x
    If Len(sendVales) > 0 Then sendVales = left(sendVales, Len(sendVales) - 1)
    For x = 0 To productos.Count(1) - 1
         sendItems = sendItems & "{" & _
                                "   'productId': '" & productos(x, 0) & "'," & _
                                "   'productPosuId': '" & productos(x, 1) & "'," & _
                                "   'productSapCode': '" & productos(x, 2) & "'," & _
                                "   'name': '" & productos(x, 3) & "'," & _
                                "   'shortDescription': '" & productos(x, 4) & "'," & _
                                "   'brand': '" & productos(x, 5) & "'," & _
                                "   'quantity': '" & productos(x, 6) & "'," & _
                                "   'quantityUnits': '" & productos(x, 7) & "'," & _
                                "   'fractionated': '" & productos(x, 8) & "'," & _
                                "   'unitPrice': '" & productos(x, 9) & "'," & _
                                "   'totalPrice': '" & productos(x, 10) & "'," & _
                                "   'presentation': '" & productos(x, 11) & "'," & _
                                "   'familyType': '" & productos(x, 12) & "'," & _
                                "   'eanCode': '" & productos(x, 13) & "'," & _
                                "   'fractionatedPrice': '" & productos(x, 14) & "'," & _
                                "   'presentationType': '" & productos(x, 15) & "'," & _
                                "   'valFrac': '" & productos(x, 16) & "'," & _
                                "   'valorUmv': '" & productos(x, 17) & "'," & _
                                "   'quantityTotal': '" & productos(x, 18) & "'" & _
                                "},"
    Next x
    If Len(sendItems) > 0 Then sendItems = left(sendItems, Len(sendItems) - 1)
    If IsObject(agreementFields) = False Then GoTo saltaAgreement
    If agreementFields Is Nothing Then GoTo saltaAgreement
    For x = 0 To agreementFields.Count(1) - 1
        itemsConv = itemsConv & "{ " & _
                 "              'codGroupCia': '" & agreementFields(x, 0) & "'," & _
                 "              'codLocal': '" & agreementFields(x, 1) & "'," & _
                 "              'numPedVta': '" & agreementFields(x, 2) & "'," & _
                 "              'codField': '" & agreementFields(x, 3) & "'," & _
                 "              'codAgreement': '" & agreementFields(x, 4) & "'," & _
                 "              'codClient': '" & agreementFields(x, 5) & "'," & _
                 "              'usuCreaPedVtaCli': '" & agreementFields(x, 6) & "'," & _
                 "              'fecModPedVtaCli': '" & agreementFields(x, 7) & "'," & _
                 "              'usuModPedVtaCli': '" & agreementFields(x, 8) & "'," & _
                 "              'descriptionField': '" & agreementFields(x, 9) & "'," & _
                 "              'nameField': '" & agreementFields(x, 10) & "'," & _
                 "              'flgPrint': '" & agreementFields(x, 11) & "'," & _
                 "              'codValueIn': '" & agreementFields(x, 12) & "'" & _
                 "          },"
    Next x
    If Len(itemsConv) > 0 Then itemsConv = left(itemsConv, Len(itemsConv) - 1)
saltaAgreement:
    sendString = "{ 'proformaNumber': '" & proformaNumber & "'," & _
                 "  'localCode': '" & localCode & "'," & _
                 "  'ciaCode': '" & ciaCode & "'," & _
                 "  'localCodeRef': '" & localCodeRef & "'," & _
                 "  'dateCreated': '" & dateCreated & "'," & _
                 "  'dateConfirmada': '" & dateConfirmada & "'," & _
                 "  'deliveryTime': " & procString(deliveryTime) & "," & _
                 "  'dateConfirmadaEnd': '" & dateConfirmadaEnd & "'," & _
                 "  'source': '" & Source & "'," & _
                 "  'companyCode': '" & companyCode & "'," & _
                 "  'creditCardProviderId': '" & creditCardProviderId & "'," & _
                 "  'currencyId': '" & currencyId & "'," & _
                 "  'discountApplied': '" & discountApplied & "'," & _
                 "  'deliveryType': '" & deliveryType & "'," & _
                 "  'totalTax': '" & totalTax & "'," & _
                 "  'amount': '" & amount & "'," & _
                 "  'flagDC':" & flgCappa & "," & _
                 "  'flagMultiBrand':" & flagMultiBrand & ","
    sendString = sendString & "  'user': {" & _
                 "      'dni': '" & user_dni & "'," & _
                 "      'isInkaClub': '" & user_isInkaClub & "'," & _
                 "      'phone': '" & user_phone & "'," & _
                 "      'email': '" & user_email & "'," & _
                 "      'name': '" & user_name & "'," & _
                 "      'lastName': '" & user_lastName & "'," & _
                 "      'isAnonymous': '" & user_isAnonymous & "'" & _
                 "  },"
    sendString = sendString & "  'paymentMethod': {" & _
                 "      'id': '" & paymentMethod_id & "'," & _
                 "      'name': '" & paymentMethod_name & "'" & _
                 "  },"
    sendString = sendString & "  'receipt': {" & _
                 "      'ruc' : '" & receipt_ruc & "'," & _
                 "      'companyName' : '" & receipt_companyName & "'," & _
                 "      'companyAddress' : '" & receipt_companyAddress & "'," & _
                 "      'receiptType': {" & _
                 "          'name': '" & receipt_receiptType_name & "'" & _
                 "      }" & _
                 "  },"
                 
    If argDelivery.Count(1) - 1 >= 0 Then
        sendString = sendString & "  'deliveryAddress': {" & _
                 "      'street': '" & argDelivery(0, 0) & "'," & _
                 "      'number': '" & argDelivery(0, 1) & "'," & _
                 "      'apartment': '" & argDelivery(0, 2) & "'," & _
                 "      'country': '" & argDelivery(0, 3) & "'," & _
                 "      'city': '" & argDelivery(0, 4) & "'," & _
                 "      'district': '" & argDelivery(0, 5) & "'," & _
                 "      'latitude': '" & argDelivery(0, 6) & "'," & _
                 "      'longitude': '" & argDelivery(0, 7) & "'," & _
                 "      'receiverName': '" & argDelivery(0, 8) & "'"
        If gstrIndRAv4 = "1" Then
        sendString = sendString & "," & _
                 "      'notes': '" & argDelivery(0, 9) & "'," & _
                 "      'urbanization': '" & argDelivery(0, 10) & "'"
        End If
        sendString = sendString & "  },"
    End If
    
    sendString = sendString & "  'paymentAmountDto': {" & _
                 "      'productsTotalCostNoDiscount': '" & paymentAmountDto_productsTotalCostNoDiscount & "'," & _
                 "      'grossPrice': '" & paymentAmountDto_grossPrice & "'," & _
                 "      'discountApplied': '" & paymentAmountDto_discountApplied & "'," & _
                 "      'productsTotalCost': '" & paymentAmountDto_productsTotalCost & "'," & _
                 "      'deliveryCost': '" & paymentAmountDto_deliveryCost & "'," & _
                 "      'paidAmountByCoin': '" & paymentAmountDto_paidAmountByCoin & "'" & _
                 "  }," & _
                 " 'items': [" & _
                        sendItems & _
                 "  ],"
    If flg_2e_reserva = "0" Then
    Else
        sendString = sendString & "'agreementFlag': '" & agreementFlag & "'," & _
                     "      'agreementCode': '" & agreementCode & "'," & _
                     "      'agreementCodCli': '" & agreementCodCli & "'," & _
                     "      'agreementBeneficiaryPct': '" & agreementBeneficiaryPct & "', " & _
                     "      'agreementFields': [" & _
                                itemsConv & _
                     "      ]," & _
                     "     'dataCli': '" & dataCli & "'," & _
                     "     'nameOf': '" & nameOf & "',"
    End If
    sendString = sendString & _
                 "  'vouchers': [" & _
                        sendVales & _
                 "  ]," & _
                 "  'observation': '" & observation & "'," & _
                 "  'observationMotorized': '" & observationMotorized & "'"
    If argAdicionales.Count(1) - 1 >= 0 Then
'        sendString = sendString & "," & _
'                     "  'flagReserveCapacity': " & argAdicionales(0, 0) & "," & _
'                     "  'startHour': '" & argAdicionales(0, 1) & "'," & _
'                     "  'endHour': '" & argAdicionales(0, 2) & "'"
        sendString = sendString & "," & _
                     "  'flagReserveCapacity': " & argAdicionales(0, 0) & "," & _
                     "  'startHour': " & procString(argAdicionales(0, 1)) & "," & _
                     "  'endHour': " & procString(argAdicionales(0, 2)) & ""
    End If
    sendString = sendString & "}"
    
    Set objResponse = SendRequest("POST", host, URL, sendString, , "asignaDigital")
    Set asignaDigital = objResponse
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebService.asignaDigital", Err.Description
End Function
Public Function cancelaDigital(ByVal order As String, Optional orderCode As String, Optional orderObs As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim sendString As String
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS115")
    host = GetDomainHost(URL)
    URL = Replace(URL, ":order", order)
    sendString = "{" & _
                 "  'action': 'CANCEL_ORDER'," & _
                 "  'orderCancelCode':'" & orderCode & "'," & _
                 "  'orderCancelObservation':'" & orderObs & "'" & _
                 "}"
    Set objResponse = SendRequest("PATCH", host, URL, sendString, , "cancelaDigital")
    Set cancelaDigital = objResponse
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebService.cancelaDigital", Err.Description
End Function
Public Function getInfoGoogleMaps(ByVal coords As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim sendString As String
    Dim URL As String
    Dim host As String
    'I.ECASTILLO 30.06.2021
    Dim params As New XArrayDB
    Dim values As New XArrayDB
    Dim errors As New XArrayDB
    Dim i As Integer
    'F.ECASTILLO 30.06.2021
    URL = GetUrlWS("URLWS116")
    host = GetDomainHost(URL)
    URL = Replace(URL, ":coords", coords)
    'I.ECASTILLO 30.06.2021
    
    params.ReDim 0, -1, 0, 9
    values.ReDim 0, -1, 0, 9
    errors.ReDim 0, -1, 0, 9
    i = 0
    params.AppendRows
    values.AppendRows
    
    params(i, 0) = "" & "Function"
    values(i, 0) = "" & "getInfoGoogleMaps"
    params(i, 1) = "" & "URL"
    values(i, 1) = "" & URL
    'F.ECASTILLO 30.06.2021
    Set objResponse = SendRequest("POST", host, URL, , , "getInfoGoogleMaps")
    Set getInfoGoogleMaps = objResponse
    'I.ECASTILLO 30.06.2021
    params(i, 2) = "" & "Response"
    values(i, 2) = "" & "getInfoGoogleMaps"
    
    createLog gstrFlagLogFile3, gstrFlagLogBD3, "log", "dlv_unificado_gmaps", params, values, errors
    'F.ECASTILLO 30.06.2021
    Exit Function
Err:
    'I.ECASTILLO 30.06.2021
    params(i, 2) = "" & "Error"
    values(i, 2) = "" & Err.Number & " | " & Err.Description
    
    createLog gstrFlagLogFile3, gstrFlagLogBD3, "log", "dlv_unificado_gmaps", params, values, errors, "1"
    'F.ECASTILLO 30.06.2021
    Err.Raise Err.Number, "clsWebService.getInfoGoogleMaps", Err.Description
End Function
'F.ECASTILLO 27.10.2020
Public Function asignaDigitalSAP(ByVal proformaNumber As String, ByVal localCode As String, ByVal ciaCode As String, ByVal localCodeRef As String, ByVal dateCreated As String, ByVal dateConfirmada As String, ByVal processStartConfirm As String, ByVal Source As String, ByVal companyCode As String, _
                              ByVal creditCardProviderId As String, ByVal currencyId As String, ByVal discountApplied As String, ByVal deliveryType As String, ByVal totalTax As String, ByVal amount As String, ByVal flgCappa As String, ByVal user_dni As String, ByVal user_isInkaClub As String, _
                              ByVal user_phone As String, ByVal user_email As String, ByVal user_name As String, ByVal user_lastName As String, ByVal user_isAnonymous As String, ByVal paymentMethod_id As String, ByVal paymentMethod_name As String, _
                              ByVal receipt_ruc As String, ByVal receipt_companyName As String, ByVal receipt_companyAddress As String, ByVal receipt_receiptType_name As String, ByVal argDelivery As XArrayDB, _
                              ByVal paymentAmountDto_productsTotalCostNoDiscount As String, ByVal paymentAmountDto_grossPrice As String, ByVal paymentAmountDto_discountApplied As String, ByVal paymentAmountDto_productsTotalCost As String, ByVal paymentAmountDto_deliveryCost As String, ByVal paymentAmountDto_paidAmountByCoin As String, _
                              ByVal productos As XArrayDB, ByVal vales As XArrayDB, _
                              ByVal agreementFlag As String, ByVal agreementCode As String, ByVal agreementCodCli As String, ByVal agreementBeneficiaryPct As String, ByVal agreementFields As XArrayDB, ByVal dataCli As String, ByVal nameOf As String, _
                              ByVal observation As String, ByVal observationMotorized As String, ByVal flagMultiBrand As String, ByVal deliveryTime As String, ByVal dateConfirmadaEnd As String, ByVal argAdicionales As XArrayDB) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim sendString As String
    Dim sendItems As String
    Dim sendVales As String
    Dim itemsConv As String
    Dim x, i, j As Integer
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS118")
    host = GetDomainHost(URL)
    
    Dim flg_2e_reserva
    flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
    
    
    For x = 0 To vales.Count(1) - 1
        sendVales = sendVales & "{" & _
                                "   'paymentMethod': '" & vales(x, 0) & "'," & _
                                "   'voucherNumber': '" & vales(x, 1) & "'" & _
                                "},"
    Next x
    If Len(sendVales) > 0 Then sendVales = left(sendVales, Len(sendVales) - 1)
    For x = 0 To productos.Count(1) - 1
         sendItems = sendItems & "{" & _
                                "   'productId': '" & productos(x, 0) & "'," & _
                                "   'productPosuId': '" & productos(x, 1) & "'," & _
                                "   'productSapCode': '" & productos(x, 2) & "'," & _
                                "   'name': '" & productos(x, 3) & "'," & _
                                "   'shortDescription': '" & productos(x, 4) & "'," & _
                                "   'brand': '" & productos(x, 5) & "'," & _
                                "   'quantity': '" & productos(x, 6) & "'," & _
                                "   'quantityUnits': '" & productos(x, 7) & "'," & _
                                "   'fractionated': '" & productos(x, 8) & "'," & _
                                "   'unitPrice': '" & productos(x, 9) & "'," & _
                                "   'totalPrice': '" & productos(x, 10) & "'," & _
                                "   'presentation': '" & productos(x, 11) & "'," & _
                                "   'familyType': '" & productos(x, 12) & "'," & _
                                "   'eanCode': '" & productos(x, 13) & "'," & _
                                "   'fractionatedPrice': '" & productos(x, 14) & "'," & _
                                "   'presentationType': '" & productos(x, 15) & "'," & _
                                "   'valFrac': '" & productos(x, 16) & "'," & _
                                "   'valorUmv': '" & productos(x, 17) & "'," & _
                                "   'quantityTotal': '" & productos(x, 18) & "'" & _
                                "},"
    Next x
    If Len(sendItems) > 0 Then sendItems = left(sendItems, Len(sendItems) - 1)
    If IsObject(agreementFields) = False Then GoTo saltaAgreement
    If agreementFields Is Nothing Then GoTo saltaAgreement
    For x = 0 To agreementFields.Count(1) - 1
        itemsConv = itemsConv & "{ " & _
                 "              'codGroupCia': '" & agreementFields(x, 0) & "'," & _
                 "              'codLocal': '" & agreementFields(x, 1) & "'," & _
                 "              'numPedVta': '" & agreementFields(x, 2) & "'," & _
                 "              'codField': '" & agreementFields(x, 3) & "'," & _
                 "              'codAgreement': '" & agreementFields(x, 4) & "'," & _
                 "              'codClient': '" & agreementFields(x, 5) & "'," & _
                 "              'usuCreaPedVtaCli': '" & agreementFields(x, 6) & "'," & _
                 "              'fecModPedVtaCli': '" & agreementFields(x, 7) & "'," & _
                 "              'usuModPedVtaCli': '" & agreementFields(x, 8) & "'," & _
                 "              'descriptionField': '" & agreementFields(x, 9) & "'," & _
                 "              'nameField': '" & agreementFields(x, 10) & "'," & _
                 "              'flgPrint': '" & agreementFields(x, 11) & "'," & _
                 "              'codValueIn': '" & agreementFields(x, 12) & "'" & _
                 "          },"
    Next x
    If Len(itemsConv) > 0 Then itemsConv = left(itemsConv, Len(itemsConv) - 1)
saltaAgreement:
    sendString = "{ 'proformaNumber': '" & proformaNumber & "'," & _
                 "  'localCode': '" & localCode & "'," & _
                 "  'ciaCode': '" & ciaCode & "'," & _
                 "  'localCodeRef': '" & localCodeRef & "'," & _
                 "  'dateCreated': '" & dateCreated & "'," & _
                 "  'dateConfirmada': '" & dateConfirmada & "'," & _
                 "  'deliveryTime': '" & deliveryTime & "'," & _
                 "  'dateConfirmadaEnd': '" & dateConfirmadaEnd & "'," & _
                 "  'source': '" & Source & "'," & _
                 "  'companyCode': '" & companyCode & "'," & _
                 "  'creditCardProviderId': '" & creditCardProviderId & "'," & _
                 "  'currencyId': '" & currencyId & "'," & _
                 "  'discountApplied': '" & discountApplied & "'," & _
                 "  'deliveryType': '" & deliveryType & "'," & _
                 "  'totalTax': '" & totalTax & "'," & _
                 "  'amount': '" & amount & "'," & _
                 "  'flagDC':" & flgCappa & "," & _
                 "  'flagMultiBrand':" & flagMultiBrand & ","
    sendString = sendString & "  'user': {" & _
                 "      'dni': '" & user_dni & "'," & _
                 "      'isInkaClub': '" & user_isInkaClub & "'," & _
                 "      'phone': '" & user_phone & "'," & _
                 "      'email': '" & user_email & "'," & _
                 "      'name': '" & user_name & "'," & _
                 "      'lastName': '" & user_lastName & "'," & _
                 "      'isAnonymous': '" & user_isAnonymous & "'" & _
                 "  },"
    sendString = sendString & "  'paymentMethod': {" & _
                 "      'id': '" & paymentMethod_id & "'," & _
                 "      'name': '" & paymentMethod_name & "'" & _
                 "  },"
    sendString = sendString & "  'receipt': {" & _
                 "      'ruc' : '" & receipt_ruc & "'," & _
                 "      'companyName' : '" & receipt_companyName & "'," & _
                 "      'companyAddress' : '" & receipt_companyAddress & "'," & _
                 "      'receiptType': {" & _
                 "          'name': '" & receipt_receiptType_name & "'" & _
                 "      }" & _
                 "  },"
                 
    If argDelivery.Count(1) - 1 >= 0 Then
        sendString = sendString & "  'deliveryAddress': {" & _
                 "      'street': '" & argDelivery(0, 0) & "'," & _
                 "      'number': '" & argDelivery(0, 1) & "'," & _
                 "      'apartment': '" & argDelivery(0, 2) & "'," & _
                 "      'country': '" & argDelivery(0, 3) & "'," & _
                 "      'city': '" & argDelivery(0, 4) & "'," & _
                 "      'district': '" & argDelivery(0, 5) & "'," & _
                 "      'latitude': '" & argDelivery(0, 6) & "'," & _
                 "      'longitude': '" & argDelivery(0, 7) & "'," & _
                 "      'receiverName': '" & argDelivery(0, 8) & "'"
        If gstrIndRAv4 = "1" Then
        sendString = sendString & "," & _
                 "      'notes': '" & argDelivery(0, 9) & "'," & _
                 "      'urbanization': '" & argDelivery(0, 10) & "'"
        End If
        sendString = sendString & "  },"
    End If
    
    sendString = sendString & "  'paymentAmountDto': {" & _
                 "      'productsTotalCostNoDiscount': '" & paymentAmountDto_productsTotalCostNoDiscount & "'," & _
                 "      'grossPrice': '" & paymentAmountDto_grossPrice & "'," & _
                 "      'discountApplied': '" & paymentAmountDto_discountApplied & "'," & _
                 "      'productsTotalCost': '" & paymentAmountDto_productsTotalCost & "'," & _
                 "      'deliveryCost': '" & paymentAmountDto_deliveryCost & "'," & _
                 "      'paidAmountByCoin': '" & paymentAmountDto_paidAmountByCoin & "'" & _
                 "  }," & _
                 " 'items': [" & _
                        sendItems & _
                 "  ],"
    If flg_2e_reserva = "0" Then
    Else
        sendString = sendString & "'agreementFlag': '" & agreementFlag & "'," & _
                     "      'agreementCode': '" & agreementCode & "'," & _
                     "      'agreementCodCli': '" & agreementCodCli & "'," & _
                     "      'agreementBeneficiaryPct': '" & agreementBeneficiaryPct & "', " & _
                     "      'agreementFields': [" & _
                                itemsConv & _
                     "      ]," & _
                     "     'dataCli': '" & dataCli & "'," & _
                     "     'nameOf': '" & nameOf & "',"
    End If
    sendString = sendString & _
                 "  'vouchers': [" & _
                        sendVales & _
                 "  ]," & _
                 "  'observation': '" & observation & "'," & _
                 "  'observationMotorized': '" & observationMotorized & "'"
    If argAdicionales.Count(1) - 1 >= 0 Then
'        sendString = sendString & "," & _
'                     "  'flagReserveCapacity': " & argAdicionales(0, 0) & "," & _
'                     "  'startHour': '" & argAdicionales(0, 1) & "'," & _
'                     "  'endHour': '" & argAdicionales(0, 2) & "'"
        sendString = sendString & "," & _
                     "  'flagReserveCapacity': " & argAdicionales(0, 0) & "," & _
                     "  'startHour': " & procString(argAdicionales(0, 1)) & "," & _
                     "  'endHour': " & procString(argAdicionales(0, 2)) & ""
    End If
    sendString = sendString & "}"
        
    Set objResponse = SendRequest("POST", host, URL, sendString, , "asignaDigitalSAP")
    Set asignaDigitalSAP = objResponse
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebService.asignaDigitalSAP", Err.Description
End Function

'I.ECASTILLO 17.12.2020
Public Function listaMetodosSegmentos_Ubigeo(ByVal Ubigeo As String, Optional canal As String, Optional Cia As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim sendString As String
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS120")
    host = GetDomainHost(URL)
    URL = Replace(URL, ":ubigeo", Ubigeo)
    URL = Replace(URL, ":channel", canal)
    URL = Replace(URL, ":cia", Cia)
    URL = URL & "&cb=" & Timer() * 100
    Set objResponse = SendRequest("GET", host, URL, , , "listaMetodosSegmentos_Ubigeo")
    Set listaMetodosSegmentos_Ubigeo = objResponse
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebservice.listaMetodosSegmentos_Ubigeo", Err.Description
End Function
Public Function listaMetodosSegmentos_Coords(ByVal Lat As String, ByVal Lng As String, Optional canal As String, Optional Cia As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim sendString As String
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS121")
    host = GetDomainHost(URL)
    URL = Replace(URL, ":lat", Lat)
    URL = Replace(URL, ":long", Lng)
    URL = Replace(URL, ":channel", canal)
    URL = Replace(URL, ":cia", Cia)
    URL = URL & "&cb=" & Timer() * 100
    Set objResponse = SendRequest("GET", host, URL, , , "listaMetodosSegmentos_Coords")
    Set listaMetodosSegmentos_Coords = objResponse
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebservice.listaMetodosSegmentos_Coords", Err.Description
End Function
Public Function listaMetodosSegmentos_Local(ByVal CodLocal As String, Optional canal As String, Optional Cia As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim sendString As String
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS122")
    host = GetDomainHost(URL)
    URL = Replace(URL, ":codLocal", CodLocal)
    URL = Replace(URL, ":channel", canal)
    URL = Replace(URL, ":cia", Cia)
    URL = URL & "&cb=" & Timer() * 100
    Set objResponse = SendRequest("GET", host, URL, , , "listaMetodosSegmentos_Local")
    Set listaMetodosSegmentos_Local = objResponse
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebservice.listaMetodosSegmentos_Local", Err.Description
End Function

Public Function listaMetodosSegmentosMF_Ubigeo(ByVal Ubigeo As String, Optional canal As String, Optional Cia As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim sendString As String
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS123")
    host = GetDomainHost(URL)
    URL = Replace(URL, ":ubigeo", Ubigeo)
    URL = Replace(URL, ":channel", canal)
    URL = Replace(URL, ":cia", Cia)
    URL = URL & "&cb=" & Timer() * 100
    Set objResponse = SendRequest("GET", host, URL, , , "listaMetodosSegmentosMF_Ubigeo")
    Set listaMetodosSegmentosMF_Ubigeo = objResponse
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebservice.listaMetodosSegmentosMF_Ubigeo", Err.Description
End Function
Public Function listaMetodosSegmentosMF_Coords(ByVal Lat As String, ByVal Lng As String, Optional canal As String, Optional Cia As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim sendString As String
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS124")
    host = GetDomainHost(URL)
    URL = Replace(URL, ":lat", Lat)
    URL = Replace(URL, ":long", Lng)
    URL = Replace(URL, ":channel", canal)
    URL = Replace(URL, ":cia", Cia)
    URL = URL & "&cb=" & Timer() * 100
    Set objResponse = SendRequest("GET", host, URL, , , "listaMetodosSegmentosMF_Coords")
    Set listaMetodosSegmentosMF_Coords = objResponse
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebservice.listaMetodosSegmentosMF_Coords", Err.Description
End Function
Public Function listaMetodosSegmentosMF_Local(ByVal CodLocal As String, Optional canal As String, Optional Cia As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim sendString As String
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS125")
    host = GetDomainHost(URL)
    URL = Replace(URL, ":codLocal", CodLocal)
    URL = Replace(URL, ":channel", canal)
    URL = Replace(URL, ":cia", Cia)
    URL = URL & "&cb=" & Timer() * 100
    Set objResponse = SendRequest("GET", host, URL, , , "listaMetodosSegmentosMF_Local")
    Set listaMetodosSegmentosMF_Local = objResponse
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebservice.listaMetodosSegmentosMF_Local", Err.Description
End Function
'F.ECASTILLO 17.12.2020

Public Function validaSegmento(ByVal dia As String, ByVal hora As String, ByVal CodLocal As String, _
                               ByVal Tipo As String, ByVal canal As String, ByVal Valor As String, _
                               ByVal valorFin As String) As Dictionary
On Error GoTo Err
    Dim objResponse As Dictionary
    Dim sendString As String
    Dim URL As String
    Dim host As String
    URL = GetUrlWS("URLWS117")
    host = GetDomainHost(URL)
    URL = Replace(URL, ":day", dia)
    URL = Replace(URL, ":hour", hora)
    URL = Replace(URL, ":localCode", CodLocal)
    URL = Replace(URL, ":type", Tipo)
    URL = Replace(URL, ":channel", canal)
    URL = Replace(URL, ":value", Valor)
    URL = Replace(URL, ":valueEnd", valorFin)
    URL = URL & "&cb=" & Timer() * 100
    Set objResponse = SendRequest("GET", host, URL, , , "validaSegmento")
    Set validaSegmento = objResponse
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebservice.validaSegmento", Err.Description
End Function

'******************************************************************************************
Private Function GetUrlWS(ByVal Opcion As String) As String
    Dim URL As String
    URL = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", Opcion, objUsuario.CodigoEmpresa)
    
    If URL = "" Then
        Err.Raise vbObjectError, "clsWebService.GetUrlWS", "No se encontro la url del Webservice (" & Opcion & ") para la CIA " & objUsuario.CodigoEmpresa
    End If
    GetUrlWS = URL
End Function

Private Function GetKeyUrlWS(ByVal Opcion As String) As String
    Dim Key As String
    Key = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", Opcion, objUsuario.CodigoEmpresa)
    
    If Key = "" Then
        Err.Raise vbObjectError, "clsWebService.GetKeyUrlWS", "No se encontro la llave del Webservice (" & Opcion & ") para la CIA " & objUsuario.CodigoEmpresa
    End If
    GetKeyUrlWS = Key
End Function

Public Function GetDomainHost(ByVal address As String) As String
On Error GoTo Err
    Dim strOutput       As String
    Dim strTemp         As String
    Dim lngLoopCount    As Long
    Dim lngBCount       As Long
    Dim lngCharCount    As Long
    Dim strProtocol     As String
 
    strOutput$ = Replace(address, "\", "/")
    lngCharCount = Len(strOutput)
 
    If (InStrB(1, strOutput, "/")) Then
        Do Until ((strTemp = "/") Or (lngLoopCount = lngCharCount))
            lngLoopCount = lngLoopCount + 1
            strTemp = Mid$(strOutput, lngBCount + 1, 1)
            strProtocol = strProtocol & strTemp
            lngBCount = lngBCount + 1
        Loop
    End If
 
    strOutput = right$(strOutput, Len(strOutput) - lngBCount)
    lngBCount = 0
    strTemp = "/"
 
    If (InStrB(1, strOutput, "/")) Then
        Do Until strTemp <> "/"
            strTemp = Mid$(strOutput, lngBCount + 1, 1)
            If strTemp = "/" Then lngBCount = lngBCount + 1: strProtocol = strProtocol & strTemp
        Loop
    End If
 
    strOutput = right$(strOutput, Len(strOutput) - lngBCount)
    strOutput = left$(strOutput, InStr(1, strOutput, "/", vbTextCompare) - 1)
    strOutput = strProtocol & strOutput
    GetDomainHost = strOutput
    Exit Function
Err:
    Err.Raise Err.Number, "clsWebservice.GetDomainHost", Err.Description
End Function

Sub createLog(ByVal flgFile As String, ByVal flgBD As String, ByVal strExtention As String, ByVal strNameFile As String, ByVal params As XArrayDB, ByVal values As XArrayDB, Optional errores As XArrayDB, Optional isError As String)
On Error GoTo Err
    Dim strPath As String
    Dim strExtentionProc As String
    Dim strNameFileBK As String
    Dim strDateTime As String
    Dim strDateNow As String
    Dim myStamp As Date
    Dim dateDiff As String
    Dim x As Integer
    Dim i As Integer
    Dim printDateUser As Integer
    Dim strAllData As String
    If flgFile = "1" Then
    Else
        'Exit Sub
        If gstrIndCreaLogError = "1" And isError = "1" Then
        Else
            GoTo saveBD
        End If
    End If
    strNameFileBK = strNameFile
    strPath = gstrPathLog '"" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "PATHDLVLOG", "10") '"C:\Documents and Settings\aescate\Escritorio\ecastillo\demo\"
    'strExtention = "log"
    strDateNow = DateTime.Now
    If Len(Trim(strPath)) = 0 Then strPath = App.Path & "\log\"
    strDateTime = gclsOracle.Fecha_Servidor
    strDateTime = Replace(strDateTime, "/", "-")
    strNameFile = strNameFile & "_" & Format(strDateNow, "dd-mm-yyyy_hh") 'strDateTime
    Dim File As String
    Dim cntFiles As Integer
    If right$(strPath, 1) <> "\" Then strPath = strPath & "\"
    If Trim$(strExtention) = "" Then
        strExtention = "*.*"
    ElseIf left$(strExtention, 2) <> "*." Then
        strExtention = "*." & strExtention
    End If
    strExtentionProc = strNameFileBK & strExtention
    File = Dir$(strPath & strExtentionProc)
    cntFiles = 0
    Do While Len(File)
        cntFiles = cntFiles + 1
        File = Dir$
    Loop
    Dim fileOld As String
    Dim fileOldStamp As Date
    File = Dir$(strPath & strExtentionProc)
    Do While Len(File)
        If cntFiles > gintFilesLog Then
            myStamp = FileDateTime(strPath & File)
            dateDiff = DateTime.dateDiff("h", myStamp, IIf(fileOldStamp = "12:00:00 a.m.", strDateNow, fileOldStamp))
            If dateDiff > 0 Then
                fileOld = strPath & File
                fileOldStamp = myStamp
            End If
        End If
        'myStamp = FileDateTime(strPath & File)
        'dateDiff = DateTime.dateDiff("d", myStamp, strDateTime)
        'If dateDiff > 2 Then
        '    Kill (strPath & File)
        'End If
        File = Dir$
    Loop
    If Len(Trim(fileOld)) <> 0 Then Kill (fileOld)
    If left$(strExtention, 2) = "*." Then strExtention = "." & right$(strExtention, 3)
Reintenta:
    printDateUser = 0: x = 0: i = 0
    Open strPath & strNameFile & strExtention For Append As #1
gtprintDateUser:
        If printDateUser = 1 Then
            Print #1, "Date: " & strDateNow
            Print #1, "User: " & objUsuario.Codigo
        End If
        For x = 0 To params.Count(1) - 1
         For i = 0 To params.Count(2) - 1
            If Len(Trim(params(x, i))) > 0 Then
                If printDateUser = 0 Then printDateUser = 1: GoTo gtprintDateUser
                Print #1, params(x, i) & ": " & values(x, i)
                If Not IsMissing(errores) Then
                    If errores.Count(1) > 0 Then
                       If Len(Trim(errores(x, i))) > 0 Then
                           Print #1, "Error: " & errores(x, i)
                       End If
                    End If
                End If
            End If
         Next i
        Next x
        Print #1, "=============================================================================================="
    Close #1
saveBD:
    If flgBD = "1" Then
    Else
        If gstrIndCreaLogError = "1" And isError = "1" Then
        Else
            Exit Sub
        End If
    End If
    printDateUser = 0: x = 0: i = 0
gtprintDateUser2:
    If printDateUser = 1 Then
        strAllData = "Date: " & strDateNow & vbNewLine & "User: " & objUsuario.Codigo & vbNewLine
    End If
    For x = 0 To params.Count(1) - 1
        For i = 0 To params.Count(2) - 1
           If Len(Trim(params(x, i))) > 0 Then
               If printDateUser = 0 Then printDateUser = 1: GoTo gtprintDateUser2
               strAllData = strAllData & params(x, i) & ": " & values(x, i) & vbNewLine
               If Not IsMissing(errores) Then
                   If errores.Count(1) > 0 Then
                      If Len(Trim(errores(x, i))) > 0 Then
                          strAllData = strAllData & "Error: " & errores(x, i) & vbNewLine
                      End If
                   End If
               End If
           End If
        Next i
    Next x
    grabaLogDelivery strAllData
    Exit Sub
Err:
    If Err.Number = 55 Then Close #1: GoTo Reintenta
    MsgBox Err.Number & ": " & Err.Description & " : abc"
End Sub
'I.ECASTILLO 05.07.2021
Public Function DecodeUTF8(s)
  Dim i
  Dim c
  Dim n
  
  s = s & " "

  i = 1
  Do While i <= Len(s)
    c = Asc(Mid(s, i, 1))
    If c And &H80 Then
      n = 1
      Do While i + n < Len(s)
        If (Asc(Mid(s, i + n, 1)) And &HC0) <> &H80 Then
          Exit Do
        End If
        n = n + 1
      Loop
      If n = 2 And ((c And &HE0) = &HC0) Then
        c = Asc(Mid(s, i + 1, 1)) + &H40 * (c And &H1)
      Else
        c = 191
      End If
      s = left(s, i - 1) + Chr(c) + Mid(s, i + n)
    End If
    i = i + 1
  Loop
  DecodeUTF8 = s
End Function
'F.ECASTILLO
Public Function DecodeURI(ByVal EncodedURI As String) As String
    Dim bANSI() As Byte
    Dim bUTF8() As Byte
    Dim lIndex As Long
    Dim lUTFIndex As Long

    If Len(EncodedURI) = 0 Then
        Exit Function
    End If

EncodedURI = Replace$(EncodedURI, "+", " ")         ' In case encoding isn't used.
    bANSI = strConv(EncodedURI, vbFromUnicode)          ' Convert from unicode text to ANSI values
    ReDim bUTF8(UBound(bANSI))                          ' Declare dynamic array, get length
    For lIndex = 0 To UBound(bANSI)                     ' from 0 to length of ANSI
        If bANSI(lIndex) = &H25 Then                    ' If we have ASCII 37, %, then
            bUTF8(lUTFIndex) = val("&H" & Mid$(EncodedURI, lIndex + 2, 2)) ' convert hex to ANSI
            lIndex = lIndex + 2                         ' this character was encoded into two bytes
        Else
            bUTF8(lUTFIndex) = bANSI(lIndex)            ' otherwise don't need to do anything special
        End If
        lUTFIndex = lUTFIndex + 1                       ' advance utf index
    Next
    DecodeURI = FromUTF8(bUTF8, lUTFIndex)              ' convert to string
End Function

Public Function FromUTF8(ByRef utf8() As Byte, ByVal Length As Long) As String
    Dim lDataLength As Long

    lDataLength = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), Length, 0, 0)  ' Get the length of the data.
    FromUTF8 = String$(lDataLength, 0)                                         ' Create array big enough
    MultiByteToWideChar CP_UTF8, 0, VarPtr(utf8(0)), _
                        Length, StrPtr(FromUTF8), lDataLength                  '
End Function

Public Function UTF8UrlEnc(s As String) As String
Dim i As Long, w As Integer
  For i = 1 To Len(s)
    w = AscW(Mid$(s, i, 1))
    If w < 128 Then
      UTF8UrlEnc = UTF8UrlEnc & ChrW$(w)
    ElseIf w < 2048 Then
      UTF8UrlEnc = UTF8UrlEnc & "%" & Hex$(w \ 64 Or 192) & "%" & Hex$(w And 63 Or 128)
    End If
  Next i
End Function

Public Function grabaLogDelivery(ByVal vStr As String)
   On Error GoTo CtrlErr
   
   Dim gvarValores As Variant
   Dim gvarIO  As Variant
   
   gvarValores = Array(objUsuario.Codigo, Obtener(12), vStr)
                          
   gvarIO = Array(entrada, entrada, entrada)
                  
   grabaLogDelivery = gclsOracle.SP("BTLPROD.PKG_LOGS_RAC.SP_LOG_DELIVERY", _
                                    gvarValores, _
                                    gvarIO)
                  
   Exit Function
CtrlErr:
   Err.Raise Err.Number, "clsWebServices.grabaLogDelivery", Err.Description
End Function

Function FileText() As String
    Dim filename As String
    filename = "responseExample.txt"
    Dim handle As Integer
     
       ' ensure that the file exists
    If Len(Dir$(filename)) = 0 Then
        Err.Raise 53  ' File not found
    End If
     
       ' open in binary mode
    handle = FreeFile
    Open filename$ For Binary As #handle
       ' read the string and close the file
    FileText = Space$(LOF(handle))
    Get #handle, , FileText
    Close #handle
End Function

Function procString(s) As String
    'indentificar true,false,null
    'para determinar si concatenar apostrofe a cadena
    Dim idenStr As Variant
    Dim i As Long
    Dim posStr As Long
    Dim Existe As Integer
    Dim response As String
    idenStr = Array("true", "false", "null")
    
    For i = 0 To UBound(idenStr)
        posStr = InStr(1, s, idenStr(i))
        Existe = 0
        If posStr > 0 Then Existe = 1: Exit For
    Next i
    If Existe = 1 Then
        response = s
    Else
        response = "'" & s & "'"
    End If
    procString = response
End Function
