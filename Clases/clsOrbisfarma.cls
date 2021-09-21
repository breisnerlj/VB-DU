VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrbisFarma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales para almacenar los valores de las propiedades
'''Private Const mvarurl As String = "https://mifarmaws01.orbisfarma.com.mx/"
'''Private Const mvarkey As String = "67723f0e27d4869d280cb83bc6ede7d1"
Private mvarurl As String
Private mvarkey As String
Private SoapAction As String
    

Private Function PostWebservice(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal xmlBody As String) As String
    Dim objXMLDoc As MSXML2.DOMDocument60
    Dim objXmlHttp As MSXML2.XMLHTTP60
    Dim faultCode As String, faultText As String, strRet As String
    Dim withFault As Boolean
    
    ' Create objects to DOMDocument and XMLHTTP
    Set objXMLDoc = New MSXML2.DOMDocument60
    Set objXmlHttp = New MSXML2.XMLHTTP60
    
    ' Load XML
    objXMLDoc.async = False
    objXMLDoc.loadXML xmlBody

    ' Open the webservice
    objXmlHttp.Open "POST", AsmxUrl, False
    'objXmlHttp.Open "POST", "http://mifarmaopws01.orbisfarma.com.mx/Operaciones.asmx", False
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapAction & SoapActionUrl 'AsmxUrl & SoapActionUrl
    'objXmlHttp.setRequestHeader "SOAPAction", "https://mifarmaopws01.orbisfarma.com.mx/getAffiliationSearch"
    
    Debug.Print "Content-Type: text/xml; charset=utf-8"
    Debug.Print "SOAPAction: " & SoapAction & SoapActionUrl
    Debug.Print objXMLDoc.XML
    Debug.Print "============================================"
    
    ' Send XML command
    objXmlHttp.send objXMLDoc.XML

    ' Get all response text from webservice
    Debug.Print objXmlHttp.getAllResponseHeaders
    Debug.Print objXmlHttp.responseText
    Debug.Print "============================================"
    strRet = objXmlHttp.responseText
    withFault = False
    
    If InStr(1, strRet, "<soap:Fault>", vbTextCompare) Then
        withFault = True
        With objXMLDoc
            .async = False
            .preserveWhiteSpace = False
            .validateOnParse = False
            .resolveExternals = True
            .setProperty "SelectionNamespaces", "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " & _
                                                "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
                                                "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
                                                "xmlns:orb=""" & mvarurl & """"
    
            .loadXML strRet
            faultCode = .selectSingleNode("/soap:Envelope/soap:Body/soap:Fault/faultcode").Text
            faultText = .selectSingleNode("/soap:Envelope/soap:Body/soap:Fault/faultstring").Text
        End With
    End If
    
    ' Close object
    Set objXmlHttp = Nothing
    Set objXMLDoc = Nothing
    
    If withFault Then
        Err.Raise vbObjectError, faultCode, faultText
    End If
    
    ' Return result
    PostWebservice = strRet
    
End Function

'''Public Function checkSaleReturn(ByVal CardNumber As String, _
'''                                ByVal EmployeeId As String, _
'''                                ByVal TransactionId As String, _
'''                                ByVal SaleAuthNumber As String, _
'''                                ByVal TransactionItems As String) As stCheckSale
'''    Dim oXMLDoc As MSXML2.DOMDocument60 'Object
'''    Dim oRootNode As MSXML2.IXMLDOMNode
'''    Dim oNode As MSXML2.IXMLDOMNode
'''    Dim stTI As stCheckSale
'''    Dim xmlBody As String
'''
'''    ' Se arma la peticion para el WS
'''    xmlBody = _
'''        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
'''        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
'''        "  <soap:Body>" & _
'''        "    <checkSaleReturn xmlns=""@@url@@"">" & _
'''        "      <cardnumber>@@cardnumber@@</cardnumber>" & _
'''        "      <storeid>01</storeid>" & _
'''        "      <posid>01</posid>" & _
'''        "      <employeeid>@@employeeid@@</employeeid>" & _
'''        "      <transactionid>@@transactionid@@</transactionid>" & _
'''        "      <saleauthnumber>@@saleauthnumber@@</saleauthnumber>" & _
'''        "      <transactionitems>@@transactionitems@@</transactionitems>" & _
'''        "      <key>@@key@@</key>" & _
'''        "    </checkSaleReturn>" & _
'''        "  </soap:Body>" & _
'''        "</soap:Envelope>"
'''    xmlBody = Replace(xmlBody, "@@url@@", mvarurl, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@cardnumber@@", CardNumber, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@cardnumber@@", CardNumber, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@employeeid@@", EmployeeId, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@transactionid@@", TransactionId, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@saleauthnumber@@", SaleAuthNumber, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@transactionitems@@", TransactionItems, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@key@@", mvarkey, , , vbTextCompare)
'''
'''    ' Se recibe la respuesta del WS
'''    xmlBody = PostWebservice(mvarurl, "checkSaleReturn", xmlBody)
'''
'''    Set oXMLDoc = New MSXML2.DOMDocument60
'''    Set stTI = New stCheckSale
'''
'''    With oXMLDoc
'''        .async = False
'''        .preserveWhiteSpace = False
'''        .validateOnParse = False
'''        .resolveExternals = True
'''        .setProperty "SelectionNamespaces", "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " & _
'''                                            "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
'''                                            "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
'''                                            "xmlns:orb=""" & mvarurl & """"
'''        If .loadXML(xmlBody) Then
'''            Set oRootNode = .selectSingleNode("/soap:Envelope/soap:Body/orb:checkSaleReturnResponse/orb:checkSaleReturnResult/orb:stcheckSaleReturn")
'''            For Each oNode In oRootNode.childNodes
'''                CallByName stTI, oNode.nodeName, VbLet, CStr("" & oNode.Text)
'''            Next oNode
'''        End If
'''    End With
'''    Set oXMLDoc = Nothing
'''
'''    Set checkSaleReturn = stTI
'''
'''End Function

Public Function setTransactionInit(ByVal cardnumber As String, _
                                   ByVal EmployeeId As String) As stTransactionInit
    Dim oXMLDoc As MSXML2.DOMDocument60 'Object
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim stTI As stTransactionInit
    Dim xmlBody As String
    
    ' Se arma la peticion para el WS
    xmlBody = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "  <soap:Body>" & _
        "    <setTransactionInit xmlns=""@@url@@"">" & _
        "      <cardnumber>@@cardnumber@@</cardnumber>" & _
        "      <storeid>01</storeid>" & _
        "      <posid>01</posid>" & _
        "      <employeeid>@@employeeid@@</employeeid>" & _
        "      <key>@@key@@</key>" & _
        "    </setTransactionInit>" & _
        "  </soap:Body>" & _
        "</soap:Envelope>"
    xmlBody = Replace(xmlBody, "@@url@@", mvarurl, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@cardnumber@@", cardnumber, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@employeeid@@", EmployeeId, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@key@@", mvarkey, , , vbTextCompare)
    
    ' Se recibe la respuesta del WS
    xmlBody = PostWebservice(mvarurl, "setTransactionInit", xmlBody)
    
    Set oXMLDoc = New MSXML2.DOMDocument60
    Set stTI = New stTransactionInit
    
    With oXMLDoc
        .async = False
        .preserveWhiteSpace = False
        .validateOnParse = False
        .resolveExternals = True
        .setProperty "SelectionNamespaces", "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " & _
                                            "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
                                            "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
                                            "xmlns:orb=""" & mvarurl & """"
        If .loadXML(xmlBody) Then
            Set oRootNode = .selectSingleNode("/soap:Envelope/soap:Body/orb:setTransactionInitResponse/orb:setTransactionInitResult/orb:stTransactionInit")
            For Each oNode In oRootNode.childNodes
                CallByName stTI, oNode.nodeName, VbLet, CStr("" & oNode.Text)
            Next oNode
        End If
    End With
    Set oXMLDoc = Nothing
    
    Set setTransactionInit = stTI
    
End Function

Public Function getAffiliationSearch(ByVal accountnumber As String, _
                                     ByVal EmployeeId As String) As stAffiliationSearch
    Dim oXMLDoc As MSXML2.DOMDocument60 'Object
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim stTI As stAffiliationSearch
    Dim xmlBody As String
    ' Se arma la peticion para el WS
'    xmlBody = _
'    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
'    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
'    "  <soap:Body>" & _
'    "    <getAffiliationSearch xmlns=""@@url@@"">" & _
'    "      <accountnumber>@@accountnumber@@</accountnumber>" & _
'    "      <storeid>01</storeid>" & _
'    "      <posid>01</posid>" & _
'    "      <employeeid>@@employeeid@@</employeeid>" & _
'    "      <key>@@key@@</key>" & _
'    "    </getAffiliationSearch>" & _
'    "  </soap:Body>" & _
'    "</soap:Envelope>"
    '****************************************'
    xmlBody = _
    "<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:mif=""@@soapaction@@"">" & _
    "  <soap:Header/>" & _
    "  <soap:Body>" & _
    "    <mif:getAffiliationSearch>" & _
    "        <mif:accountnumber>@@accountnumber@@</mif:accountnumber>" & _
    "        <mif:storeid>01</mif:storeid>" & _
    "        <mif:posid>01</mif:posid>" & _
    "        <mif:employeeid>@@employeeid@@</mif:employeeid>" & _
    "        <mif:key>@@key@@</mif:key>" & _
    "    </mif:getAffiliationSearch>" & _
    "  </soap:Body>" & _
    "</soap:Envelope>"
    'xmlBody = Replace(xmlBody, "@@url@@", mvarurl, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@soapaction@@", SoapAction, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@employeeid@@", EmployeeId, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@accountnumber@@", accountnumber, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@key@@", mvarkey, , , vbTextCompare)
    
    ' Se recibe la respuesta del WS
    Debug.Print xmlBody
    Debug.Print "=========================="
    xmlBody = PostWebservice(mvarurl, "getAffiliationSearch", xmlBody)
    
    Set oXMLDoc = New MSXML2.DOMDocument60
    Set stTI = New stAffiliationSearch
    
    With oXMLDoc
        .async = False
        .preserveWhiteSpace = False
        .validateOnParse = False
        .resolveExternals = True
        .setProperty "SelectionNamespaces", "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " & _
                                            "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
                                            "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
                                            "xmlns:orb=""" & mvarurl & """"
        If .loadXML(xmlBody) Then
            'Set oRootNode = .selectSingleNode("/soap:Envelope/soap:Body/orb:getAffiliationSearchResponse/orb:getAffiliationSearchResult/orb:stAffiliationSearch")
            Set oRootNode = .selectSingleNode("/soap:Envelope/soap:Body").firstChild.firstChild.firstChild '.selectNodes("/getAffiliationSearchResponse/getAffiliationSearchResult/stAffiliationSearch")
            For Each oNode In oRootNode.childNodes
                CallByName stTI, oNode.nodeName, VbLet, CStr("" & oNode.Text)
            Next oNode
        End If
    End With
    Set oXMLDoc = Nothing
    
    Set getAffiliationSearch = stTI
    
End Function

'''Public Function getTransactionStatus(ByVal TransactionId As String) As stTransactionStatus
'''    Dim oXMLDoc As MSXML2.DOMDocument60 'Object
'''    Dim oRootNode As MSXML2.IXMLDOMNode
'''    Dim oNode As MSXML2.IXMLDOMNode
'''    Dim stTI As stTransactionStatus
'''    Dim xmlBody As String
'''
'''    ' Se arma la peticion para el WS
'''    xmlBody = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "  <soap:Body>" & _
        "    <getTransactionStatus xmlns=""" & mvarurl & """>" & _
        "      <transactionid>@@transactionid@@</transactionid>" & _
        "      <storeid>01</storeid>" & _
        "      <posid>01</posid>" & _
        "      <employeeid>@@employeeid@@</employeeid>" & _
        "      <key>@@key@@</key>" & _
        "    </getTransactionStatus>" & _
        "  </soap:Body>" & _
        "</soap:Envelope>"
'''    xmlBody = Replace(xmlBody, "@@transactionid@@", TransactionId, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@key@@", mvarkey, , , vbTextCompare)
'''
'''    ' Se recibe la respuesta del WS
'''    xmlBody = PostWebservice(mvarurl, getTransactionStatus, xmlBody)
'''
'''    Set oXMLDoc = New MSXML2.DOMDocument60
'''    Set stTI = New stTransactionStatus
'''
'''    With oXMLDoc
'''        .async = False
'''        .preserveWhiteSpace = False
'''        .validateOnParse = False
'''        .resolveExternals = True
'''        .setProperty "SelectionNamespaces", "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " & _
'''                                            "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
'''                                            "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
'''                                            "xmlns:orb=""" & mvarurl & """"
'''        If .loadXML(xmlBody) Then
'''            Set oRootNode = .selectSingleNode("/soap:Envelope/soap:Body/orb:getTransactionStatusResponse/orb:getTransactionStatusResult/orb:stTransactionStatus")
'''            For Each oNode In oRootNode.childNodes
'''                CallByName stTI, oNode.nodeName, VbLet, CStr("" & oNode.Text)
'''            Next oNode
'''        End If
'''    End With
'''    Set oXMLDoc = Nothing
'''
'''    Set getTransactionStatus = stTI
'''
'''End Function

Public Function setCardAffiliation(ByVal cardnumber As String, _
                                   ByVal accountnumber As String, _
                                   ByVal cardname As String, _
                                   ByVal cardlastname As String, _
                                   ByVal cardbirthdate As String, _
                                   ByVal cardgender As String, _
                                   ByVal carddoctor As String, _
                                   ByVal cardcontactpreferences As String, _
                                   ByVal cardemail As String, _
                                   ByVal cardphone As String, _
                                   ByVal cardcellphone As String, _
                                   ByVal cardaddress As String, _
                                   ByVal cardfields As String, _
                                   ByVal EmployeeId As String) As stCardAffiliation
    Dim oXMLDoc As MSXML2.DOMDocument60 'Object
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim stTI As stCardAffiliation
    Dim xmlBody As String

    ' Se arma la peticion para el WS
    xmlBody = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "  <soap:Body>" & _
        "    <setCardAffiliation xmlns=""@@url@@"">" & _
        "      <cardnumber>@@cardnumber@@</cardnumber>" & _
        "      <storeid>01</storeid>" & _
        "      <posid>01</posid>" & _
        "      <employeeid>@@employeeid@@</employeeid>" & _
        "      <accountnumber>@@accountnumber@@</accountnumber>" & _
        "      <cardname>@@cardname@@</cardname>" & _
        "      <cardlastname>@@cardlastname@@</cardlastname>" & _
        "      <cardbirthdate>@@cardbirthdate@@</cardbirthdate>" & _
        "      <cardgender>@@cardgender@@</cardgender>" & _
        "      <carddoctor>@@carddoctor@@</carddoctor>" & _
        "      <cardcontactpreferences>@@cardcontactpreferences@@</cardcontactpreferences>" & _
        "      <cardemail>@@cardemail@@</cardemail>" & _
        "      <cardphone>@@cardphone@@</cardphone>" & _
        "      <cardcellphone>@@cardcellphone@@</cardcellphone>" & _
        "      <cardaddress>@@cardaddress@@</cardaddress>" & _
        "      <cardfields>@@cardfields@@</cardfields>" & _
        "      <key>@@key@@</key>" & _
        "    </setCardAffiliation>" & _
        "  </soap:Body>" & _
        "</soap:Envelope>"
    xmlBody = Replace(xmlBody, "@@url@@", mvarurl, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@cardnumber@@", cardnumber, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@employeeid@@", EmployeeId, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@accountnumber@@", accountnumber, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@cardname@@", cardname, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@cardlastname@@", cardlastname, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@cardbirthdate@@", cardbirthdate, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@cardgender@@", cardgender, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@carddoctor@@", carddoctor, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@cardcontactpreferences@@", cardcontactpreferences, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@cardemail@@", cardemail, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@cardphone@@", cardphone, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@cardcellphone@@", cardcellphone, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@cardaddress@@", cardaddress, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@cardfields@@", cardfields, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@key@@", mvarkey, , , vbTextCompare)

    ' Se recibe la respuesta del WS
    xmlBody = PostWebservice(mvarurl, "setCardAffiliation", xmlBody)

    Set oXMLDoc = New MSXML2.DOMDocument60
    Set stTI = New stCardAffiliation

    With oXMLDoc
        .async = False
        .preserveWhiteSpace = False
        .validateOnParse = False
        .resolveExternals = True
        .setProperty "SelectionNamespaces", "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " & _
                                            "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
                                            "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
                                            "xmlns:orb=""" & mvarurl & """"
        If .loadXML(xmlBody) Then
            Set oRootNode = .selectSingleNode("/soap:Envelope/soap:Body/orb:setCardAffiliationResponse/orb:setCardAffiliationResult/orb:stCardAffiliation")
            For Each oNode In oRootNode.childNodes
                CallByName stTI, oNode.nodeName, VbLet, CStr("" & oNode.Text)
            Next oNode
        End If
    End With
    Set oXMLDoc = Nothing

    Set setCardAffiliation = stTI

End Function

'''Public Function setSaleReturn(ByVal CardNumber As String, _
'''                              ByVal EmployeeId As String, _
'''                              ByVal TransactionId As String, _
'''                              ByVal TransactionItems As String, _
'''                              ByVal transactionwithdrawal As String, _
'''                              ByVal InvoiceNumber As String, _
'''                              ByVal InvoiceDate As String, _
'''                              ByVal InvoiceAmount As String) As stSaleReturn
'''    Dim oXMLDoc As MSXML2.DOMDocument60 'Object
'''    Dim oRootNode As MSXML2.IXMLDOMNode
'''    Dim oNode As MSXML2.IXMLDOMNode
'''    Dim stTI As stSaleReturn
'''    Dim xmlBody As String
'''
'''    ' Se arma la peticion para el WS
'''    xmlBody = _
'''        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
'''        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
'''        "  <soap:Body>" & _
'''        "    <setSaleReturn xmlns=""@@url@@"">" & _
'''        "      <cardnumber>@@cardnumber@@</cardnumber>" & _
'''        "      <storeid>01</storeid>" & _
'''        "      <posid>01</posid>" & _
'''        "      <employeeid>@@employeeid@@</employeeid>" & _
'''        "      <transactionid>@@transactionid@@</transactionid>" & _
'''        "      <transactionitems>@@transactionitems@@</transactionitems>" & _
'''        "      <transactionwithdrawal>@@transactionwithdrawal@@</transactionwithdrawal>" & _
'''        "      <invoicenumber>@@invoicenumber@@</invoicenumber>" & _
'''        "      <invoicedate>@@invoicedate@@</invoicedate>" & _
'''        "      <invoiceamount>@@invoiceamount@@</invoiceamount>" & _
'''        "      <key>@@key@@</key>" & _
'''        "    </setSaleReturn>" & _
'''        "  </soap:Body>" & _
'''        "</soap:Envelope>"
'''    xmlBody = Replace(xmlBody, "@@url@@", mvarurl, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@cardnumber@@", CardNumber, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@employeeid@@", EmployeeId, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@transactionid@@", TransactionId, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@transactionitems@@", TransactionItems, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@transactionwithdrawal@@", transactionwithdrawal, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@invoicenumber@@", InvoiceNumber, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@invoicedate@@", InvoiceDate, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@invoiceamount@@", InvoiceAmount, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@key@@", mvarkey, , , vbTextCompare)
'''
'''    ' Se recibe la respuesta del WS
'''    xmlBody = PostWebservice(mvarurl, "setSaleReturn", xmlBody)
'''
'''    Set oXMLDoc = New MSXML2.DOMDocument60
'''    Set stTI = New stSaleReturn
'''
'''    With oXMLDoc
'''        .async = False
'''        .preserveWhiteSpace = False
'''        .validateOnParse = False
'''        .resolveExternals = True
'''        .setProperty "SelectionNamespaces", "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " & _
'''                                            "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
'''                                            "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
'''                                            "xmlns:orb=""" & mvarurl & """"
'''        If .loadXML(xmlBody) Then
'''            Set oRootNode = .selectSingleNode("/soap:Envelope/soap:Body/orb:setSaleReturnResponse/orb:setSaleReturnResult/orb:stSaleReturn")
'''            For Each oNode In oRootNode.childNodes
'''                CallByName stTI, oNode.nodeName, VbLet, CStr("" & oNode.Text)
'''            Next oNode
'''        End If
'''    End With
'''    Set oXMLDoc = Nothing
'''
'''    Set setSaleReturn = stTI
'''
'''End Function

'''Public Function setSaleReverse(ByVal CardNumber As String, _
'''                               ByVal TransactionId As String, _
'''                               ByVal saleauthnumber As String, _
'''                               ByVal InvoiceNumber As String) As stSaleReturn
'''    Dim oXMLDoc As MSXML2.DOMDocument60 'Object
'''    Dim oRootNode As MSXML2.IXMLDOMNode
'''    Dim oNode As MSXML2.IXMLDOMNode
'''    Dim stTI As stSaleReverse
'''    Dim xmlBody As String
'''
'''    ' Se arma la peticion para el WS
'''    xmlBody = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "  <soap:Body>" & _
        "    <setSaleReverse xmlns=""" & mvarurl & """>" & _
        "      <cardnumber>@@cardnumber@@</cardnumber>" & _
        "      <storeid>01</storeid>" & _
        "      <posid>01</posid>" & _
        "      <employeeid>@@employeeid@@</employeeid>" & _
        "      <transactionid>@@transactionid@@</transactionid>" & _
        "      <saleauthnumber>@@saleauthnumber@@</saleauthnumber>" & _
        "      <invoicenumber>@@invoicenumber@@</invoicenumber>" & _
        "      <key>@@key@@</key>" & _
        "    </setSaleReverse>" & _
        "  </soap:Body>" & _
        "</soap:Envelope>"

'''    xmlBody = Replace(xmlBody, "@@cardnumber@@", CardNumber, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@transactionid@@", TransactionId, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@saleauthnumber@@", saleauthnumber, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@invoicenumber@@", InvoiceNumber, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@key@@", mvarkey, , , vbTextCompare)
'''
'''    ' Se recibe la respuesta del WS
'''    xmlBody = PostWebservice(mvarurl, mvarsaSetSaleReverse, xmlBody)
'''
'''    Set oXMLDoc = New MSXML2.DOMDocument60
'''    Set stTI = New stSaleReverse
'''
'''    With oXMLDoc
'''        .async = False
'''        .preserveWhiteSpace = False
'''        .validateOnParse = False
'''        .resolveExternals = True
'''        .setProperty "SelectionNamespaces", "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " & _
'''                                            "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
'''                                            "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
'''                                            "xmlns:orb=""" & mvarurl & """"
'''        If .loadXML(xmlBody) Then
'''            Set oRootNode = .selectSingleNode("/soap:Envelope/soap:Body/orb:setSaleReverseResponse/orb:setSaleReverseResult/orb:setSaleReverse")
'''            For Each oNode In oRootNode.childNodes
'''                CallByName stTI, oNode.nodeName, VbLet, CStr("" & oNode.Text)
'''            Next oNode
'''        End If
'''    End With
'''    Set oXMLDoc = Nothing
'''
'''    Set setSaleReverse = stTI
'''
'''End Function

'''Public Function setTransactionCancel(ByVal CardNumber As String, _
'''                                     ByVal TransactionId As String) As stSaleReturn
'''    Dim oXMLDoc As MSXML2.DOMDocument60 'Object
'''    Dim oRootNode As MSXML2.IXMLDOMNode
'''    Dim oNode As MSXML2.IXMLDOMNode
'''    Dim stTI As stTransactionCancel
'''    Dim xmlBody As String
'''
'''    ' Se arma la peticion para el WS
'''    xmlBody = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "  <soap:Body>" & _
        "    <setTransactionCancel xmlns=""" & mvarurl & """>" & _
        "      <cardnumber>@@cardnumber@@</cardnumber>" & _
        "      <storeid>01</storeid>" & _
        "      <posid>01</posid>" & _
        "      <employeeid>@@employeeid@@</employeeid>" & _
        "      <transactionid>@@transactionid@@</transactionid>" & _
        "      <key>@@key@@</key>" & _
        "    </setTransactionCancel>" & _
        "  </soap:Body>" & _
        "</soap:Envelope>"

'''    xmlBody = Replace(xmlBody, "@@cardnumber@@", CardNumber, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@transactionid@@", TransactionId, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@key@@", mvarkey, , , vbTextCompare)
'''
'''    ' Se recibe la respuesta del WS
'''    xmlBody = PostWebservice(mvarurl, mvarsaSetTransactionCancel, xmlBody)
'''
'''    Set oXMLDoc = New MSXML2.DOMDocument60
'''    Set stTI = New stTransactionCancel
'''
'''    With oXMLDoc
'''        .async = False
'''        .preserveWhiteSpace = False
'''        .validateOnParse = False
'''        .resolveExternals = True
'''        .setProperty "SelectionNamespaces", "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " & _
'''                                            "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
'''                                            "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
'''                                            "xmlns:orb=""" & mvarurl & """"
'''        If .loadXML(xmlBody) Then
'''            Set oRootNode = .selectSingleNode("/soap:Envelope/soap:Body/orb:setTransactionCancelResponse/orb:setTransactionCancelResult/orb:stTransactionCancel")
'''            For Each oNode In oRootNode.childNodes
'''                CallByName stTI, oNode.nodeName, VbLet, CStr("" & oNode.Text)
'''            Next oNode
'''        End If
'''    End With
'''    Set oXMLDoc = Nothing
'''
'''    Set setTransactionCancel = stTI
'''
'''End Function

'''Public Function setTransactionQuote(ByVal CardNumber As String, _
'''                                    ByVal EmployeeId As String, _
'''                                    ByVal TransactionId As String, _
'''                                    ByVal TransactionItems As String) As stSaleReturn
'''    Dim oXMLDoc As MSXML2.DOMDocument60 'Object
'''    Dim oRootNode As MSXML2.IXMLDOMNode
'''    Dim oNode As MSXML2.IXMLDOMNode
'''    Dim stTI As stTransactionQuote
'''    Dim xmlBody As String
'''
'''    ' Se arma la peticion para el WS
'''    xmlBody = _
'''        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
'''        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
'''        "  <soap:Body>" & _
'''        "    <setTransactionQuote xmlns=""@@url@@"">" & _
'''        "      <cardnumber>@@cardnumber@@</cardnumber>" & _
'''        "      <storeid>01</storeid>" & _
'''        "      <posid>01</posid>" & _
'''        "      <employeeid>@@employeeid@@</employeeid>" & _
'''        "      <transactionid>@@transactionid@@</transactionid>" & _
'''        "      <transactionitems>@@transactionitems@@</transactionitems>" & _
'''        "      <key>@@key@@</key>" & _
'''        "    </setTransactionQuote>" & _
'''        "  </soap:Body>" & _
'''        "</soap:Envelope>"
'''    xmlBody = Replace(xmlBody, "@@url@@", mvarurl, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@cardnumber@@", CardNumber, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@employeeid@@", EmployeeId, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@transactionid@@", TransactionId, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@transactionitems@@", TransactionItems, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@key@@", mvarkey, , , vbTextCompare)
'''
'''    ' Se recibe la respuesta del WS
'''    xmlBody = PostWebservice(mvarurl, "setTransactionQuote", xmlBody)
'''
'''    Set oXMLDoc = New MSXML2.DOMDocument60
'''    Set stTI = New stTransactionQuote
'''
'''    With oXMLDoc
'''        .async = False
'''        .preserveWhiteSpace = False
'''        .validateOnParse = False
'''        .resolveExternals = True
'''        .setProperty "SelectionNamespaces", "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " & _
'''                                            "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
'''                                            "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
'''                                            "xmlns:orb=""" & mvarurl & """"
'''        If .loadXML(xmlBody) Then
'''            Set oRootNode = .selectSingleNode("/soap:Envelope/soap:Body/orb:setTransactionQuoteResponse/orb:setTransactionQuoteResult/orb:stTransactionQuote")
'''            For Each oNode In oRootNode.childNodes
'''                CallByName stTI, oNode.nodeName, VbLet, CStr("" & oNode.Text)
'''            Next oNode
'''        End If
'''    End With
'''    Set oXMLDoc = Nothing
'''
'''    Set setTransactionQuote = stTI
'''
'''End Function

'''Public Function setTransactionSale(ByVal CardNumber As String, _
'''                                   ByVal EmployeeId As String, _
'''                                   ByVal TransactionId As String, _
'''                                   ByVal TransactionItems As String, _
'''                                   ByVal transactionwithdrawal As String, _
'''                                   ByVal InvoiceNumber As String, _
'''                                   ByVal InvoiceDate As String, _
'''                                   ByVal InvoiceAmount As String) As stSaleReturn
'''    Dim oXMLDoc As MSXML2.DOMDocument60 'Object
'''    Dim oRootNode As MSXML2.IXMLDOMNode
'''    Dim oNode As MSXML2.IXMLDOMNode
'''    Dim stTI As stTransactionSale
'''    Dim xmlBody As String
'''
'''    ' Se arma la peticion para el WS
'''    xmlBody = _
'''        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
'''        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
'''        "  <soap:Body>" & _
'''        "    <setTransactionSale xmlns=""@@url@@"">" & _
'''        "      <cardnumber>@@cardnumber@@</cardnumber>" & _
'''        "      <storeid>01</storeid>" & _
'''        "      <posid>01</posid>" & _
'''        "      <employeeid>@@employeeid@@</employeeid>" & _
'''        "      <transactionid>@@transactionid@@</transactionid>" & _
'''        "      <transactionitems>@@transactionitems@@</transactionitems>" & _
'''        "      <transactionwithdrawal>@@transactionwithdrawal@@</transactionwithdrawal>" & _
'''        "      <invoicenumber>@@invoicenumber@@</invoicenumber>" & _
'''        "      <invoicedate>@@invoicedate@@</invoicedate>" & _
'''        "      <invoiceamount>@@invoiceamount@@</invoiceamount>" & _
'''        "      <key>@@key@@</key>" & _
'''        "    </setTransactionSale>" & _
'''        "  </soap:Body>" & _
'''        "</soap:Envelope>"
'''    xmlBody = Replace(xmlBody, "@@url@@", mvarurl, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@cardnumber@@", CardNumber, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@employeeid@@", EmployeeId, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@transactionid@@", TransactionId, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@transactionitems@@", TransactionItems, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@transactionwithdrawal@@", transactionwithdrawal, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@invoicenumber@@", InvoiceNumber, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@invoicedate@@", InvoiceDate, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@invoiceamount@@", InvoiceAmount, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@key@@", mvarkey, , , vbTextCompare)
'''
'''    ' Se recibe la respuesta del WS
'''    xmlBody = PostWebservice(mvarurl, "setTransactionSale", xmlBody)
'''
'''    Set oXMLDoc = New MSXML2.DOMDocument60
'''    Set stTI = New stTransactionSale
'''
'''    With oXMLDoc
'''        .async = False
'''        .preserveWhiteSpace = False
'''        .validateOnParse = False
'''        .resolveExternals = True
'''        .setProperty "SelectionNamespaces", "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " & _
'''                                            "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
'''                                            "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
'''                                            "xmlns:orb=""" & mvarurl & """"
'''        If .loadXML(xmlBody) Then
'''            Set oRootNode = .selectSingleNode("/soap:Envelope/soap:Body/orb:setTransactionSaleResponse/orb:setTransactionSaleResult/orb:stTransactionSale")
'''            For Each oNode In oRootNode.childNodes
'''                CallByName stTI, oNode.nodeName, VbLet, CStr("" & oNode.Text)
'''            Next oNode
'''        End If
'''    End With
'''    Set oXMLDoc = Nothing
'''
'''    Set setTransactionSale = stTI
'''
'''End Function

Public Function setTransactionSaleOffline(ByVal cardnumber As String, _
                                          ByVal EmployeeId As String, _
                                          ByVal TransactionId As String, _
                                          ByVal transactionitems As String, _
                                          ByVal TransactionWithDrawal As String, _
                                          ByVal InvoiceNumber As String, _
                                          ByVal InvoiceDate As String, _
                                          ByVal InvoiceAmount As String) As stSaleReturn
    Dim oXMLDoc As MSXML2.DOMDocument60 'Object
    Dim oRootNode As MSXML2.IXMLDOMNode
    Dim oNode As MSXML2.IXMLDOMNode
    Dim stTI As stTransactionSaleOffline
    Dim xmlBody As String

    ' Se arma la peticion para el WS
    xmlBody = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "  <soap:Body>" & _
        "    <setTransactionSaleOffline xmlns=""@@url@@"">" & _
        "      <cardnumber>@@cardnumber@@</cardnumber>" & _
        "      <storeid>01</storeid>" & _
        "      <posid>01</posid>" & _
        "      <employeeid>@@employeeid@@</employeeid>" & _
        "      <transactionid>@@transactionid@@</transactionid>" & _
        "      <transactionitems>@@transactionitems@@</transactionitems>" & _
        "      <transactionwithdrawal>@@transactionwithdrawal@@</transactionwithdrawal>" & _
        "      <invoicenumber>@@invoicenumber@@</invoicenumber>" & _
        "      <invoicedate>@@invoicedate@@</invoicedate>" & _
        "      <invoiceamount>@@invoiceamount@@</invoiceamount>" & _
        "      <key>@@key@@</key>" & _
        "    </setTransactionSaleOffline>" & _
        "  </soap:Body>" & _
        "</soap:Envelope>"
    xmlBody = Replace(xmlBody, "@@url@@", mvarurl, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@cardnumber@@", cardnumber, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@employeeid@@", EmployeeId, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@transactionid@@", TransactionId, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@transactionitems@@", transactionitems, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@transactionwithdrawal@@", TransactionWithDrawal, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@invoicenumber@@", InvoiceNumber, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@invoicedate@@", InvoiceDate, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@invoiceamount@@", InvoiceAmount, , , vbTextCompare)
    xmlBody = Replace(xmlBody, "@@key@@", mvarkey, , , vbTextCompare)

    ' Se recibe la respuesta del WS
    xmlBody = PostWebservice(mvarurl, "setTransactionSaleOffline", xmlBody)

    Set oXMLDoc = New MSXML2.DOMDocument60
    Set stTI = New stTransactionSaleOffline

    With oXMLDoc
        .async = False
        .preserveWhiteSpace = False
        .validateOnParse = False
        .resolveExternals = True
        .setProperty "SelectionNamespaces", "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " & _
                                            "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
                                            "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
                                            "xmlns:orb=""" & mvarurl & """"
        If .loadXML(xmlBody) Then
            Set oRootNode = .selectSingleNode("/soap:Envelope/soap:Body/orb:setTransactionSaleOfflineResponse/orb:setTransactionSaleOfflineResult/orb:stTransactionSaleOffline")
            For Each oNode In oRootNode.childNodes
                CallByName stTI, oNode.nodeName, VbLet, CStr("" & oNode.Text)
            Next oNode
        End If
    End With
    Set oXMLDoc = Nothing

    Set setTransactionSaleOffline = stTI

End Function

'''Public Function setTransactionUpdate(ByVal CardNumber As String, _
'''                                     ByVal EmployeeId As String, _
'''                                     ByVal TransactionId As String, _
'''                                     ByVal TransactionItems As String) As stSaleReturn
'''    Dim oXMLDoc As MSXML2.DOMDocument60 'Object
'''    Dim oRootNode As MSXML2.IXMLDOMNode
'''    Dim oNode As MSXML2.IXMLDOMNode
'''    Dim stTI As stTransactionUpdate
'''    Dim xmlBody As String
'''
'''    ' Se arma la peticion para el WS
'''    xmlBody = _
'''        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
'''        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
'''        "  <soap:Body>" & _
'''        "    <setTransactionUpdate xmlns=""@@url@@"">" & _
'''        "      <cardnumber>@@cardnumber@@</cardnumber>" & _
'''        "      <storeid>01</storeid>" & _
'''        "      <posid>01</posid>" & _
'''        "      <employeeid>@@employeeid@@</employeeid>" & _
'''        "      <transactionid>@@transactionid@@</transactionid>" & _
'''        "      <transactionitems>@@transactionitems@@</transactionitems>" & _
'''        "      <key>@@key@@</key>" & _
'''        "    </setTransactionUpdate>" & _
'''        "  </soap:Body>" & _
'''        "</soap:Envelope>"
'''    xmlBody = Replace(xmlBody, "@@url@@", mvarurl, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@cardnumber@@", CardNumber, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@employeeid@@", EmployeeId, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@transactionid@@", TransactionId, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@transactionitems@@", TransactionItems, , , vbTextCompare)
'''    xmlBody = Replace(xmlBody, "@@key@@", mvarkey, , , vbTextCompare)
'''
'''    ' Se recibe la respuesta del WS
'''    xmlBody = PostWebservice(mvarurl, "setTransactionUpdate", xmlBody)
'''
'''    Set oXMLDoc = New MSXML2.DOMDocument60
'''    Set stTI = New stTransactionUpdate
'''
'''    With oXMLDoc
'''        .async = False
'''        .preserveWhiteSpace = False
'''        .validateOnParse = False
'''        .resolveExternals = True
'''        .setProperty "SelectionNamespaces", "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " & _
'''                                            "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
'''                                            "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
'''                                            "xmlns:orb=""" & mvarurl & """"
'''        If .loadXML(xmlBody) Then
'''            Set oRootNode = .selectSingleNode("/soap:Envelope/soap:Body/orb:setTransactionUpdateResponse/orb:setTransactionUpdateResult/orb:stTransactionUpdate")
'''            For Each oNode In oRootNode.childNodes
'''                CallByName stTI, oNode.nodeName, VbLet, CStr("" & oNode.Text)
'''            Next oNode
'''        End If
'''    End With
'''    Set oXMLDoc = Nothing
'''
'''    Set setTransactionUpdate = stTI
'''
'''End Function

Public Function getFormatoDocIdentidad(docIdentidadIn As String) As String
    Dim docIdentidad As String
    If (docIdentidadIn <> "") Then
        docIdentidadIn = Trim$(docIdentidadIn)
        Select Case Len(docIdentidadIn)
        Case 8
            docIdentidad = "D0" + docIdentidadIn
        Case 9 To 10
            If (InStr(docIdentidadIn, "D") = 1) Or (InStr(docIdentidadIn, "E") = 1) Then
                docIdentidad = docIdentidadIn
            Else
                docIdentidad = "E" + docIdentidadIn
            End If
        Case Else
            docIdentidad = docIdentidadIn
        End Select
    Else
        docIdentidad = "0"
    End If
    getFormatoDocIdentidad = docIdentidad
End Function

Public Function getFormatoDocIdentidadInversa(docIdentidad As String) As String
    Dim doc As String
    If (docIdentidad <> "") Then
        If (Len(docIdentidad) = 10) Then
            If (InStr(docIdentidad, "D") = 1) Then
                doc = Mid$(docIdentidad, 3)
            Else
                doc = Mid$(docIdentidad, 2)
            End If
        Else
            doc = ""
        End If
    Else
        doc = docIdentidad
    End If
    getFormatoDocIdentidadInversa = doc
End Function

Public Function getTipoDocIdentidad(docIdentidad As String) As String
    Dim doc As String
    If (docIdentidad <> "") Then
        If (Len(docIdentidad) = 10) Then
            If (InStr(docIdentidad, "D") = 1) Then
                doc = "002"
            End If
            If (InStr(docIdentidad, "E") = 1) Then
                doc = "004"
            End If
        End If
    Else
        doc = docIdentidad
    End If
    getTipoDocIdentidad = doc
End Function

Public Function getFormatearFecha(ByVal Fecha As String) As String
    Dim vdate As String
    If (IsDate(Fecha)) Then
        vdate = CStr(Format(Fecha, "yyyyMMdd"))
    End If
    getFormatearFecha = vdate
End Function

Public Function getFormatearFechaFV(ByVal Fecha As String) As String
    Dim vdate As String
    Dim Anio As Integer, Mes As Integer, Dia As Integer
    
    If (Fecha <> "" And Len(Trim$(Fecha)) = 8) Then
        vfecha = Trim$(Fecha)
        If (IsNumeric(Fecha)) Then
            Anio = CInt(Val(Mid$(Fecha, 1, 4)))
            Mes = CInt(Val(Mid$(Fecha, 5, 2)))
            Dia = CInt(Val(Mid$(Fecha, 7, 2)))
            vfecha = CStr(Dia) & "/" & CStr(Mes) & "/" & CStr(Anio)
        End If
    End If
    
    getFormatearFechaFV = vfecha
End Function

Private Sub Class_Initialize()
    mvarurl = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "INDFV409", objUsuario.CodigoEmpresa)
    'mvarurl = IIf(mvarurl <> "", "https://mifarmaws01.orbisfarma.com.mx/", mvarurl)
    If mvarurl = "" Then
        Err.Raise vbObjectError, "clsOrbisFarma.Initialize", "No se encontro la url del Webservice de Orbis (INDFV409) para la CIA " & objUsuario.CodigoEmpresa
    End If
    mvarkey = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "INDFV410", objUsuario.CodigoEmpresa)
    'mvarkey = IIf(mvarkey <> "", "67723f0e27d4869d280cb83bc6ede7d1", mvarkey)
    If mvarkey = "" Then
        Err.Raise vbObjectError, "clsOrbisFarma.Initialize", "No se encontro la KEY del Webservice de Orbis (INDFV409) para la CIA " & objUsuario.CodigoEmpresa
    End If
    SoapAction = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "INDFV411", objUsuario.CodigoEmpresa)
    If SoapAction = "" Then Err.Raise vbObjectError, "clsOrbisFarma.Initialize", "No se encontro la url para la acción de la Webservices de Orbis (INDFV411) para la CIA " & objUsuario.CodigoEmpresa
End Sub
