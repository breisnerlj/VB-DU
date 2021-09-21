VERSION 5.00
Begin VB.UserControl ucGMap 
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6090
   ScaleHeight     =   4515
   ScaleWidth      =   6090
End
Attribute VB_Name = "ucGMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Implementation of the Google-Static-API (limited to 1000 requests per User and Day)
'There's no dependencies to a Browser-Control - just the plain AsynRead-functionality
'of a VB6-Usercontrol (the only two API-calls are used, to blit with HalfTone-Quality)
'Author: Olaf Schmidt (2012)
'2013... adjustments to the location-search-api, which now requires a new URL:
'        "http://maps.googleapis.com/maps/api/geocode/xml?&sensor=false&address=" & UTF8-encoded-Address

Option Explicit

Public Enum MapType
  mt_roadmap
  mt_satellite
  mt_hybrid
  mt_terrain
End Enum

Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nDestWidth As Long, ByVal nDestHeight As Long, ByVal hdcSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal hSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Event MouseMove(ByVal GMouseCoordLatLng As String)
Event MouseUp(ByVal GMouseCoordLatLng As String)
Event DblClick(ByVal GMouseCoordLatLng As String)
 
Private Const mSize& = 640 'this is the max (free usable) SquareSize of the GMap-Static-API
Private Const PI# = 3.14159265358979, TwoPI# = 6.28318530717959
Private Const D2RFac# = 1.74532925199433E-02

Private BackBuf As VB.PictureBox
Private mGZoom As Long, mGPoint As String, mMapType As MapType
Private mLat As Single, mLng As Single
Private mPxlX As Long, mPxlY As Long
Private MDownPoint, LastGMouseMovePoint As String, LastGSearchPoint As String
Private InfoGSearch As String, tagParent As String, firstTagChild As String, secondTagChild As String
Private flgOld As Boolean
Public Markers As New Collection
Public respaldo As String
Public flgBuscaCoords As Boolean
Private objWS As New clsWebService

Dim xyDown As POINTAPI
Dim xyUp As POINTAPI
Public xyIFDist As Integer

Private Sub UserControl_Initialize()
  ScaleMode = vbPixels
  Set BackBuf = Controls.Add("VB.PictureBox", "BackBuf")
  BackBuf.BorderStyle = 0
  BackBuf.AutoRedraw = True
  BackBuf.Move 0, 0, mSize, mSize
  'mGPoint = "-12.083106,-77.012768"
  mGPoint = "20.703879,-40.993700"
  flgOld = True
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
On Error GoTo manipulado
    Dim XML As String
    
    Dim street As String
    Dim Number As String
    Dim apartment As String
    Dim country As String
    Dim city As String
    Dim district As String
    Dim latitude As String
    Dim longitude As String
    Dim receiverName As String
    
    If AsyncProp.StatusCode <> vbAsyncStatusCodeEndDownloadData Then Exit Sub
    If flgOld Then
        Select Case TypeName(AsyncProp.Value)
            Case "Byte()"
              'Dim XML As String
                Dim strAsync As String
                Debug.Print AsyncProp.Target
                strAsync = strConv(AsyncProp.Value, vbUnicode)
                Debug.Print strAsync
                strAsync = objWS.DecodeUTF8(strAsync)
                Debug.Print strAsync
                XML = ReadTagContent(strConv(AsyncProp.Value, vbUnicode), "location")
                LastGSearchPoint = ReadTagContent(XML, "lat") & "," & ReadTagContent(XML, "lng")
                street = ReadTagContent(strAsync, "formatted_address")
                Number = ReadTagContent(strAsync, "address_component", "street_number") 'resp("data")("results")(X)("address_components")(i)("long_name")
                If Len(Trim(Number)) > 0 Then Number = ReadTagContent(Number, "long_name")
                district = ReadTagContent(strAsync, "address_component", "locality") 'resp("data")("results")(X)("address_components")(i)("long_name")
                If Len(Trim(district)) > 0 Then district = ReadTagContent(district, "long_name")
                city = ReadTagContent(strAsync, "address_component", "administrative_area_level") 'resp("data")("results")(X)("address_components")(i)("long_name")
                If Len(Trim(city)) > 0 Then city = ReadTagContent(city, "long_name")
'                If Len(Trim(city)) = 0 Then
'                    If InStr(1, resp("data")("results")(X)("address_components")(i)("types")(j), "administrative_area_level") > 0 Then
'                        city = resp("data")("results")(X)("address_components")(i)("long_name")
'                    End If
'                End If
                country = ReadTagContent(strAsync, "address_component", "country") 'resp("data")("results")(X)("address_components")(i)("long_name")
                If Len(Trim(country)) > 0 Then country = ReadTagContent(country, "long_name")
                latitude = ReadTagContent(XML, "lat") 'resp("data")("results")(X)("geometry")("location")("lat")
                longitude = ReadTagContent(XML, "lng") 'resp("data")("results")(X)("geometry")("location")("lng")
                apartment = ""
                receiverName = ""
                objVenta.dc_street = street
                objVenta.dc_number = Number
                objVenta.dc_apartment = apartment
                objVenta.dc_country = country
                objVenta.dc_departamentBK = procesaCadena(objVenta.dc_departamentBK)
                objVenta.dc_city = city
                objVenta.dc_city = procesaCadena(objVenta.dc_city)
                objVenta.dc_city = IIf(Len(Trim(objVenta.dc_city)) = 0, objVenta.dc_departamentBK, objVenta.dc_city)
                objVenta.dc_city = procesaCadena(objVenta.dc_city)
                objVenta.dc_district = district
                objVenta.dc_district = procesaCadena(objVenta.dc_district)
                objVenta.dc_district = IIf(Len(Trim(objVenta.dc_district)) = 0, objVenta.dc_city, objVenta.dc_district)
                objVenta.dc_district = procesaCadena(objVenta.dc_district)
                objVenta.dc_latitude = latitude
                objVenta.dc_longitude = longitude
'                LastGSearchPoint = ReadTagContent(strConv(AsyncProp.Value, vbUnicode), "location")
            Case "Picture"
                If AsyncProp.BytesRead < 8000 Then Exit Sub
                Set BackBuf.Picture = AsyncProp.Value
                UserControl_Paint
        End Select
    Else
        Select Case TypeName(AsyncProp.Value)
            Case "Byte()"
                'Dim XML As String
                If Len(tagParent) = 0 Then Exit Sub
                'XML = strConv("Malecón del Parque Salazar 610, Miraflores 15074, Perú", vbUnicode)
                'XML = ReadTagContent2(strConv(AsyncProp.Value, vbUnicode), tagParent)
                XML = ReadTagContent(objWS.DecodeURI(strConv(AsyncProp.Value, vbUnicode)), tagParent)
                If Len(firstTagChild) = 0 Then
                    InfoGSearch = XML
                Else
                    InfoGSearch = ReadTagContent(XML, firstTagChild) & IIf(Len(secondTagChild) <> 0, "," & ReadTagContent(XML, secondTagChild), "")
                End If
                'InfoGSearch = ReadTagContent(DecodeURI(strConv(AsyncProp.Value, vbUnicode)), tagParent)
        End Select
    End If
    Exit Sub
manipulado:
    'If Err Then
    'Resume Next
    
        MsgBox ("Dirección no encontrada, por favor verifique | " & vbNewLine & Err.Number & " -- " & Err.Description), vbCritical, App.ProductName
        'Refresh
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)

    Dim buf As String
    Dim PuntosCord() As String
    Dim params As New XArrayDB
    Dim values As New XArrayDB
    Dim Err As New XArrayDB
    Dim i As Integer
    Dim strType As String
    
    params.ReDim 0, -1, 0, 9
    values.ReDim 0, -1, 0, 9
    Err.ReDim 0, -1, 0, 9
    i = 0
    params.AppendRows
    values.AppendRows
    Err.AppendRows
    
    frm_VTA_PreviaTomaPedido.Label3.Caption = AsyncProp.Status
    
    Select Case AsyncProp.StatusCode
        Case vbAsyncStatusCodeSendingRequest
            frm_VTA_PreviaTomaPedido.Label1.Caption = "Solicitando datos"
        Case vbAsyncStatusCodeFindingResource
            frm_VTA_PreviaTomaPedido.Label1.Caption = "Buscando ubicación de descarga"
        Case vbAsyncStatusCodeConnecting
           frm_VTA_PreviaTomaPedido.Label1.Caption = "Conectado"
        Case vbAsyncStatusCodeBeginDownloadData
            frm_VTA_PreviaTomaPedido.Label1.Caption = "Iniciar descarga"
            If AsyncProp.BytesMax = 0 Then
                buf = "0 / desconocido"
            Else
                buf = "0 /" & FormatNumber(AsyncProp.BytesMax, 0)

            End If
        Case vbAsyncStatusCodeEndDownloadData
            frm_VTA_PreviaTomaPedido.flgMapa = True
            frm_VTA_PreviaTomaPedido.Label1.Caption = "Descarga completada"
            If AsyncProp.BytesMax = 0 Then
                buf = FormatNumber(AsyncProp.BytesRead, 0) & "/ unknown"
            Else
                buf = FormatNumber(AsyncProp.BytesRead, 0) & _
                "/" & FormatNumber(AsyncProp.BytesMax, 0)
            End If
        Case vbAsyncStatusCodeDownloadingData
            frm_VTA_PreviaTomaPedido.Label1.Caption = "Descargando ........."
            If AsyncProp.BytesMax = 0 Then
                buf = FormatNumber(AsyncProp.BytesRead, 0) & "/ unknown"
            Else
                buf = FormatNumber(AsyncProp.BytesRead, 0) & _
                "/" & FormatNumber(AsyncProp.BytesMax, 0)
            End If
        Case vbAsyncStatusCodeError
            frm_VTA_PreviaTomaPedido.Label1.Caption = "Se produjo un error. La transferencia se interrumpirá"
    End Select
    'Label3.Caption = buf
    PuntosCord = Split(GPoint, ",")
    objVenta.Latitud = PuntosCord(0)
    objVenta.Longitud = PuntosCord(1)
    If objVenta.Latitud <> "" Or objVenta.Longitud <> "" Then
        frm_VTA_PreviaTomaPedido.cmdTomaPedido.Enabled = True
    Else
        frm_VTA_PreviaTomaPedido.cmdTomaPedido.Enabled = False
    End If
    frm_VTA_PreviaTomaPedido.Label4.Caption = buf & " - Coords : " & GPoint
    
    Select Case AsyncProp.AsyncType
        Case vbAsyncTypeByteArray
            strType = "ByteArray"
        Case vbAsyncTypeFile
            strType = "File"
        Case vbAsyncTypePicture
            strType = "Picture"
    End Select
    params(i, 0) = "" & "AsyncType"
    values(i, 0) = "" & AsyncProp.AsyncType & " | " & strType
    params(i, 1) = "" & "BytesMax"
    values(i, 1) = "" & AsyncProp.BytesMax
    params(i, 2) = "" & "BytesRead"
    values(i, 2) = "" & AsyncProp.BytesRead
    params(i, 3) = "" & "PropertyName"
    values(i, 3) = "" & AsyncProp.PropertyName
    params(i, 4) = "" & "Status"
    values(i, 4) = "" & AsyncProp.Status
    params(i, 5) = "" & "StatusCode"
    values(i, 5) = "" & AsyncProp.StatusCode & " | " & frm_VTA_PreviaTomaPedido.Label1.Caption
    params(i, 6) = "" & "Target"
    values(i, 6) = "" & AsyncProp.Target
    params(i, 7) = "" & "Value"
    values(i, 7) = "" & frm_VTA_PreviaTomaPedido.Label4.Caption
    
    objWS.createLog gstrFlagLogFile3, gstrFlagLogBD3, "log", "dlv_unificado_gmaps", params, values, Err
End Sub
Private Function ReadTagContent_2(sXML As String, Tag As String) As String
'    Dim oXMLDoc As MSXML2.DOMDocument60 'Object
'    Dim oRootNode As MSXML2.IXMLDOMNode
'    Dim oNode As MSXML2.IXMLDOMNode
'    Dim oRootNode2 As MSXML2.IXMLDOMNode
'    Dim oNode2 As MSXML2.IXMLDOMNode
'    Dim Response As String
'    Dim Status As String
'    Set oXMLDoc = New MSXML2.DOMDocument60
'    With oXMLDoc
'        If .loadXML(sXML) Then
'            Status = .selectSingleNode("/GeocodeResponse/status").Text
'            If Status = "ZERO_RESULTS" Then
'                'MsgBox ("Dirección no encontrada, por favor verifique"), vbCritical, App.ProductName
'                'Err.Raise vbObjectError, App.ProductName, "Dirección no encontrada, por favor verifique", vbCritical
'                'MsgBox "No se encontró la dirección, verifique por favor"
'                'Exit Function
'            Else
'                If Tag = "location" Then
'                    Response = .selectSingleNode("/GeocodeResponse/result/geometry/location/lat").Text & "," & .selectSingleNode("/GeocodeResponse/result/geometry/location/lng").Text
'                    Debug.Print Response
'                ElseIf Tag = "formatted_address" Then
'                    Response = .selectSingleNode("/GeocodeResponse/result/formatted_address").Text
'                End If
'            End If
'        End If
'    End With
'    Debug.Print Response
'    Set oXMLDoc = Nothing
'    ReadTagContent = Response
End Function
Private Function ReadTagContent(sXML As String, Tag As String, Optional pType As String) As String

    Dim Result As String
    Dim Existe As String
    Dim Status As String
    Dim Parent As String
    Dim pos As Integer
    Dim lenXML As Integer
    Dim PosBK As Integer
    'On Error GoTo manipulado
    If Tag = "location" Or Tag = "formatted_address" Then
        Status = Mid$(sXML, InStr(sXML, "<status>") + Len("status") + 2)
        Status = left$(Status, InStr(Status, "</status>") - 1)
        Status = Trim$(Replace(Replace(Status, vbCr, ""), vbLf, ""))
        If Status = "OK" Then
            GoTo continua
        End If
    Else
        GoTo continua
    End If
    '***
continua:
    If Len(pType) > 0 Then
        Existe = "0"
        pos = 0
        PosBK = -1
        lenXML = Len(sXML)
        While Existe = "0" And PosBK < pos
            'Debug.Print sXML
            PosBK = pos
            pos = InStr(IIf(pos = 0, 1, pos), sXML, "<" & Tag & ">") + Len(Tag) + 2
            Result = Mid$(sXML, pos)
            'Debug.Print Result
            pos = pos + InStr(Result, "</" & Tag & ">") - 1 + Len(Tag) + 3
            Result = left$(Result, InStr(Result, "</" & Tag & ">") - 1)
            'Debug.Print Result
            If InStr(1, Result, "<type>" & pType) > 0 Then
                Existe = "1"
            End If
        Wend
        If Existe = "0" Then
            Result = ""
        End If
    Else
        Result = Mid$(sXML, InStr(sXML, "<" & Tag & ">") + Len(Tag) + 2)
        Result = left$(Result, InStr(Result, "</" & Tag & ">") - 1)
    End If
    
    Result = Replace(Replace(Result, vbCr, ""), vbLf, "")
    ReadTagContent = Trim$(Result)
End Function

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  x = (x / ScaleWidth - 0.5) * mSize: Y = (Y / ScaleHeight - 0.5) * mSize
  If Button = 1 Then MDownPoint = Array(CLng(x), CLng(Y))
  
    Dim l
    l = GetCursorPos(xyDown)
'    Debug.Print CStr(xyDown.X) & ", " & CStr(xyDown.Y)
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  x = (x / ScaleWidth - 0.5) * mSize: Y = (Y / ScaleHeight - 0.5) * mSize
  LastGMouseMovePoint = Trim(str(PxlYToLat(mPxlY + Y))) & "," & Trim(str(PxlXToLng(mPxlX + x)))
  RaiseEvent MouseMove(LastGMouseMovePoint)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  x = (x / ScaleWidth - 0.5) * mSize: Y = (Y / ScaleHeight - 0.5) * mSize
  If Not IsEmpty(MDownPoint) Then
    Dim dx As Long, dy As Long
    dx = MDownPoint(0) - x: dy = MDownPoint(1) - Y
    MDownPoint = Empty
    GPoint = Trim(str(PxlYToLat(mPxlY + dy))) & "," & Trim(str(PxlXToLng(mPxlX + dx)))
  End If
    Dim l
    l = GetCursorPos(xyUp)
    
'    Debug.Print CStr(xyDown.X) & ", " & CStr(xyDown.Y)
'    Debug.Print CStr(xyUp.X) & ", " & CStr(xyUp.Y)
    xyIFDist = 0
    If (xyDown.x + 15 < xyUp.x Or xyDown.x - 15 > xyUp.x) And xyDown.x <> xyUp.x Then xyIFDist = 1
    If (xyDown.Y + 15 < xyUp.Y Or xyDown.Y - 15 > xyUp.Y) And xyDown.Y <> xyUp.Y And xyIFDist = 0 Then xyIFDist = 1
    
    'Or xyDown.Y <> xyUp.Y + 30 Then xyIFDist = 1
'    Debug.Print xyIFDist
  RaiseEvent MouseUp(LastGMouseMovePoint)
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick(LastGMouseMovePoint)
End Sub

Private Sub UserControl_Resize()
  UserControl_Paint
End Sub
Private Sub UserControl_Paint()
  'Sleep 15
  SetStretchBltMode UserControl.hdc, 4
  'Sleep 15
  StretchBlt hdc, 0, 0, ScaleWidth, ScaleHeight, BackBuf.hdc, 0, 0, BackBuf.Width, BackBuf.Height, vbSrcCopy
End Sub

'Zoom-related Props
Public Property Let GZoom(ByVal newValue As Long)
  If newValue < 0 Then newValue = 0
  If newValue > 20 Then newValue = 20
  If mGZoom = newValue Then Exit Property
  mGZoom = newValue
  mPxlX = LngToPxlX
  mPxlY = LatToPxlY
  'Refresh
End Property
Public Property Get GZoom() As Long
  GZoom = mGZoom
End Property
Public Property Get RealZoom() As Long
  RealZoom = 2 ^ mGZoom
End Property

'Lat,Long "Csv-String-Point"-related Props
Public Property Let GPoint(ByVal newValue As String)
Dim Sarr() As String
respaldo = GPoint
  If mGPoint = newValue Then Exit Property
  mGPoint = IIf(newValue = "", GPoint, newValue)
  Sarr = Split(mGPoint, ",") 'Split(newValue, ",")
  If UBound(Sarr) <> 1 Then Err.Raise vbObjectError, , _
                            "not a valid Lat,Long-Point-Definition"
  mLat = val(Sarr(0))
  mLng = val(Sarr(1))

  mPxlX = LngToPxlX
  mPxlY = LatToPxlY
End Property
Public Property Get GPoint() As String
  GPoint = mGPoint
End Property
Public Property Get Lat() As Double
  Lat = mLat
End Property
Public Property Get Lng() As Double
  Lng = mLng
End Property

'maptype-Props
Public Property Let MapType(ByVal newValue As MapType)
  If newValue < 0 Then newValue = 0
  If newValue > 3 Then newValue = 3
  mMapType = newValue
  'Refresh
End Property
Public Property Get MapType() As MapType
  MapType = mMapType
End Property
Public Function GetMapType(Optional MapType) As String
  If IsMissing(MapType) Then MapType = mMapType
  Select Case MapType
    Case mt_roadmap:   GetMapType = "Normal"
    Case mt_satellite: GetMapType = "Satelital"
    Case mt_hybrid:    GetMapType = "Híbrido"
    Case mt_terrain:   GetMapType = "Terreno"
  End Select
End Function

'all the Pxl to GeoCoord-formulas found in PHP-code from Fabrice Bernhard
Public Function LngToPxlX(Optional Lng, Optional GZoom, Optional GImgWidth) As Long
  If IsMissing(Lng) Then Lng = mLng
  If IsMissing(GZoom) Then GZoom = mGZoom
  If IsMissing(GImgWidth) Then GImgWidth = mSize
  If Lng > 180 Then Lng = 180 Else If Lng < -180 Then Lng = -180
  
  LngToPxlX = (D2RFac * Lng + PI) * 256 / TwoPI * 2 ^ GZoom
End Function

Public Function LatToPxlY(Optional Lat, Optional GZoom, Optional GImgHeight) As Long
  If IsMissing(Lat) Then Lat = mLat
  If IsMissing(GZoom) Then GZoom = mGZoom
  If IsMissing(GImgHeight) Then GImgHeight = mSize
  If Lat > 85 Then Lat = 85 Else If Lat < -85 Then Lat = -85
  
  LatToPxlY = (PI - Log(Tan(PI / 4 + D2RFac * Lat / 2))) * 256 / TwoPI * 2 ^ GZoom
End Function

Public Function PxlXToLng(Optional PxlX, Optional GZoom, Optional GImgWidth) As Single
  If IsMissing(PxlX) Then PxlX = mPxlX
  If IsMissing(GZoom) Then GZoom = mGZoom
  If IsMissing(GImgWidth) Then GImgWidth = mSize
  PxlX = PxlX / 2 ^ GZoom
  If PxlX > GImgWidth Then PxlX = GImgWidth Else If PxlX < 0 Then PxlX = 0
  
  PxlXToLng = (PxlX / 256 * TwoPI - PI) / D2RFac
End Function
 
Public Function PxlYToLat(Optional PxlY, Optional GZoom, Optional GImgHeight) As Single
  If IsMissing(PxlY) Then PxlY = mPxlY
  If IsMissing(GZoom) Then GZoom = mGZoom
  If IsMissing(GImgHeight) Then GImgHeight = mSize
  PxlY = PxlY / 2 ^ GZoom
  If PxlY > GImgHeight Then PxlY = GImgHeight Else If PxlY < 0 Then PxlY = 0
  
  PxlYToLat = (2 * Atn(Exp(PI - PxlY / 256 * TwoPI)) - PI / 2) / D2RFac
End Function

Public Function FindLatLngPointFromTextLocation(TextLocation As String) As String
On Error GoTo Err
    Dim ReqURL As String
    Dim address As String
    'I.ECASTILLO 30.06.2021
    Dim params As New XArrayDB
    Dim values As New XArrayDB
    Dim errors As New XArrayDB
    Dim i As Integer
    'F.ECASTILLO 30.06.2021
    flgOld = True
    address = objWS.UTF8UrlEnc(TextLocation)
    ReqURL = gstrVarURL3 '"https://maps.googleapis.com/maps/api/geocode/xml?&key=AIzaSyAkYwj59-AuMUiPvy0k4uAtSIF0ysP88Ic&sensor=false&address=" & ReqURL & "&components=country:PE" 'UTF8UrlEnc(TextLocation)
    ReqURL = Replace(ReqURL, ":key", gstrVarKEY1)
    ReqURL = Replace(ReqURL, ":address", address)
    'I.ECASTILLO 30.06.2021
    params.ReDim 0, -1, 0, 9
    values.ReDim 0, -1, 0, 9
    errors.ReDim 0, -1, 0, 9
    i = 0
    params.AppendRows
    values.AppendRows
    
    params(i, 0) = "" & "Function"
    values(i, 0) = "" & "FindLatLngPointFromTextLocation"
    params(i, 1) = "" & "URL"
    values(i, 1) = "" & ReqURL
    countCalls "FindLatLngPointFromTextLocation"
    'F.ECASTILLO 30.06.2021
    AsyncRead ReqURL, vbAsyncTypeByteArray, CStr(Timer), vbAsyncReadForceUpdate + vbAsyncReadSynchronousDownload
    FindLatLngPointFromTextLocation = LastGSearchPoint
    'I.ECASTILLO 30.06.2021
    params(i, 2) = "" & "Response"
    values(i, 2) = "" & LastGSearchPoint
    
    objWS.createLog gstrFlagLogFile3, gstrFlagLogBD3, "log", "dlv_unificado_gmaps", params, values, errors
    Exit Function
Err:
    params(i, 2) = "" & "Error"
    values(i, 2) = "" & Err.Number & " | " & Err.Description
    
    objWS.createLog gstrFlagLogFile3, gstrFlagLogBD3, "log", "dlv_unificado_gmaps", params, values, errors, "1"
    'F.ECASTILLO 30.06.2021
End Function

Public Sub SetCenterToTextLocation(TextLocation As String)
    countCalls "SetCenterToTextLocation"
    GPoint = FindLatLngPointFromTextLocation(TextLocation & ", Peru")
  
  'Refresh
End Sub

Public Function GetInfoFromLatLng(coord As String) As String
On Error GoTo Err
    Dim ReqURL As String
    'I.ECASTILLO 30.06.2021
    Dim params As New XArrayDB
    Dim values As New XArrayDB
    Dim errors As New XArrayDB
    Dim i As Integer
    'F.ECASTILLO 30.06.2021
    flgOld = False
    tagParent = "formatted_address"
    'firstTagChild = "lat"
    'secondTagChild = "lng"
    ReqURL = gstrVarURL2 '"https://maps.googleapis.com/maps/api/geocode/xml?&key=AIzaSyAkYwj59-AuMUiPvy0k4uAtSIF0ysP88Ic&sensor=false&latlng=" & coord
    ReqURL = Replace(ReqURL, ":key", gstrVarKEY1)
    ReqURL = ReqURL & coord
    'ReqURL = "https://maps.googleapis.com/maps/api/geocode/json?key=AIzaSyAkYwj59-AuMUiPvy0k4uAtSIF0ysP88Ic&latlng=" & coord
    'I.ECASTILLO 30.06.2021
    params.ReDim 0, -1, 0, 9
    values.ReDim 0, -1, 0, 9
    errors.ReDim 0, -1, 0, 9
    i = 0
    params.AppendRows
    values.AppendRows
    
    params(i, 0) = "" & "Function"
    values(i, 0) = "" & "GetInfoFromLatLng"
    params(i, 1) = "" & "URL"
    values(i, 1) = "" & ReqURL
    countCalls "GetInfoFromLatLng"
    'F.ECASTILLO 30.06.2021
    AsyncRead ReqURL, vbAsyncTypeByteArray, CStr(Timer), vbAsyncReadForceUpdate + vbAsyncReadSynchronousDownload
    GetInfoFromLatLng = InfoGSearch
    'I.ECASTILLO 30.06.2021
    params(i, 2) = "" & "Response"
    values(i, 2) = "" & InfoGSearch
    
    objWS.createLog gstrFlagLogFile3, gstrFlagLogBD3, "log", "dlv_unificado_gmaps", params, values, errors
    Exit Function
    'F.ECASTILLO 30.06.2021
Err:
    params(i, 2) = "" & "Error"
    values(i, 2) = "" & Err.Number & " | " & Err.Description
    
    objWS.createLog gstrFlagLogFile3, gstrFlagLogBD3, "log", "dlv_unificado_gmaps", params, values, errors, "1"
End Function

Public Sub AddMarker(GPosLatLng As String, ByVal Color As Long, MarkerChar As String)
'    MsgBox "Agrega marcador 8"
  Markers.Add "&markers=color:0x" & Color2Hex(Color) & "%7Clabel:" & left$(MarkerChar, 1) & "%7C" & GPosLatLng
End Sub

Public Function Refresh() As Long
    Dim ReqURL As String, m
    Static Counter As Long
    'I.ECASTILLO 30.06.2021
    Dim params As New XArrayDB
    Dim values As New XArrayDB
    Dim errors As New XArrayDB
    Dim i As Integer
    'F.ECASTILLO 30.06.2021
    Counter = Counter + 1
    ReqURL = gstrVarURL1 '"https://maps.googleapis.com/maps/api/staticmap?sensor=false&format=jpg"
    ReqURL = ReqURL & "&center=" & GPoint
    ReqURL = ReqURL & "&zoom=" & GZoom
    ReqURL = ReqURL & "&size=" & mSize & "x" & mSize
    ReqURL = ReqURL & "&maptype=" & GetMapType
    ReqURL = ReqURL & "&key=" & gstrVarKEY1 '"AIzaSyAkYwj59-AuMUiPvy0k4uAtSIF0ysP88Ic"
    flgOld = True
    For Each m In Markers: ReqURL = ReqURL & m: Next
    Debug.Print ReqURL
    'I.ECASTILLO 30.06.2021
    params.ReDim 0, -1, 0, 9
    values.ReDim 0, -1, 0, 9
    errors.ReDim 0, -1, 0, 9
    i = 0
    params.AppendRows
    values.AppendRows
    
    params(i, 0) = "" & "Function"
    values(i, 0) = "" & "Refresh"
    params(i, 1) = "" & "URL"
    values(i, 1) = "" & ReqURL
    'F.ECASTILLO 30.06.2021
    On Error Resume Next
    countCalls "Refresh"
    AsyncRead ReqURL, vbAsyncTypePicture, "C=" & Counter, vbAsyncReadForceUpdate + vbAsyncReadResynchronize
    'I.ECASTILLO 30.06.2021
    params(i, 2) = "" & "Response"
    values(i, 2) = "" & "StaticMap"
    
    objWS.createLog gstrFlagLogFile3, gstrFlagLogBD3, "log", "dlv_unificado_gmaps", params, values, errors
    If Err Then
        params(i, 2) = "" & "Error"
        values(i, 2) = "" & Err.Number & " | " & Err.Description
        
        objWS.createLog gstrFlagLogFile3, gstrFlagLogBD3, "log", "dlv_unificado_gmaps", params, values, errors, "1"
        MsgBox (Err.Number & " : " & Err.Description), vbCritical, App.ProductName
        Err.Clear
    End If
    Refresh = Counter
'    getInfoGoogleMaps
End Function
Private Sub getInfoGoogleMaps()
On Error GoTo Err
    Dim resp As Dictionary
    Dim objWS As New clsWebService
    Dim coord As String
    Dim x, i, j As Integer
    Dim street As String
    Dim Number As String
    Dim apartment As String
    Dim country As String
    Dim city As String
    Dim district As String
    Dim latitude As String
    Dim longitude As String
    Dim receiverName As String
    coord = objVenta.Latitud & "," & objVenta.Longitud
    Set resp = objWS.getInfoGoogleMaps(coord)
    If Not resp Is Nothing Then
        If IsObject(resp("data")) = False Then
            Exit Sub
        End If
        For x = 1 To resp("data")("results").Count()
            If Len(Trim(district)) <= 0 Or Len(Trim(city)) <= 0 Then
                street = resp("data")("results")(x)("formatted_address")
                For i = 1 To resp("data")("results")(x)("address_components").Count()
                    Debug.Print resp("data")("results")(x)("address_components")(i)("types")(1)
                    For j = 1 To resp("data")("results")(x)("address_components")(i)("types").Count()
                        If resp("data")("results")(x)("address_components")(i)("types")(j) = "street_number" Then Number = resp("data")("results")(x)("address_components")(i)("long_name")
                        If resp("data")("results")(x)("address_components")(i)("types")(j) = "locality" Then district = resp("data")("results")(x)("address_components")(i)("long_name")
                        If resp("data")("results")(x)("address_components")(i)("types")(j) = "administrative_area_level_2" Then city = resp("data")("results")(x)("address_components")(i)("long_name")
                        If Len(Trim(city)) = 0 Then
                            If InStr(1, resp("data")("results")(x)("address_components")(i)("types")(j), "administrative_area_level") > 0 Then
                                city = resp("data")("results")(x)("address_components")(i)("long_name")
                            End If
                        End If
                        If resp("data")("results")(x)("address_components")(i)("types")(j) = "country" Then country = resp("data")("results")(x)("address_components")(i)("long_name")
                    Next j
                Next i
                latitude = resp("data")("results")(x)("geometry")("location")("lat")
                longitude = resp("data")("results")(x)("geometry")("location")("lng")
                apartment = ""
                receiverName = ""
            End If
        Next x
        objVenta.dc_street = street
        objVenta.dc_number = Number
        objVenta.dc_apartment = apartment
        objVenta.dc_country = country
        objVenta.dc_city = city
        objVenta.dc_district = district
        objVenta.dc_latitude = latitude
        objVenta.dc_longitude = longitude
        'objVenta.dc_receiverName = receiverName
        'frm_VTA_PreviaTomaPedido.strDireccion_old = objVenta.dc_street
    End If
Err:
    Err.Raise Err.Number, "ucGMap.getInfoGoogleMaps", Err.Description
End Sub
 
'small Helper-Functions
Private Function Color2Hex(Color As Long) As String
  Color2Hex = right("0" & Hex(Color \ 65536), 2) & _
              right("0" & Hex(Color \ 256 And 255), 2) & _
              right("0" & Hex(Color And 255), 2)
End Function


Public Function DefaultCoord(bool As Boolean, Optional coord As String = "0,0")
    If bool Then
        GPoint = coord
    End If
End Function




Public Function countCalls(ByVal strMethod As String)
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If garrCallGoogleMaps.Count(1) < 0 Then Exit Function
    Dim i As Integer
    Dim encontro As Boolean
    aux = garrCallGoogleMaps.Count(1)
    While i < aux
        If garrCallGoogleMaps(i, 0) = strMethod Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = garrCallGoogleMaps.Count(1) 'Llena con la ultima posicion disponible
        End If
        i = i + 1
    Wend
    If encontro = False Then
        garrCallGoogleMaps.AppendRows
    End If

j:
    If garrCallGoogleMaps.Count(1) = 0 Then ultimo = 0: garrCallGoogleMaps.AppendRows
    
    garrCallGoogleMaps(ultimo, 0) = strMethod
    garrCallGoogleMaps(ultimo, 1) = IIf(ultimo = 0, 1, garrCallGoogleMaps(ultimo, 1) + 1)
    
End Function

Function procesaCadena(ByVal str As String) As String
    Dim newStr As Variant
    Dim response As String
    newStr = Split(str, ",")
    response = newStr(0)
    response = Replace(response, "¿", "")
    response = Trim(response)
    procesaCadena = response
End Function
