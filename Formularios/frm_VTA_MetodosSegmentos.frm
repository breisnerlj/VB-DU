VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_VTA_MetodosSegmentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Metodos y Segmentos"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   8460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRET 
      Caption         =   "Retiro en Tienda"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   6360
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "EXP"
      TabPicture(0)   =   "frm_VTA_MetodosSegmentos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdSegmentos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin vbp_Ventas.ctlGrillaArray grdSegmentos 
         Height          =   4335
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   7646
         MenuPopUp       =   0   'False
         Resalte         =   0   'False
      End
   End
   Begin VB.Label lblValTipoServicio 
      Caption         =   "La capacidad que elegiste ya no está disponible,selecciona otro horario."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "frm_VTA_MetodosSegmentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objProducto As New clsProducto
Dim data As New Dictionary
Dim xListaArray As New XArrayDB
Public Parametro As String
Public Tipo As String '1 => ubigeo, 2 => coordenadas, 3 => local
Public permiteCerrar As String
Private grdLeft As Double
Private grdTop As Double
Private flgTabClick As Boolean
Dim objWS As New clsWebService
Dim objLocal As New clsLocal
Public CodLocalPrecio As String
Public CodLocalReferencia As String
Public lat1, lng1 As String

'
Private strGrupo As String
Private strCapacidad As String
Private strPedidos As String
Private strHora As String
Private strValor As String
Private strSegmento As String
Private strSegundaHora As String 'I.CVIERA 20.03.2021
Private strHabilitado As String
Private strTipoCode As String
Private strDia As String
Public strPrecioTipo As String
Public MS_continua As Boolean
Private strDeliveryTime As String
Public strMessage As String
Public strDateTime As String
Private strStarHour As String
Private strEndHour As String


Private Sub chkRET_Click()
    If chkRET.Value = 1 Then
        SSTab1.Enabled = False
        grdSegmentos.Enabled = False
    Else
        SSTab1.Enabled = True
        grdSegmentos.Enabled = False
    End If
End Sub

Private Sub cmdAceptar_Click()
'I.CVIERA 21.12.2020
    CodLocalPrecio = "" '"LocalCode" | localCode
    strDia = "" '"Dia" | day
    strMessage = "" '"Fecha" | message
    strStarHour = "" '"starHour" | starHour
    strEndHour = "" '"endHour" | endHour
    strDeliveryTime = "" '"Tiempo de dlv" | deliveryTime
    strSegmento = "" '"Segmento" | time
    strValor = "" '"Valor" | value
    strSegundaHora = "" '"Valor_Fin" | valueEnd
    strCapacidad = ""
    'strValor2 = ""
    strGrupo = ""
    strPedidos = ""
    strHora = ""
    strHabilitado = ""
    strTipoCode = ""
    strPrecioTipo = ""
    Dim ii As Integer
    Dim flgCapacidadDisp As String
    Dim agregoProducto As Integer
    Dim strCiaDespacho As String
    strDateTime = ""
    '"Dia", "Fecha", "Tiempo de Dlv", "Segmento", "Valor", "LocalCode", "Valor_Fin"
    Dim x, i, j, a, b As Integer
    Dim objLocal As New clsLocal
    'I.ECASTILLO 05.01.2021
    Dim rsCia As oraDynaset
    Dim Cia
    'F.ECASTILLO 05.01.2021
    strDateTime = DateTime.Now
    'strDateTime = Format(strDateTime, "dddd, dd mmmm")
    objVenta.bk_chkRET = chkRET.Value
    If grdSegmentos.ApproxCount = 0 Then GoTo salir
    CodLocalPrecio = objLocal.GetCodBTL(grdSegmentos.Columns("LocalCode").Value)
    strDia = grdSegmentos.Columns("Dia").Value
    strMessage = grdSegmentos.Columns("Fecha").Value
    strStarHour = grdSegmentos.Columns("starHour").Value
    strEndHour = grdSegmentos.Columns("endHour").Value
    strDeliveryTime = grdSegmentos.Columns("Tiempo de Dlv").Value
    strSegmento = grdSegmentos.Columns("Segmento").Value
    strValor = grdSegmentos.Columns("Valor").Value
    strSegundaHora = grdSegmentos.Columns("Valor_Fin").Value 'ECASTILLO 28.05.2021
    CodLocalReferencia = CodLocalPrecio
    'I.ECASTILLO 05.01.2021
    objVenta.respetaLocal = True
    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, CodLocalReferencia)
    If (rsCia.RecordCount > 0) Then
      Cia = CStr(rsCia(1))
    End If
    Set rsCia = Nothing
    'cboCia.BoundText = cia
    mdiPrincipal.ctlCliente1.sCia = Cia
    mdiPrincipal.ctlCliente1.LimpiarLocalDespacho
    mdiPrincipal.ctlCliente1.LocalAsignado = ""
    mdiPrincipal.ctlCliente1.LocalDespacho = ""
    'F.ECASTILLO 05.01.2021
    mdiPrincipal.ctlCliente1.reasignaLocales CodLocalPrecio, CodLocalReferencia
'    mdiPrincipal.ctlCliente1.LocalAsignado = CodLocalPrecio
'    mdiPrincipal.ctlCliente1.LocalDespacho = CodLocalReferencia
    If objVenta.bk_chkRET = "1" Then
'        strTipoCode = "RET"
'        strDia = "" & Format(strDateTime, "YYYY-mm-dd")
'        strMessage = "" & Format(strDateTime, "dddd, dd mmmm")
'        strStarHour = ""
'        strEndHour = ""
'        strDeliveryTime = ""
'        strSegmento = ""
'        strValor = ""
'        strSegundaHora = ""
'        GoTo continuarFlujo
        GoTo salir
    End If

    i = 0
    For i = 1 To data("data").Count()
        For x = 1 To data("data")(i)("deliveryMethods").Count()
            strTipoCode = data("data")(i)("deliveryMethods")(x)("serviceTypeCode")
            strPrecioTipo = data("data")(i)("deliveryMethods")(x)("deliveryAmount")
            objVenta.bk_amount = strPrecioTipo
            If SSTab1.Caption = "EXP" Then
                If SSTab1.Caption = strTipoCode Then
                    Exit For
                End If
            ElseIf SSTab1.Caption = "AM_PM" Then
                If SSTab1.Caption = strTipoCode Then
                    Exit For
                End If
            ElseIf SSTab1.Caption = "PROG" Then
                If SSTab1.Caption = strTipoCode Then
                    Exit For
                End If
            End If
        Next x
    Next i
    If frm_VTA_Busqueda.grdProductos.ApproxCount > 0 Then
        If frm_VTA_Busqueda.grdProductos.Columns("COD_PRODUCTO").Value = "09938" Then
            frm_VTA_Busqueda.grdProductos.Columns("PRECIO").Value = Format(strPrecioTipo, "###,###.00")
            frm_VTA_Busqueda.grdProductos.Columns("FLG_SEG").Value = 1
        End If
    End If
    Dim Buscar As String ' | 05.01.2021.REVISAR
    'AGREGAR AL CARRITO DE FORMA AUTOMATICA EL PRODUCTO DLV
    If objVenta.bk_ServiceType <> strTipoCode Then
        Dim od As oraDynaset
        Set od = objProducto.Lista(objUsuario.CodigoEmpresa, _
                                    mdiPrincipal.ctlCliente1.LocalAsignado, _
                                    "003", _
                                    Trim("525527"), "", "", _
                                    mdiPrincipal.ctlCliente1.LocalDespacho, _
                                    objVenta.CodModalidadVenta, _
                                    objVenta.CodigoConvenio, _
                                    mdiPrincipal.ctlCliente1.sCia, _
                                    "0")
        If od(0) <> -1 Then
            frmPedido.grdPedido.Array1 = objVenta.AgregaProducto(od(0), _
                                                                    od(22), 1, 0, _
                                                                    od(4), objVenta.CodigoTipoVenta, _
                                                                    0, , , , , , _
                                                                    "1", _
                                                                    0, _
                                                                    "", _
                                                                    "", , _
                                                                    od(4), , , , , 0)
            agregoProducto = 1
        End If
        'I.ECASTILLO 05.01.2021 | busca producto dlv en los productos del carrito y actualiza precio
        If agregoProducto = 1 Then
            Dim nuevoPrecio As Double
            For i = 0 To objVenta.Producto.UpperBound(1)
                Buscar = objVenta.Producto(i, 0)
                If Buscar = "09938" Then
                    nuevoPrecio = IIf(objVenta.Producto(i, 6) = "0", objVenta.Producto(i, 5) * CDbl(val(strPrecioTipo)), (CDbl(val(strPrecioTipo)) / 1) * objVenta.Producto(i, 3))
                    objVenta.AgregaProducto objVenta.Producto(i, 0), objVenta.Producto(i, 1), _
                                            objVenta.Producto(i, 3), objVenta.Producto(i, 6), _
                                            nuevoPrecio, objVenta.Producto(i, 5), _
                                            objVenta.Producto(i, 7), , , , , , _
                                            objVenta.Producto(i, 21), objVenta.Producto(i, 13), _
                                            objVenta.Producto(i, 22), objVenta.Producto(i, 23), , _
                                            nuevoPrecio, , , objVenta.Producto(i, 2), , , _
                                            1
                End If
            Next i
            If objVenta.Producto.UpperBound(1) >= 0 Then
                frmPedido.Cal_Promo
                frmPedido.Cal_Montos
                frmPedido.grdPedido.Rebind
                'frmPedido.grdPedido.Refresh
            End If
        End If
    End If
    'F.ECASTILLO 05.01.2021
continuarFlujo:
    MS_continua = True
    frm_VTA_PreviaTomaPedido.flgContinua = True
    mdiPrincipal.ctlCliente1.seleccionManualLocal = True
    objVenta.flgDatosCapacidad = True
    If grdSegmentos.ApproxCount <> 0 Then
        objVenta.bk_codLocalCapacidad = grdSegmentos.Columns("LocalCode").Value
        objVenta.bk_FechaCapacidad = strDia
        objVenta.bk_message = strMessage
        objVenta.bk_starHour = strStarHour
        objVenta.bk_endHour = strEndHour
        objVenta.bk_deliveryTime = strDeliveryTime
        objVenta.bk_segmento = strSegmento
        objVenta.bk_HoraCapacidad = strValor
        objVenta.bk_HoraCapacidad2 = strSegundaHora 'I.CVIERA 20.03.2021
    End If
    If Len(Trim(objVenta.bk_codLocalCapacidad)) = 0 Then
        objVenta.bk_codLocal = mdiPrincipal.ctlCliente1.LocalDespacho
    Else
        objVenta.bk_codLocal = objLocal.GetCodBTL(objVenta.bk_codLocalCapacidad)
    End If
    objVenta.bk_ServiceType = strTipoCode
    objVenta.bk_flgPactado = 1
    'I.ECASTILLO 27.10.2020
    '17.12.2020 | se comenta función, volver a usar
    objVenta.isLocalDcCappa = "" & objLocal.GetIndCDCAP(objVenta.bk_codLocal)
    'F.ECASTILLO 27.10.2020
    strCiaDespacho = mdiPrincipal.ctlCliente1.sCia
    
    If strCiaDespacho = "" Then strCiaDespacho = objUsuario.CodigoEmpresa
    objVenta.isDCSAP = "" & objLocal.GetEstConfig(strCiaDespacho, objVenta.bk_codLocal, "STOCK_DC_SAP")
        
    ii = 0
    While ii < objVenta.xMetodoSegmento.Count(1)
        objVenta.xMetodoSegmento.DeleteRows (0)
    Wend
    objVenta.AgregaMetodoSegmento objVenta.bk_codLocal, objVenta.bk_ServiceType, _
                                  objVenta.bk_amount, objVenta.bk_FechaCapacidad, _
                                  objVenta.bk_message, objVenta.bk_deliveryTime, _
                                  objVenta.bk_HoraCapacidad, objVenta.bk_HoraCapacidad2, _
                                  objVenta.bk_segmento, objVenta.bk_starHour, _
                                  objVenta.bk_endHour
    If gstrIndRAv3 = "2" Then
        flgCapacidadDisp = objVenta.validaCapacidad
        lblValTipoServicio.Visible = False
        If flgCapacidadDisp <> "1" Then
            lblValTipoServicio.Visible = True
            Exit Sub
        End If
    End If
    
    Unload Me
salir:
    MS_continua = True
    frm_VTA_PreviaTomaPedido.flgContinua = True
    mdiPrincipal.ctlCliente1.seleccionManualLocal = True
    If objVenta.bk_chkRET = "1" Then
        strTipoCode = "RET"
        strDia = "" & Format(strDateTime, "YYYY-mm-dd")
        strMessage = "" & Format(strDateTime, "dddd, dd mmmm")
        strDeliveryTime = ""
        strSegmento = ""
        strValor = DateTime.Now
        strValor = Format(strValor, "hh:mm:ss")
        strSegundaHora = DateTime.Now
        strSegundaHora = Format(strSegundaHora, "hh:mm:ss")
        
        objVenta.bk_ServiceType = strTipoCode
        objVenta.bk_FechaCapacidad = strDia
        objVenta.bk_message = strMessage
        objVenta.bk_starHour = strStarHour
        objVenta.bk_endHour = strEndHour
        objVenta.bk_deliveryTime = strDeliveryTime
        objVenta.bk_segmento = strSegmento
        objVenta.bk_HoraCapacidad = strValor
        objVenta.bk_HoraCapacidad2 = strSegundaHora
        objVenta.isLocalDcCappa = "" & objLocal.GetIndCDCAP(mdiPrincipal.ctlCliente1.LocalDespacho)
        objVenta.bk_flgPactado = 1
        If Len(Trim(objVenta.bk_codLocalCapacidad)) = 0 Then
            objVenta.bk_codLocal = mdiPrincipal.ctlCliente1.LocalDespacho
        Else
            objVenta.bk_codLocal = objLocal.GetCodBTL(objVenta.bk_codLocalCapacidad)
        End If
        
        'I.ECASTILLO 27.10.2020
        '17.12.2020 | se comenta función, volver a usar
        objVenta.isLocalDcCappa = "" & objLocal.GetIndCDCAP(objVenta.bk_codLocal)
        'F.ECASTILLO 27.10.2020
        strCiaDespacho = mdiPrincipal.ctlCliente1.sCia
        
        If strCiaDespacho = "" Then strCiaDespacho = objUsuario.CodigoEmpresa
        objVenta.isDCSAP = "" & objLocal.GetEstConfig(strCiaDespacho, objVenta.bk_codLocal, "STOCK_DC_SAP")
    
        ii = 0
        While ii < objVenta.xMetodoSegmento.Count(1)
            objVenta.xMetodoSegmento.DeleteRows (0)
        Wend
        objVenta.AgregaMetodoSegmento objVenta.bk_codLocal, objVenta.bk_ServiceType, _
                                      objVenta.bk_amount, objVenta.bk_FechaCapacidad, _
                                      objVenta.bk_message, objVenta.bk_deliveryTime, _
                                      objVenta.bk_HoraCapacidad, objVenta.bk_HoraCapacidad2, _
                                      objVenta.bk_segmento, objVenta.bk_starHour, _
                                      objVenta.bk_endHour
        If gstrIndRAv3 = "2" Then
            flgCapacidadDisp = objVenta.validaCapacidad
            lblValTipoServicio.Visible = False
            If flgCapacidadDisp <> "1" Then
                lblValTipoServicio.Visible = True
                Exit Sub
            End If
        End If
        Unload Me
    End If
'F.CVIERA 21.12.2020
End Sub

'I.ECASTILLO 17.12.2020
Private Sub Form_Activate()
    Dim coords, Lat, Lng As String
    Dim x, i, j, xx, ii, jj As Integer
    Dim rsCia As oraDynaset
    Dim sCia As String
    Dim channel As String
    Dim Cia As String
    
    
    If gstrIndRAv3 = "2" Then
        Dim flgCapacidadDisp As String
        If Len(objVenta.bk_ServiceType) > 0 And objVenta.bk_ServiceType <> "RET" Then
            flgCapacidadDisp = objVenta.validaCapacidad
            lblValTipoServicio.Visible = False
            If flgCapacidadDisp <> "1" Then
                lblValTipoServicio.Visible = True
            End If
        End If
    End If
    
    MS_continua = False
    grdLeft = grdSegmentos.left
    grdTop = grdSegmentos.top
    flgTabClick = False
    'I.ECASTILLO 17.09.2021 PARAMETRIZAR MARCA
    Set rsCia = gclsOracle.FN_Cursor("btlprod.pkg_local.get_cia_x_local", 0, mdiPrincipal.ctlCliente1.LocalDespacho)
    If (rsCia.RecordCount > 0) Then
      sCia = CStr(rsCia(1))
    End If
    Set rsCia = Nothing
    'ENVIAR CANAL Y CIA
    Cia = "" & objLocal.GetMarcaLocal(mdiPrincipal.ctlCliente1.LocalDespacho, 2)
    Cia = Trim(Cia)
    'F.ECASTILLO 17.09.2021
    channel = "CALL"
    Sleep 75
    If Tipo = "1" Then
        If Len(Trim(Cia)) = 0 Then
            Select Case sCia
                Case "94", "93", "92", "1DLV"
                    Cia = "IKF"
                    Set data = objWS.listaMetodosSegmentos_Ubigeo(Parametro, channel, Cia)
                Case ""
                    Set data = Nothing
                Case Else
                    Cia = "MF"
                    Set data = objWS.listaMetodosSegmentosMF_Ubigeo(Parametro, channel, Cia)
            End Select
        Else
            Set data = objWS.listaMetodosSegmentos_Ubigeo(Parametro, channel, Cia)
        End If
    ElseIf Tipo = "2" Then
        coords = Split(Parametro, ",")
        Lat = coords(0)
        lat1 = Trim(Lat)
        Lng = coords(1)
        lng1 = Trim(Lng)
        If Len(Trim(Cia)) = 0 Then
            Select Case sCia
                Case "94", "93", "92", "1DLV"
                    Cia = "IKF"
                    Set data = objWS.listaMetodosSegmentos_Coords(Lat, Lng, channel, Cia)
                Case ""
                    Set data = Nothing
                Case Else
                    Cia = "MF"
                    Set data = objWS.listaMetodosSegmentosMF_Coords(Lat, Lng, channel, Cia)
            End Select
        Else
            Set data = objWS.listaMetodosSegmentos_Coords(Lat, Lng, channel, Cia)
        End If
    ElseIf Tipo = "3" Then
        If Len(Trim(Cia)) = 0 Then
            Select Case sCia
                Case "94", "93", "92", "1DLV"
                    Cia = "IKF"
                    Set data = objWS.listaMetodosSegmentos_Local(Parametro, channel, Cia)
                Case ""
                    Set data = Nothing
                Case Else
                    Cia = "MF"
                    Set data = objWS.listaMetodosSegmentosMF_Local(Parametro, channel, Cia)
            End Select
        Else
            Set data = objWS.listaMetodosSegmentos_Local(Parametro, channel, Cia)
        End If
    End If
    setFormatGrid
    xListaArray.ReDim 0, -1, 0, 8
    grdSegmentos.Array1 = xListaArray
    If Not data Is Nothing Then
        If IsObject(data("data")) = False Then
            Exit Sub
        End If
        x = 0

        i = 0
        SSTab1.Tab = 0
        SSTab1.Caption = "-----"
        For i = 1 To data("data").Count()
            If IsObject(data("data")(i)) = False Then
                For x = 1 To data("data")("deliveryMethods").Count()
                    If x = 1 Then
                        SSTab1.Caption = data("data")("deliveryMethods")(x)("serviceTypeCode")
                    Else
                        SSTab1.Tabs = SSTab1.Tabs + 1
                        SSTab1.Tab = SSTab1.Tabs - 1
                        SSTab1.Caption = data("data")("deliveryMethods")(x)("serviceTypeCode")
                    End If
                Next x
            Else
                For x = 1 To data("data")(i)("deliveryMethods").Count()
                    If x = 1 Then
                        SSTab1.Caption = data("data")(i)("deliveryMethods")(x)("serviceTypeCode")
                    Else
                        SSTab1.Tabs = SSTab1.Tabs + 1
                        SSTab1.Tab = SSTab1.Tabs - 1
                        SSTab1.Caption = data("data")(i)("deliveryMethods")(x)("serviceTypeCode")
                    End If
                Next x
            End If
        Next i
    Else
        SSTab1.Caption = "-----"
    End If
    flgTabClick = True
    SSTab1.Tab = 0
    SSTab1_Click 1
End Sub

Function setFormatGrid()
    Dim arrCampos As Variant
    Dim arrCaption As Variant
    Dim arrAncho As Variant
    Dim arrAlineacion As Variant
    'dia
    'mensaje
    'capacidad
    'tiempo delivery
    'segmento
    'valor
    'local
    arrCampos = Array("", "", "", "", "", "", "", "", "")
    arrCaption = Array("Dia", "Fecha", "Tiempo de Dlv", "Segmento", "Valor", "LocalCode", "Valor_Fin", "starHour", "endHour")
    arrAncho = Array(1500, 2000, 1500, 2000, 0, 0, 0, 0, 0)
    arrAlineacion = Array(dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft, dbgLeft)
    
    grdSegmentos.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    grdSegmentos.Columns("Valor").Visible = False
    grdSegmentos.Columns("LocalCode").Visible = False
    grdSegmentos.Columns("Valor_Fin").Visible = False
    grdSegmentos.Columns("starHour").Visible = False
    grdSegmentos.Columns("endHour").Visible = False
End Function

Private Sub Form_Unload(Cancel As Integer)
    If MS_continua = False Then
        frm_VTA_PreviaTomaPedido.flgContinua = False
        mdiPrincipal.ctlCliente1.seleccionManualLocal = False
        objVenta.flgDatosCapacidad = False
    End If
    Unload Me
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Dim x, i, j, xx, ii, jj As Integer
    If flgTabClick = False Then
        Exit Sub
    End If
    xListaArray.ReDim 0, -1, 0, 9
    grdSegmentos.left = grdLeft
    grdSegmentos.top = grdTop
    grdSegmentos.MoveFirst
    grdSegmentos.Limpiar
    grdSegmentos.Array1 = xListaArray
    If Not data Is Nothing Then
        If IsObject(data("data")) = False Then
            Exit Sub
        End If
        x = 0
        i = 0
        For i = 1 To data("data").Count()
            If IsObject(data("data")(i)) = False Then
                For x = 1 To data("data")("deliveryMethods").Count()
                    If SSTab1.Caption = data("data")("deliveryMethods")(x)("serviceTypeCode") Then
                        j = 0
                        ii = 0
                        If IsObject(data("data")("deliveryMethods")(x)("capabilities")) = True Then
                            For j = 1 To data("data")("deliveryMethods")(x)("capabilities").Count()
                                xx = 0
                                For xx = 1 To data("data")("deliveryMethods")(x)("capabilities")(j)("schedules").Count()
                                    xListaArray.AppendRows
                                    'dia
                                    'mensaje
                                    'tiempo delivery
                                    'segmento
                                    'valor
                                    'local
                                    xListaArray(ii, 0) = data("data")("deliveryMethods")(x)("capabilities")(j)("day")
                                    xListaArray(ii, 1) = data("data")("deliveryMethods")(x)("capabilities")(j)("message")
                                    xListaArray(ii, 2) = data("data")("deliveryMethods")(x)("capabilities")(j)("deliveryTime")
                                    xListaArray(ii, 3) = data("data")("deliveryMethods")(x)("capabilities")(j)("schedules")(xx)("time")
                                    xListaArray(ii, 4) = data("data")("deliveryMethods")(x)("capabilities")(j)("schedules")(xx)("value")
                                    xListaArray(ii, 5) = data("data")("localCode")
                                    xListaArray(ii, 6) = data("data")("deliveryMethods")(x)("capabilities")(j)("schedules")(xx)("valueEnd")
                                    xListaArray(ii, 7) = data("data")("deliveryMethods")(x)("capabilities")(j)("startHour")
                                    xListaArray(ii, 8) = data("data")("deliveryMethods")(x)("capabilities")(j)("endHour")
                                    ii = ii + 1
                                Next xx
                            Next j
                        End If
                    End If
                Next x
            Else
                For x = 1 To data("data")(i)("deliveryMethods").Count()
                    ii = 0
                    If SSTab1.Caption = data("data")(i)("deliveryMethods")(x)("serviceTypeCode") Then
                        j = 0
                        ii = 0
                        If IsObject(data("data")(i)("deliveryMethods")(x)("capabilities")) = True Then
                            For j = 1 To data("data")(i)("deliveryMethods")(x)("capabilities").Count()
                                xx = 0
                                For xx = 1 To data("data")(i)("deliveryMethods")(x)("capabilities")(j)("schedules").Count()
                                    xListaArray.AppendRows
                                    'dia
                                    'mensaje
                                    'tiempo delivery
                                    'segmento
                                    'valor
                                    'local
                                    xListaArray(ii, 0) = data("data")(i)("deliveryMethods")(x)("capabilities")(j)("day")
                                    xListaArray(ii, 1) = data("data")(i)("deliveryMethods")(x)("capabilities")(j)("message")
                                    xListaArray(ii, 2) = data("data")(i)("deliveryMethods")(x)("capabilities")(j)("deliveryTime")
                                    xListaArray(ii, 3) = data("data")(i)("deliveryMethods")(x)("capabilities")(j)("schedules")(xx)("time")
                                    xListaArray(ii, 4) = data("data")(i)("deliveryMethods")(x)("capabilities")(j)("schedules")(xx)("value")
                                    xListaArray(ii, 5) = data("data")(i)("localCode")
                                    xListaArray(ii, 6) = data("data")(i)("deliveryMethods")(x)("capabilities")(j)("schedules")(xx)("valueEnd")
                                    xListaArray(ii, 7) = data("data")(i)("deliveryMethods")(x)("capabilities")(j)("startHour")
                                    xListaArray(ii, 8) = data("data")(i)("deliveryMethods")(x)("capabilities")(j)("endHour")
                                    ii = ii + 1
                                Next xx
                            Next j
                        End If
                    End If
                Next x
            End If
        Next i
    End If
    
    grdSegmentos.Array1 = xListaArray
    grdSegmentos.Rebind
    grdSegmentos.Refresh
End Sub
'F.ECASTILLO 17.12.2020
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        If permiteCerrar = "1" Then
            Unload Me
        Else
            frm_VTA_PreviaTomaPedido.flgContinua = False
            mdiPrincipal.ctlCliente1.seleccionManualLocal = False
            objVenta.flgDatosCapacidad = False
            MS_continua = False
            Unload Me
        End If
    End Select
End Sub
