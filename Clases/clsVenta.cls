VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim objMaquina As New clsMaquina
Dim objDocumento As New clsDocumento
Dim objWS As New clsWebService
Dim objLocal As New clsLocal
Dim Cumplio As Boolean
Public CUANTAMAXIMA As Integer
'*************************************************************************************
' Variables que se retornaran como propiedad, Solo seran usadas dentro de la clase
'*************************************************************************************
Private xServicio As New XArrayDB ' Variable que se devuelve como propiedad los producto
Private xMaquina As New XArrayDB  'Variable que se devuelve como propiedad las maquinas
Private xProducto As New XArrayDB ' Variable que se devuelve como propiedad los producto
Public xProductoRegalo As New XArrayDB ' Variable que se devuelve como propiedad los producto
Public xProductoBK As New XArrayDB
Private xProductoLoteDetalle As New XArrayDB ' Variable que se devuelve como propiedad los producto As New XArrayDB ' Variable que se devuelve como propiedad los producto
Public xCuantaMaxima As New XArrayDB 'ECASTILLO 04.07.2020
Private xConvenio As New XArrayDB ' Variable que se devuelve como propiedad los Convenio
Private xFormaPago As New XArrayDB ' Variable que se devuelve como propiedad los Convenio
Private XDistribucion As New XArrayDB ' Variable que se devuelve como propiedad los Convenio
Private xTransferencia As New XArrayDB 'Variable arreglo para las transferencias de rute
Public xMetodoSegmento As New XArrayDB
Private xDiagnostico As New XArrayDB
Private arrDiagnostico As Variant
Private xCobroResponsabilidad As New XArrayDB
Private xRecetarioMagistral  As New XArrayDB
Private xEspeciesValoradas  As New XArrayDB
Private xEntregaTerceros As New XArrayDB
Private xDatoAdicional As New XArrayDB
Private xProductoCobro As New XArrayDB
Private xProductoLote As New XArrayDB
Private lNumMaximoUnidades As Integer
Private xArrayPlanVital As New XArrayDB
Dim lLocalOrigenVital As String
Dim lCIE10Vital As String
Dim lCMPVital As String
Dim sEscaneoTarjeta As String
Dim lCodigoConvenioVital As String
Private strCodigoConvenio As String
Private strNombreConvenio As String
Private strCodigoBeneficiario As String
Private strCodigoTipoVenta As TipoVenta
Private sPuntoInicial As String
Private sNumTarjetaPuntos As String
Private strNombreBeneficiario As String
Private strFlgTipoConvenio As String
Private strTipoCliente As String
Private ImpdblPctBeneficiario As Double
Private strMotivoNotaCredito As String
Private strRuc As String
Private strLocalAtencion As String
Private strLocalDespacho As String
Private strRazonSocial As String
Private strNumeroDocumentoID As String
Private strNombreCliente As String
Private strApellidoCliente As String
Private strDireccionCliente As String
Private strCodigoCliente As String
Private strCodigoClienteDLV As String
Private strDireccionClienteSocial As String
Private strCodigoDocumentoVenta As String
Private strProductoTransferido As String
Private strUbigeoEntrega As String
Private strCodModalidadVenta As String
Private strMontoTotalTotal As String
'-----------
Private byteNumElementosCobro As Byte
Private byteNumElementosFormaPago As Byte
Private byteNumElementosProd As Byte
Private byteNumElementosRM As Byte
Private byteNumElementosServ As Byte
Private byteNumElementosDatoAdic As Byte
Private byteNumElementosEspVal As Byte
Private byteNumElementosPlanVital As Byte

'----------para convenios
Private strflgBeneficiarios As String
Private dblPctBeneficiario As Double
Private dblPctBeneficiarioReal As Double
Private dblPctBeneficiarioOriginal As Double
Private flgIntPctBeneficiario As Integer
Private strFlgRepartidor As String
Private strFlgMedico As String
Private strFlgPolitica As String
Private strCodConvenio As String
Private strCodBeneficiario As String
Private strNomConvenio As String
Private strNomBeneficiario As String
Private strCodEmpresa As String
Private strCodMedico As String
Private strFlgValidaLinCre As String
'-----------para clientes
Private strCodCliente As String
'-----------para documento
Private strCodTipDocVta As String
'-----------Para el documento de referencia
Private strCodDocRef As String
Private strNumDocRef As String
Private strBtlRef As String
Private strFlgNcrDev As String
'---------- Para los conceptos y subconceptos de nc
Private strConcepto As String
Private strSubConcepto As String
'-----------para delivery
Private strNumPedidoPadre As String
Private strFlgTransfLocal As String
Private strFlgEntregaLocal As String
'Private strFlgPedUrgente As String
Private strFlgReservaStk As String
Private strCodDireccionCli As String
Private strCodMotorizado As String
Private lMontoRedime As String
Private bEsVentaMonedero As Boolean
Private sNumeroTarjetaMonedero As String
Private mvarPuntosTarjetaMonedero As Double
Private strFlgUrgente As String
Private strFlgPactado As String
Private strDesAuxCliTlf As String
Private strDesAuxCliNombre As String
Private flgEntregaTercero As String
Private strDesAuxCliDirecc As String
Private strDesReferenciaCli As String
Private datFchHoraPactEntr As String
Private datFchHoraPactRecog As String
Private strDesAuxRecogeTlf As String
Private strDesAuxRecogeNombre As String
Private strDesAuxRecogeDirecc As String
Private strDesAuxRecogeRef As String
Private strDesAuxMotorizado As String
Private strObsNotaVerificacion As String
Private strObsNotaRuteo As String
Private strObsNotaLocal As String
Private strObsNotaMotorizado As String
Private strHrPactEntrada As String
Private strHrPactRecojo As String
Private strHoraPactEntr As String
Private strHoraPactRecogo As String
'------------cobro por responsabilidad
Private strCodMotivoCobro As String
'-------------cliente
Private strOut_CodigoCliente As String
Private strOut_NombreCliente As String
Private strOut_NumeroId As String
Private strOut_Tipo As String
Private strOut_Direccion As String
Private strOut_CodDocumentoId As String
Private strNumeroReceta As String
Private strFechaReceta As String
Private strLocalReceta As String
'----------- cliente DLV
Private strNombreClienteDLV  As String
Private strDireccionClienteDLV As String
Private strDesDistritoDLV As String
Private strDesUrbanizacionDLV As String
Private strObservacionClienteDLV As String
Private strLineaCred As String
Private strConsumo As String

Private strBenef As String
Private strUbigeo As String

Private dblNewPctBenef As Double
Private strObservacion As String
'------------


Private strDevEfectivoNC As String
'----------- Datos del Titular de la tarjeta
Private strNomTitular As String
Private strNumDNI As String
Private strCodMotorizadoConv As String


Private strSerieTKT As String
Private strSloganBTL As String

Public Enum TipoVenta
    Venta_Regular = 1
    Venta_Convenio = 2
    Venta_Mayorista = 3
    Cobro_Responsabilidad = 4
    Canje_Canje = 5
    Servicio = 6
    Cobranza_VtaCred = 7
    Cotizaciones = 8
    FactServicios = 9
    Guias_Remision = 10
    Recetario = 17
    Pedido_DLV = 12
    Cajero_Corresponsal = 19
End Enum

Public ptmModalidad As TipoVenta

'** Variables para el convenio 30/12/2008 **'
Dim strFlgImprimeImp As String
Dim strFlgPrecioMenor As String
Dim strFlgPrecioDeducible As String

Dim dblPrecioPublico As Double
Private m_sFlgPlanVital As String
Private m_sNomArchVital As String
Private m_sFecAteVital As String
Private m_sNumAteVital As String
Private m_sCodCentralVital As String
Private m_sNumPedidoVital As String
Private m_sCodPostalVital As String
Private m_sDesRefVital As String
Private m_sFonoBenefVital As String
Private m_sSexoBenefVital As String
Private m_dCtdPedidoVital As Double
Private m_sNumPolizaVital As String
Private m_sNumCertifVital As String
Private m_sCodParentesco As String
Private m_sCodDeducibleVital As String
Private m_sTipoFacturaVital As String
Private m_sFecNacVital As String
Private m_sCodInternoVital As String
Private m_sDesObsVital As String
Private m_dPctCoPagoVital As Double
Private m_sNombBenef As String
Private m_sApePatBenef As String
Private m_sApeMatBenef As String
Private m_sDirBenef As String
Private m_sArchivoVital As String

Private strCodProducto_BusGen As String
Private strDesProducto_BusGen As String

' Para indicar si va dar o no comision

Public vExisteDNI_RENIEC As String
Dim vstrStock1 As String

'Para ver si excede el stock 15/10/2012 MLEVANO
Public evaluaStockGuiaLote As Integer



Private strLatitud As String
Private strLongitud As String
'
Public strTipoDocumento_Pts As String
Public strNumDocumento_Pts As String
Public ObservacionLocal As String
'I.ECASTILLO 27.10.2020
Public isLocalDcCappa As String
Public flgDatosCapacidad As Boolean
Public flgReservoCap As String
Public flgDiasFuturos As String
Public dc_cod_forma_pago As String
Public dc_street As String
Public dc_number As String
Public dc_apartment As String
Public dc_country As String
Public dc_city As String
Public dc_district As String
Public dc_latitude As String
Public dc_longitude As String
Public dc_receiverName As String
Public dc_orderDigital As String
Public dc_departamentBK As String
Public dc_referencia As String
Public dc_urbanizacion As String
'F.ECASTILLO 27.10.2020
'I.ECASTILLO 17.12.2020
Public respetaLocal As Boolean
Public bk_codBeneficiario As String
Public flg_2e_reserva_local As String
Public bk_strProforma As String
Public bk_codLocal As String
Public bk_codLocalCapacidad As String
Public bk_FechaCapacidad As String
Public bk_HoraCapacidad As String
Public bk_HoraCapacidad2 As String
Public bk_ServiceType As String
Public bk_chkRET As String
Public bk_flgPactado As String
Public bk_deliveryTime As String
Public bk_codCliente As String
Public bk_Ubigeo As String
Public bk_message As String
Public bk_segmento As String
Public bk_amount As String
Public bk_starHour As String
Public bk_endHour As String
'F.ECASTILLO 17.12.2020
Public bk_DescSufijoDir As String
Public bk_AbrSufijoDir As String
Public bk_SufijoDir As String

Public busquedaNVeces As Integer
Public strBusqueda As String
Public isDCSAP As String





'*************************************************************************************
' Cuando se inicializa la clase se crean todas las variables
'*************************************************************************************
Private Sub Class_Initialize()
    xMaquina.ReDim 0, -1, 0, 7
    xDiagnostico.ReDim 0, -1, 0, 3
    'I.CVIERA 31.12.2020
    'xProducto.ReDim 0, -1, 0, 29 'Inicializa los productos seleccionados
    'xProducto.ReDim 0, -1, 0, 30 'Inicializa los productos seleccionados
    'F.CVIERA 31.12.2020
    'I.ECASTILLO 17.12.2020
    xProducto.ReDim 0, -1, 0, 31
    'F.ECASTILLO 17.12.2020
    xProductoRegalo.ReDim 0, -1, 0, 29 'Inicializa los productos regalos seleccionados
    xProductoRegaloBK.ReDim 0, -1, 0, 29
    XDistribucion.ReDim 0, -1, 0, 15 'Inicializa los productos seleccionados
    xTransferencia.ReDim 0, -1, 0, 15 'Inicializa el arreglo de transferencia
    xConvenio.ReDim 0, -1, 0, 2 'Inicializa el array de convenios
    xFormaPago.ReDim 0, -1, 0, 35 '33 'Inicializa el array de Forma Pago
    xCobroResponsabilidad.ReDim 0, -1, 0, 3 'Inicializa el array de Conbro por Responsabilidad
    xRecetarioMagistral.ReDim 0, -1, 0, 9
    xDatoAdicional.ReDim 0, -1, 0, 6
    xProductoCobro.ReDim 0, -1, 0, 11 'Inicializa el array de los productos para el cobro por responasibilidad
    xServicio.ReDim 0, -1, 0, 13 'Inicializa la variable servicios
    xEspeciesValoradas.ReDim 0, -1, 0, 5
    xProductoLote.ReDim 0, -1, 0, 27
    
    xProductoLoteDetalle.ReDim 0, -1, 0, 23
    
    xArrayPlanVital.ReDim 0, -1, 0, 3 'Setea el array de la carga de archivo txt plan vital
    Me.CodigoDocumentoVenta = objUsuario.TipDocDefault
    'I.CVIERA 31.12.2020
    'xProductoBK.ReDim 0, -1, 0, 30
    'xProductoBK.ReDim 0, -1, 0, 31
    'F.CVIERA 31.12.2020
    'I.ECASTILLO 17.12.2020
    xProductoBK.ReDim 0, -1, 0, 32
    'F.ECASTILLO 17.12.2020
    xCuantaMaxima.ReDim 0, -1, 0, 2 'ECASTILLO 04.07.2020
    
    xMetodoSegmento.ReDim 0, -1, 0, 11
End Sub

Public Property Get Longitud() As String
    Longitud = Trim(strLongitud)
End Property
Public Property Let Longitud(str As String)
    strLongitud = Trim(str)
End Property
Public Property Get Latitud() As String
    Latitud = Trim(strLatitud)
End Property
Public Property Let Latitud(str As String)
    strLatitud = Trim(str)
End Property


'**** Proiedad que permite devolver el arreglo para las maquinas ****'
Public Property Get Maquina() As XArrayDB
    Set Maquina = xMaquina
End Property
'********************************************************************'

Public Property Get EntregaTerceros() As XArrayDB
    Set EntregaTerceros = xEntregaTerceros
End Property

Public Property Get Transferencia() As XArrayDB
    Set Transferencia = xTransferencia
End Property


'*************************************************************************************
' Propiedad que contiene todas los productos que se han seleccionado para la venta
'*************************************************************************************

Public Sub limpiarProducto()
    xProducto.Clear
    'I.CVIERA 31.12.2020
    'xProducto.ReDim 0, -1, 0, 29
    'xProducto.ReDim 0, -1, 0, 30
    'F.CVIERA 31.12.2020
    'I.ECASTILLO 17.12.2020
    xProducto.ReDim 0, -1, 0, 31
    'F.ECASTILLO 17.12.2020
    xProductoRegalo.Clear
    xProductoRegalo.ReDim 0, -1, 0, 29
    xProductoRegaloBK.Clear
    xProductoRegaloBK.ReDim 0, -1, 0, 29
    xProductoBK.Clear
    'I.CVIERA 31.12.2020
    'xProductoBK.ReDim 0, -1, 0, 30
    'xProductoBK.ReDim 0, -1, 0, 31
    'F.CVIERA 31.12.2020
    'I.ECASTILLO 17.12.2020
    xProductoBK.ReDim 0, -1, 0, 32
    'F.ECASTILLO 17.12.2020
    xCuantaMaxima.Clear 'ECASTILLO 04.07.2020
    xCuantaMaxima.ReDim 0, -1, 0, 2
    xMetodoSegmento.Clear
    xMetodoSegmento.ReDim 0, 1, 0, 11
End Sub
Public Property Get MetodoSegmento() As XArrayDB
    Set MetodoSegmento = xMetodoSegmento
End Property

Public Property Get Producto() As XArrayDB
    Set Producto = xProducto 'Devuelve la variable de la clase en forma de XArray
End Property
Public Property Get Producto2() As XArrayDB
    Call ObtieneRegalos
    Set Producto2 = xProducto
End Property
Public Property Get EspeciesValoradas() As XArrayDB
    Set EspeciesValoradas = xEspeciesValoradas 'Devuelve la variable de la clase en forma de XArray
End Property
Public Property Get ProductoLote() As XArrayDB
    Set ProductoLote = xProductoLote
End Property

Public Property Get ProductoLoteDetalle() As XArrayDB
    Set ProductoLoteDetalle = xProductoLoteDetalle
End Property

Public Property Get Distribucion() As XArrayDB
    If strProductoTransferido = "" Then
        Set Distribucion = XDistribucion 'Devuelve la variable de la clase en forma de XArray
    Else
        Dim xTemp As New XArrayDB
        Dim i As Integer
        i = 0
        xTemp.ReDim 0, -1, 0, 15 'Inicializa los productos seleccionados
        While i < XDistribucion.Count(1)
            If XDistribucion(i, 1) = strProductoTransferido Then
                Dim j As Integer
                j = xTemp.Count(1)
                xTemp.AppendRows
                Dim t As Integer
                    For t = 0 To 9
                        xTemp(j, t) = XDistribucion(i, t)
                    Next t
            End If
            i = i + 1
        Wend
        Set Distribucion = xTemp
    End If
End Property
Public Property Let ProductoTransferido(mvar As String)
    strProductoTransferido = mvar
End Property
Public Property Get FormaPago() As XArrayDB
    Set FormaPago = xFormaPago 'Devuelve la variable de la clase en forma de XArray
End Property
Public Property Get CobroResponsabilidad() As XArrayDB
    Set CobroResponsabilidad = xCobroResponsabilidad 'Devuelve la variable de la clase en forma de XArray
End Property

Public Property Get RecetarioMagistral() As XArrayDB
    Set RecetarioMagistral = xRecetarioMagistral
End Property

Public Property Get ProductoCobro() As XArrayDB
    Set ProductoCobro = xProductoCobro
End Property

Public Property Get Servicios() As XArrayDB
On Error GoTo handle
    Set Servicios = xServicio
    Exit Function
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Property


Public Function AgregaEspecieValorada(codigoProducto As String, _
                            NumeroFormulario As String, _
                            placa As String _
                            ) As XArrayDB
On Error GoTo handle
    If codigoProducto = "" Then Err.Raise 1500, "clsVenta", "Debe de ingresar un codigo de Producto"
    If NumeroFormulario = "" Then Err.Raise 1500, "clsVenta", "Debe de ingresar el número de formulario"
    If placa = "" Then Err.Raise 1500, "clsVenta", "Debe de ingresar la placa del formulario"
    
    Dim ultimo As Integer ' declara variable contador
    xEspeciesValoradas.AppendRows
    ultimo = xEspeciesValoradas.Count(1) - 1
    xEspeciesValoradas(ultimo, 0) = codigoProducto ' Asigna el codigo de producto
    xEspeciesValoradas(ultimo, 1) = NumeroFormulario ' Asigna la descripcion del producto
    xEspeciesValoradas(ultimo, 2) = placa ' Asigna si la cantidad es fraccion o Unidad
    Set AgregaEspecieValorada = xEspeciesValoradas
    Exit Function
handle:
    Set AgregaEspecieValorada = xEspeciesValoradas
    Err.Raise Err.Number, "clsVenta", Err.Description
End Function

Public Function AgregaEntregaTerceros(ByVal Nombre As String, _
                            ByVal Direccion As String, _
                            ByVal Referencia As String, _
                            ByVal Telefono As String) As XArrayDB
On Error GoTo handle
    xEntregaTerceros.ReDim 0, -1, 0, 3  'Inicializa el array para pedidos con entrega a terceros
    Dim ultimo As Integer
    xEntregaTerceros.AppendRows
    ultimo = xEntregaTerceros.Count(1) - 1
    xEntregaTerceros(ultimo, 0) = Nombre
    xEntregaTerceros(ultimo, 1) = Direccion
    xEntregaTerceros(ultimo, 2) = Referencia
    xEntregaTerceros(ultimo, 3) = Telefono
    Set AgregaEntregaTerceros = xEntregaTerceros
    Exit Function
handle:
    Err.Raise Err.Number, "clsVentaAgregaEntregaTerceros", Err.Description
    Set xEntregaTerceros = Nothing
End Function


'*************************************************************************************
' Metodo para seleccionar un producto para la venta, retorna como variable el propio array
'*************************************************************************************
Public Function AgregaProducto(Codigo As String, _
                               Descripcion As String, _
                               cantidad As Integer, _
                               FlagFraccion As String, _
                               Precio As Double, _
                               TipoVenta As TipoVenta, _
                               Regalo As eProdRegalo, _
                               Optional Dato1 As String = "", _
                               Optional Dato2 As String = "", _
                               Optional Dato3 As String = "", _
                               Optional Dato4 As String = "", _
                               Optional Dato5 As String = "", _
                               Optional IndicadorReceta As String = "", _
                               Optional pctComision As Double = 0, _
                               Optional NroLote As String = "", _
                               Optional FchVmto As String = "", _
                               Optional flgFarmaco As String = "0", _
                               Optional PrecPublic As Double, Optional CadenaPromocion As String = "", Optional CadenaPromocionDcto As String = "", Optional FracEnabled As String = "", Optional CantPuntos As String, Optional CantPuntosRedimir As Double, Optional flg_seg As String, Optional flg_Transf As String, Optional flg_valid_frac As String) As XArrayDB
    
    Dim objProducto As New clsProducto
    
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If xProducto.Count(1) < 0 Then Exit Function
        
    
    Dim i As Integer
    i = 0
    Dim encontro As Boolean
    Dim PctComi As Double

    Dim oldCant As Integer
    Dim oldFrac As String
    
    Dim ultimoDet As Integer
    Dim ultimoCant As Integer
    Dim aux2 As Integer
    
    Dim j As Integer
    Dim k As Integer
    
    Dim canti As Double
    Dim prec As Double
    Dim FraccionaEnabled As String
    Dim cantidadFraccionable As Integer
    Dim flgFraccion As String
    Dim strPromocionBK As String
    'Autor: Juan Arturo Escate Espichan
    'Fecha: 14/04/2008
    'Proposito: esto es para que no se confunda navsat , para que sea solo un documento
    
''''''    If xProducto.Count(1) <> 0 Then
''''''        Dim mi As Integer
''''''        mi = xProducto.Count(1)
''''''        Dim b As Integer
''''''        b = 0
''''''        Dim numServ As Integer
''''''        Dim numOtros As Integer
''''''        numServ = 0
''''''        numOtros = 0
''''''        While b < mi
''''''            If Val(xProducto(b, 5)) = Val(servicio) Then
''''''                numServ = numServ + 1
''''''            Else
''''''                numOtros = numOtros + 1
''''''            End If
''''''            b = b + 1
''''''
''''''        Wend
''''''        If numServ > 0 And numOtros <> o Then
''''''            MsgBox "No se permite cobrar servicios con venta de producto", vbCritical, App.ProductName
''''''            Exit Function
''''''        End If
''''''            '
''''''    End If
    '-------------------------------------------------------------------------------------
    
    '''''''''''''PctComi = objProducto.pctComision(Codigo, objUsuario.CodigoLocal, Format(ptmTipoPrecio, "000"))
    
    PctComi = pctComision
    
    aux = xProducto.Count(1)
    'MsgBox xProducto.Count(1)
    'MsgBox xProductoLote.Count(1)
'    ========================AGREGADO PARA LOTES POR PRODUCTO========== MLEVANO 28/08/2012
    If strCodigoTipoVenta = 10 Then
        While i < aux
            If xProducto(i, 0) = Codigo Then 'And xProducto(i, 7) = Regalo And xProducto(i, 22) = NroLote Then
                ultimo = i
                encontro = True
GoTo m
            Else
                encontro = False
                ultimo = xProducto.Count(1) 'Llena con la ultima posicion disponible
            End If
        i = i + 1
        Wend
        
        If encontro = False Then
            xProducto.AppendRows
        End If
m:
        If xProducto.Count(1) = 0 Then ultimo = 0: xProducto.AppendRows
        
        aux2 = xProductoLote.Count(1)
            While j < aux2
                If xProductoLote(j, 0) = Codigo And xProducto(i, 22) = NroLote Then
                    ultimoDet = j
                    encontro = True
GoTo n
                Else
                    encontro = False
                    ultimoDet = xProductoLote.Count(1) - 1 'Llena con la ultima posicion disponible
                End If
                j = j + 1
            Wend
n:
    
        cantidadFraccionable = objProducto.DevFraccionProducto(Codigo)
        If aux2 <> 0 Then
            While k < aux2
                If xProductoLote(k, 0) = Codigo Then
                    If xProductoLote(k, 2) = "F" Then
                        canti = canti + xProductoLote(k, 3)
                        prec = prec + xProductoLote(k, 4)
                    Else
                        canti = canti + (xProductoLote(k, 3) * cantidadFraccionable)
                        prec = prec + xProductoLote(k, 4)
                    End If
                End If
                
                k = k + 1
            Wend
        Else
            While k < xProducto.Count(1)
                If xProducto(k, 0) = Codigo Then
                    If xProducto(k, 2) = "F" Then
                        canti = canti + xProducto(k, 3)
                        prec = prec + xProducto(k, 4)
                    Else
                        canti = canti + (xProducto(k, 3) * cantidadFraccionable)
                        prec = prec + xProducto(k, 4)
                    End If
                End If
                k = k + 1
            Wend
        End If
        
        k = 0
            
                If canti Mod cantidadFraccionable = 0 Then
                   canti = canti / cantidadFraccionable
                    FraccionaEnabled = "U"
                    flgFraccion = "0"
                Else
                    canti = canti
                    If FracEnabled = "1" Then
                        FraccionaEnabled = "F"
                       flgFraccion = "1"
                    ElseIf FracEnabled = "0" Then
                        FraccionaEnabled = "U"
                        flgFraccion = "0"
                   Else
                        FraccionaEnabled = xProducto(ultimo, 2)
                        flgFraccion = xProducto(ultimo, 6)
                    End If
                End If
              
               
      

            
    
'    ===========================================================================
    Else
'    ===========================================================================
        
        
        'gclsOracle.FN_Valor("BTLPROD.pkg_usuario.FN_ES_QUIMICO", txtUsuario.Text)
        If flg_valid_frac = "0" Or flg_valid_frac = "" Then
            While i < aux
                'If xProducto(i, 0) = Codigo And xProducto(i, 5) = TipoVenta Then
                If xProducto(i, 0) = Codigo And xProducto(i, 7) = Regalo Then
                    ultimo = i
                    encontro = True
                    GoTo j
                Else
                    encontro = False
                    ultimo = xProducto.Count(1) 'Llena con la ultima posicion disponible
                End If
            i = i + 1
            Wend
        Else
            While i < aux
                'If xProducto(i, 0) = Codigo And xProducto(i, 5) = TipoVenta Then
                If xProducto(i, 0) = Codigo And xProducto(i, 7) = Regalo And xProducto(i, 6) = FlagFraccion Then
                    ultimo = i
                    encontro = True
                    GoTo j
                Else
                    encontro = False
                    ultimo = xProducto.Count(1) 'Llena con la ultima posicion disponible
                End If
            i = i + 1
            Wend
        End If
    
    
        If encontro = False Then
            xProducto.AppendRows
        End If
j:
        If xProducto.Count(1) = 0 Then ultimo = 0: xProducto.AppendRows
        
        If xProducto(ultimo, 14) = 0 Then
            oldCant = val(cantidad)
        Else
            oldCant = xProducto(ultimo, 3)
        End If
        If xProducto(ultimo, 15) = "" Then
            oldFrac = FlagFraccion
        Else
            oldFrac = IIf(xProducto(ultimo, 2) = "F", "1", "0")
        End If
    
    End If
        strPromocionBK = IIf(CadenaPromocion <> "", CadenaPromocion, xProducto(ultimo, 26))
        xProducto(ultimo, 0) = Codigo ' Asigna el codigo de producto
        xProducto(ultimo, 1) = Descripcion ' Asigna la descripcion del producto
        If strCodigoTipoVenta = 10 Then
            xProducto(ultimo, 2) = FraccionaEnabled
        Else
            xProducto(ultimo, 2) = IIf(FlagFraccion = "1", "F", "U") ' Asigna si la cantidad es fraccion o Unidad
        End If
        
        If strCodigoTipoVenta = 10 Then
            xProducto(ultimo, 3) = val(canti) ' Asigna la cantidad seleccionada del producto
            xProducto(ultimo, 4) = Format(prec, "###,##0.00") ' Asigna el tipo de Venta
        Else
            xProducto(ultimo, 3) = val(cantidad) ' Asigna la cantidad seleccionada del producto
            xProducto(ultimo, 4) = Format(Precio, "###,##0.00") ' Asigna el tipo de Venta
        End If
    
        xProducto(ultimo, 5) = TipoVenta ' Asigna el tipo de Venta
        
        If strCodigoTipoVenta = 10 Then
            xProducto(ultimo, 6) = flgFraccion
        Else
            xProducto(ultimo, 6) = FlagFraccion ' Asigna si la cantidad es fraccion o Unidad
        End If
        
        xProducto(ultimo, 7) = Regalo 'valor libre
        xProducto(ultimo, 8) = ""
        xProducto(ultimo, 9) = 0
        xProducto(ultimo, 10) = ""
        xProducto(ultimo, 11) = ""
        xProducto(ultimo, 12) = ""
        xProducto(ultimo, 13) = PctComi
        
        If Regalo = Producto_Normal Then
            xProducto(ultimo, 14) = 0
            xProducto(ultimo, 15) = 0
        Else
            xProducto(ultimo, 14) = oldCant
            xProducto(ultimo, 15) = oldFrac
        End If
        xProducto(ultimo, 16) = Dato1 'Formulario de SOAT o Especie Valorada
        xProducto(ultimo, 17) = Dato2 'Número de placa o de referencia de la especie Valorada
        xProducto(ultimo, 18) = Dato3
        xProducto(ultimo, 19) = Dato4
        xProducto(ultimo, 20) = Dato5
        xProducto(ultimo, 21) = IndicadorReceta
        xProducto(ultimo, 22) = NroLote
        xProducto(ultimo, 23) = FchVmto
        xProducto(ultimo, 24) = flgFarmaco
        xProducto(ultimo, 25) = Format(PrecPublic, "###,##0.00") 'Precio Publico
        xProducto(ultimo, 26) = strPromocionBK 'CadenaPromocion 'Cadena de Promociones 'ECASTILLO 04.07.2020
        xProducto(ultimo, 27) = CadenaPromocionDcto 'Cadena de descuento por promocion
        xProducto(ultimo, 28) = val(CantPuntos) 'cantidad de puntos por product
        xProducto(ultimo, 29) = val(CantPuntosRedimir)
        Dim flg_2e_reserva
        flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
        If flg_2e_reserva = "0" Then
        Else
            'I CVIERA 31.12.2020
            If val(flg_seg) <> 0 Then
                xProducto(ultimo, 30) = val(flg_seg)
            End If
            'F. CVIERA 31.12.2020
            'I.ECASTILLO 17.12.2020
            If Len(Trim(flg_Transf)) > 0 Then
                xProducto(ultimo, 31) = val(flg_Transf)
            End If
            'F.ECASTILLO 17.12.2020
        End If
'        If Val(cantidad) = 0 Then 'ECASTILLO 04.07.2020
'            EliminaProducto Codigo, TipoVenta
'        End If
        
        If strCodigoTipoVenta = 10 Then
            If xProducto(ultimo, 7) = "1" Then
                xProducto.DeleteRows (ultimo)
            End If
            
            If Regalo <> Producto_Regalo Then
                If objProducto.EvaluaCtdVta(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, Codigo, canti, flgFraccion) = "0" Then
                   vstrStock1 = objProducto.Stock(Codigo, objUsuario.CodigoLocal, "1")
                   MsgBox "La cantidad solicitada es mayor a la existente" & Chr(13) & "Stock Actual: " & vstrStock1, vbExclamation + vbOKOnly, App.ProductName
                   Exit Function
                End If
            End If
                
            
        End If
        
    '    If tipoventa = 3 Then
    '        Dim odyn As oraDynaset
    '        Set odyn = objProducto.ListaExcluyeMayorista(codigo)
    '        odyn.MoveFirst
    '        While Not odyn.EOF
    '            If odyn("COD_PRODUCTO") = codigo Then
    '                EliminaProducto2 codigo, tipoventaG
    '                Exit Function
    '            End If
    '            odyn.MoveNext
    '        Wend
    '    End If
    Call ReasignaRegalos(xProducto, strPromocionBK)
    Set AgregaProducto = xProducto
End Function
Public Function ReasignaRegalos(Xarray As XArrayDB, strPromocionBK As String)
    'PROCESA SOLO REGALOS, AGREGANDOLOS EN UN NUEVO ARRAY CON UNA COLUMNA NUEVA
    'CODIGO PADRE, QUE INDICA CUAL ES EL PETITORIO, SE AGREGAN SOLO REGALOS CON CANT 0
    Dim i, x As Integer
    While i < xProducto.Count(1)
        Debug.Print "=== " & xProducto(i, 0) & "|" & xProducto(i, 3) & "|" & xProducto(i, 7) & "|" & xProducto(i, 26) & "|" & xProducto(i, 30) & "|" & xProducto(i, 31)
        i = i + 1
    Wend
    i = 0
    While i < xProducto.Count(1)
        If xProducto(i, 7) <> 0 Then 'SOLO REGALOS
            Dim indice As Integer
            Dim existe As Boolean
            x = 0
            existe = False
            While x < xProductoBK.Count(1) And existe = False
                'SI EXISTE REGALO,Y CANTIDAD > 0, REMUEVE DE ARRAY
                If xProductoBK(x, 0) = xProducto(i, 0) And xProductoBK(x, 7) = xProducto(i, 7) And xProducto(i, 3) <> 0 Then
                    indice = x
                    existe = True
                ElseIf xProductoBK(x, 0) = xProducto(i, 0) And xProductoBK(x, 3) = xProducto(i, 3) Then
                    indice = x
                    existe = True
                End If
                x = x + 1
            Wend
            If existe = True And xProducto(i, 3) <> 0 Then: xProductoBK.DeleteRows (indice)
            'SI NO EXISTE, REGISTRA EN LA ULTIMA POSICION DISPONIBLE
            If existe = False And xProducto(i, 3) = 0 Then
                Dim j As Integer
                j = 0
                xProductoBK.AppendRows
                indice = xProductoBK.UpperBound(1)
                While j < xProducto.Count(2)
                    If j = 26 Then
                        xProductoBK(indice, j) = strPromocionBK
                    Else
                        xProductoBK(indice, j) = xProducto(i, j)
                    End If
                    j = j + 1
                Wend
                j = 0
                While j < xProducto.Count(1)
                    If xProducto(j, 7) = 0 And xProducto(j, 26) = xProductoBK(indice, 26) Then
                        'I. CVIERA 31.12.2020
'                        xProductoBK(indice, 30) = xProducto(j, 0) 'NUEVO INDICE PARA RELACIONAR CON SU PETITORIO
                        'xProductoBK(indice, 31) = xProducto(j, 0) 'NUEVO INDICE PARA RELACIONAR CON SU PETITORIO
                        'F. CVIERA 31.12.2020
                        'I.ECASTILLO 17.12.2020
                        xProductoBK(indice, 32) = xProducto(j, 0) 'NUEVO INDICE PARA RELACIONAR CON SU PETITORIO
                        'F.ECASTILLO 17.12.2020
                    End If
                    j = j + 1
                Wend
                'LUEGO DE AGREGARLO AL NUEVO ARRAY, SE REMUEVE DEL ORIGINAL
                xProducto.DeleteRows (i)
                i = i - 1
            End If
            If existe = True Then
                If xProducto(i, 3) = 0 Then: xProducto.DeleteRows (i): i = i - 1
            End If
        ElseIf xProducto(i, 3) = 0 And xProducto(i, 7) = 0 Then
            'CUANDO SE INGRESA CANT 0 EN EL PETITORIO ELIMINA LOS REGALOS
            'SE DETERMINA EL REGALO CON LA COLUMNA 30 QUE CONTIENE EL CODIGO DEL PRODUCTO (PETITORIO)
            'Debug.Print "### " & xProductoBK(0, 0) & "|" & xProductoBK(0, 3) & "|" & xProductoBK(0, 7) & "|" & xProductoBK(0, 26) & "|" & xProductoBK(0, 30)
            x = 0
            While x < xProductoBK.Count(1)
'                If xProductoBK(x, 30) = xProducto(i, 0) Or xProductoBK(x, 26) = strPromocionBK Or xProductoBK(x, 30) = "" Then xProductoBK.DeleteRows (x) 'ECASTILLO 04.07.2020
                'I. CVIERA 31.12.2020
                'If xProductoBK(x, 31) = xProducto(i, 0) Or xProductoBK(x, 26) = strPromocionBK Or xProductoBK(x, 31) = "" Then xProductoBK.DeleteRows (x) 'ECASTILLO 04.07.2020
                'F. CVIERA 31.12.2020
                'I.ECASTILLO 17.12.2020
                If xProductoBK(x, 32) = xProducto(i, 0) Or xProductoBK(x, 26) = strPromocionBK Or xProductoBK(x, 32) = "" Then xProductoBK.DeleteRows (x) 'ECASTILLO 04.07.2020
                'F.ECASTILLO 17.12.2020
                x = x + 1
            Wend
            x = 0
            While x < xProductoRegalo.Count(1)
                If xProductoRegalo(x, 18) = strPromocionBK Then: xProductoRegalo.DeleteRows (x): x = x - 1
                x = x + 1
            Wend
            x = 0
            While x < xProductoRegaloBK.Count(1)
                If xProductoRegaloBK(x, 18) = strPromocionBK Then: xProductoRegaloBK.DeleteRows (x): x = x - 1
                x = x + 1
            Wend
        End If
        i = i + 1
    Wend
    
    i = 0
    'Debug.Print vbNewLine
'    While i < xProducto.Count(1)
'        Debug.Print "*** " & xProducto(i, 0) & "|" & xProducto(i, 3) & "|" & xProducto(i, 7) & "|" & xProducto(i, 26)
'        i = i + 1
'    Wend
'    i = 0
'    Debug.Print vbNewLine
'    While i < xProductoBK.Count(1)
'        Debug.Print "=== " & xProductoBK(i, 0) & "|" & xProductoBK(i, 3) & "|" & xProductoBK(i, 7) & "|" & xProductoBK(i, 26) & "|" & xProductoBK(i, 30)
'        i = i + 1
'    Wend
End Function
Public Function ObtieneRegalos()
    Dim i, x, j As Integer
    For i = 0 To xProductoBK.UpperBound(1)
        Dim indice As Integer
        Dim existe As Boolean
        x = 0
        While x < xProducto.Count(1) And existe = False
            'SI EXISTE REGALO,Y CANTIDAD > 0, REMUEVE DE ARRAY
            If xProducto(x, 0) = xProductoBK(i, 0) And xProducto(x, 7) = xProductoBK(i, 7) Then
                indice = x
                existe = True
            ElseIf xProducto(x, 26) = xProductoBK(i, 26) And xProducto(x, 7) = xProductoBK(i, 7) Then 'ECASTILLO 04.07.2020
                indice = x
                existe = True
            End If
            x = x + 1
        Wend
        If existe = False Then 'ECASTILLO 04.07.2020
            xProducto.AppendRows
            indice = xProducto.UpperBound(1)
            j = 0
            'I CVIERA 31.12.2020
'            While j < xProductoBK.Count(2) And j <> 30
            'While j < xProductoBK.Count(2) And j <> 31
            'F CVIERA 31.12.2020
            'I.ECASTILLO 17.12.2020
            While j < xProductoBK.Count(2) And j <> 32
            'F.ECASTILLO 17.12.2020
                xProducto(indice, j) = xProductoBK(i, j)
                j = j + 1
            Wend
            
        End If
    Next i
    i = 0
'    Debug.Print vbNewLine
    While i < xProducto.Count(1)
        Debug.Print "+++ " & xProducto(i, 0) & "|" & xProducto(i, 3) & "|" & xProducto(i, 7) & "|" & xProducto(i, 26) & "|" & xProducto(i, 30)
        i = i + 1
    Wend
End Function

'*************************************************************************************
' Metodo para seleccionar un producto para la venta, retorna como variable el propio array
'*************************************************************************************
Public Function LimpiaDistribucion()
    XDistribucion.ReDim 0, -1, 0, 15 'Inicializa los productos seleccionados
End Function
Public Function AgregaDistribucion(CodigoLocal As String, _
                            Codigo As String, _
                            Descripcion As String, _
                            cantidad As String, _
                            FlagFraccion As String, _
                            Precio As Double, _
                            TipoVenta As TipoVenta, _
                            CodigoLocalOrigen As String, _
                            CodigoLocalDestino As String, _
                            CodigoLocalSapOrigen As String, _
                            CodigoLocalSapDestino As String _
                            ) As XArrayDB
                            
    
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If CodigoLocalOrigen = "" Then Err.Raise 200, "clsVenta", "El codigo de local es vacio"
    
    If XDistribucion.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    aux = XDistribucion.Count(1)
    While i < aux
        If XDistribucion(i, 8) = CodigoLocalOrigen And XDistribucion(i, 1) = Codigo And XDistribucion(i, 7) = FlagFraccion Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = XDistribucion.Count(1) 'Llena con la ultima posicion disponible
        End If
        i = i + 1
    Wend
    If encontro = False Then
        XDistribucion.AppendRows
    End If
        
j:
    If XDistribucion.Count(1) = 0 Then ultimo = 0: XDistribucion.AppendRows
    XDistribucion(ultimo, 0) = CodigoLocal 'objUsuario.CodigoLocal ' Asigna el codigo de Local
    XDistribucion(ultimo, 1) = Codigo ' Asigna el codigo de producto
    XDistribucion(ultimo, 2) = Descripcion ' Asigna la descripcion del producto
    XDistribucion(ultimo, 3) = IIf(FlagFraccion = "1", "F", "U") ' Asigna si la cantidad es fraccion o Unidad
    XDistribucion(ultimo, 4) = cantidad ' Asigna la cantidad seleccionada del producto
    XDistribucion(ultimo, 5) = Format(Precio, "###,##0.00") ' Asigna el tipo de Venta
    XDistribucion(ultimo, 6) = TipoVenta ' Asigna el tipo de Venta
    XDistribucion(ultimo, 7) = FlagFraccion ' Asigna si la cantidad es fraccion o Unidad
    XDistribucion(ultimo, 8) = CodigoLocalOrigen  ' Asigna el codigo de Local
    XDistribucion(ultimo, 9) = CodigoLocalDestino ' Asigna el codigo de Local
    XDistribucion(ultimo, 10) = CodigoLocalSapOrigen ' Asigna el codigo de Local SAP
    XDistribucion(ultimo, 11) = CodigoLocalSapDestino ' Asigna el codigo de Local SAP
    Set AgregaDistribucion = XDistribucion
    
End Function

Public Function AgregaTransferencia(CodigoLocal As String, _
                            Codigo As String, _
                            Descripcion As String, _
                            cantidad As String, _
                            FlagFraccion As String, _
                            Precio As Double, _
                            TipoVenta As TipoVenta, _
                            CodigoLocalOrigen As String, _
                            CodigoLocalDestino As String, _
                            CodigoLocalSapOrigen As String, _
                            CodigoLocalSapDestino As String _
                            ) As XArrayDB
                            
    
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If CodigoLocalOrigen = "" Then Err.Raise 200, "clsVenta", "El codigo de local es vacio"
    
    If xTransferencia.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    aux = xTransferencia.Count(1)
    While i < aux
        If xTransferencia(i, 8) = CodigoLocalOrigen And xTransferencia(i, 1) = Codigo And xTransferencia(i, 7) = FlagFraccion Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xTransferencia.Count(1) 'Llena con la ultima posicion disponible
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xTransferencia.AppendRows
    End If
        
j:
    If xTransferencia.Count(1) = 0 Then ultimo = 0: xTransferencia.AppendRows
    xTransferencia(ultimo, 0) = CodigoLocal 'objUsuario.CodigoLocal ' Asigna el codigo de Local
    xTransferencia(ultimo, 1) = Codigo ' Asigna el codigo de producto
    xTransferencia(ultimo, 2) = Descripcion ' Asigna la descripcion del producto
    xTransferencia(ultimo, 3) = IIf(FlagFraccion = "1", "F", "U") ' Asigna si la cantidad es fraccion o Unidad
    xTransferencia(ultimo, 4) = cantidad ' Asigna la cantidad seleccionada del producto
    xTransferencia(ultimo, 5) = Format(Precio, "###,##0.00") ' Asigna el tipo de Venta
    xTransferencia(ultimo, 6) = TipoVenta ' Asigna el tipo de Venta
    xTransferencia(ultimo, 7) = FlagFraccion ' Asigna si la cantidad es fraccion o Unidad
    xTransferencia(ultimo, 8) = CodigoLocalOrigen  ' Asigna el codigo de Local
    xTransferencia(ultimo, 9) = CodigoLocalDestino ' Asigna el codigo de Local
    xTransferencia(ultimo, 10) = CodigoLocalSapOrigen ' Asigna el codigo de Local SAP
    xTransferencia(ultimo, 11) = CodigoLocalSapDestino ' Asigna el codigo de Local SAP
    Set AgregaTransferencia = xTransferencia
    
End Function


'*************************************************************************************
' Metodo para seleccionar un producto para la venta, retorna como variable el propio array
'*************************************************************************************
Public Function AgregaServicio(CodigoPadre As String, _
                            DescripcionPadre As String, _
                            CodigoHijo As String, _
                            DescripcionHijo As String, _
                            NumOperacion As String, _
                            TipoVenta As TipoVenta, _
                            Monto As String, _
                            TipoCambio As String, _
                            Importe As String, _
                            Producto As String, _
                            NumeroRecibo As String, _
                            NumeroSuministro As String, _
                            FlagTipoPago As String, Optional ByVal CantidadServicio As Integer = 1 _
                            ) As XArrayDB
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    'Dim strCajaAbierta As String
    
    
    If xServicio.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    ''Dim PctComi As Double
    ''Dim objproducto As New clsProducto
    
    '** BITACORA **'
    
    'Autor: Juan Arturo Escate Espichan
    'Fecha: 14/04/2008
    'Proposito: esto es para que no se confunda navsat , para que sea solo un documento

    If EsNavsat(CodigoPadre) = "1" And xServicio.Count(1) <> 0 Then
            MsgBox "Solo se permite una transaccion de Navsat por documento", vbCritical, App.ProductName
            Exit Function
    End If
    
    'IF xProducto.Find(0,5, Servicio)<>
    '-----------------------------------------------------------------------------------------------------
    
    
    
    ''PctComi = objproducto.PctComision(Codigo, objUsuario.CodigoLocal, Format(ptmTipoPrecio, "000"))
    
    aux = xServicio.Count(1)
    While i < aux
        'If xProducto(i, 0) = Codigo And xProducto(i, 5) = TipoVenta Then
        If xServicio(i, 0) = CodigoPadre And xServicio(i, 2) = CodigoHijo And xServicio(i, 4) = NumOperacion Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xServicio.Count(1) 'Llena con la ultima posicion disponible
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xServicio.AppendRows
    End If
j:
    If xServicio.Count(1) = 0 Then ultimo = 0: xServicio.AppendRows
    xServicio(ultimo, 0) = CodigoPadre
    xServicio(ultimo, 1) = DescripcionPadre
    xServicio(ultimo, 2) = CodigoHijo
    xServicio(ultimo, 3) = DescripcionHijo
    xServicio(ultimo, 4) = NumOperacion
    xServicio(ultimo, 5) = TipoVenta
    xServicio(ultimo, 6) = Monto
    xServicio(ultimo, 7) = TipoCambio
    xServicio(ultimo, 8) = Importe
    xServicio(ultimo, 9) = Producto
    xServicio(ultimo, 10) = NumeroRecibo
    xServicio(ultimo, 11) = NumeroSuministro
    xServicio(ultimo, 12) = FlagTipoPago
    xServicio(ultimo, 13) = CantidadServicio
    Set AgregaServicio = xServicio
    
End Function

Public Sub CargaConceptoRemesa()
    Dim odynRemesa  As oraDynaset
    Dim strConcepto$, strDesConcepto$, strCta1$, strCta2$, strCta3$, strTipoCta$, strFP_P$, strFP_H$
    
    On Error GoTo CtrlErr
    
    'Llamada a la Consulta'
    Set odynRemesa = objLiquidacion.ListaConcepto
    
    odynRemesa.MoveFirst
    While Not odynRemesa.EOF
            
        strConcepto = odynRemesa("COD_CONCEPTO").Value
        strDesConcepto = odynRemesa("DES_CONCEPTO").Value
        strCta1 = odynRemesa("NUM_CTA_CTE_01").Value
        strCta2 = odynRemesa("NUM_CTA_CTE_02").Value
        strCta3 = odynRemesa("NUM_CTA_CTE_03").Value
        strTipoCta = odynRemesa("FLG_TIPO_CTA").Value
        strFP_P = odynRemesa("COD_FORMA_PAGO").Value
        strFP_H = odynRemesa("COD_HIJO").Value
            
        objLiquidacion.AgregaConceptoRemesa strConcepto, strDesConcepto, strCta1, strCta2, strCta3, strTipoCta, strFP_P, strFP_H
        odynRemesa.MoveNext
    Wend
    Exit Sub
CtrlErr:
    Err.Raise Err.Number, "clsLiquidacion", Err.Description
End Sub

'-- Recetario Magistral --'
Public Function AgregaRecetarioM(ByVal codProd As String, _
                                 ByVal CodIns As String, _
                                 ByVal cantidad As String, _
                                 ByVal FlagFraccion As String, _
                                 ByVal Precio As String, _
                                 ByVal TipoVta As String, _
                                 ByVal PctBase As String, _
                                 ByVal CantBase As String, _
                                 ByVal UndMedida As String) As XArrayDB
                            
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If xRecetarioMagistral.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    aux = xRecetarioMagistral.Count(1)
    While i < aux
        'If xProducto(i, 0) = Codigo And xProducto(i, 5) = TipoVenta Then
        If xRecetarioMagistral(i, 1) = CodIns Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xRecetarioMagistral.Count(1) 'Llena con la ultima posicion disponible
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xRecetarioMagistral.AppendRows
    End If
j:
    If xRecetarioMagistral.Count(1) = 0 Then ultimo = 0: xRecetarioMagistral.AppendRows
    xRecetarioMagistral(ultimo, 0) = codProd ' Asigna el codigo de producto
    xRecetarioMagistral(ultimo, 1) = CodIns ' Asigna la descripcion del producto
    xRecetarioMagistral(ultimo, 2) = FlagFraccion ' Asigna si la cantidad es fraccion o Unidad
    xRecetarioMagistral(ultimo, 3) = cantidad ' Asigna la cantidad seleccionada del producto
    xRecetarioMagistral(ultimo, 9) = Format(Precio, "###,##0.00")
    xRecetarioMagistral(ultimo, 5) = TipoVta
    xRecetarioMagistral(ultimo, 6) = IIf(Trim(PctBase = ""), "0", PctBase)
    xRecetarioMagistral(ultimo, 7) = IIf(Trim(CantBase = ""), "0", CantBase)
    xRecetarioMagistral(ultimo, 8) = UndMedida
    Set AgregaRecetarioM = xRecetarioMagistral
    
End Function

'*************************************************************************************
' Metodo para correción de precios
'*************************************************************************************
Public Function CorreccionPrecio(Codigo As String, _
                            Descripcion As String, _
                            cantidad As Integer, _
                            FlagFraccion As String, _
                            Precio As Double, _
                            TipoVenta As TipoVenta, _
                            Regalo As String, _
                            TipoDcto As eTipoValor, _
                            ImpDcto As Double, _
                            PrcAnt As Double, _
                            CodAuto As String, _
                            CodUsu As String) As XArrayDB
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If xProducto.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    Dim PctComi As Double
    Dim objProducto As New clsProducto
    
    PctComi = objProducto.pctComision(Codigo, objUsuario.CodigoLocal, Format(ptmTipoPrecio, "000"))
        
    
    
    aux = xProducto.Count(1)
    While i < aux
        'If xProducto(i, 0) = Codigo And xProducto(i, 5) = TipoVenta Then
        If xProducto(i, 0) = Codigo Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xProducto.Count(1) 'Llena con la ultima posicion disponible
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xProducto.AppendRows
    End If
j:
    If xProducto.Count(1) = 0 Then ultimo = 0: xProducto.AppendRows
    xProducto(ultimo, 0) = Codigo ' Asigna el codigo de producto
    xProducto(ultimo, 1) = Descripcion ' Asigna la descripcion del producto
    xProducto(ultimo, 2) = IIf(FlagFraccion = "1", "F", "U") ' Asigna si la cantidad es fraccion o Unidad
    xProducto(ultimo, 3) = val(cantidad) ' Asigna la cantidad seleccionada del producto
    xProducto(ultimo, 4) = Format(Precio, "###,##0.00") ' Asigna el tipo de Venta
    xProducto(ultimo, 5) = TipoVenta ' Asigna el tipo de Venta
    xProducto(ultimo, 6) = FlagFraccion ' Asigna si la cantidad es fraccion o Unidad
    xProducto(ultimo, 7) = Regalo 'valor libre
    
    xProducto(ultimo, 8) = TipoDcto
    xProducto(ultimo, 9) = ImpDcto
    xProducto(ultimo, 10) = PrcAnt
    xProducto(ultimo, 11) = CodAuto
    xProducto(ultimo, 12) = CodUsu
    
    xProducto(ultimo, 13) = PctComi
    Set CorreccionPrecio = xProducto
    'Set xProductoData = xProducto
End Function




Public Function EliminaProducto(Codigo As String, TipoVenta As TipoVenta) As XArrayDB
    'MsgBox strCodigoTipoVenta
    Dim objProducto As New clsProducto
    Dim ultimo As Integer ' declara variable contador
    If xProducto.Count(1) <= 0 Then Exit Function
    Dim aux As Integer
    Dim i As Integer
    aux = xProducto.Count(1)
    i = 0
    While i < xProducto.Count(1)
        If xProducto(i, 0) = Codigo And xProducto(i, 5) = TipoVenta Then
            ultimo = i
        End If
        i = i + 1
    Wend
    ''esto es cuando es un regalo elimine la promociones de los productos
    If xProducto(ultimo, 7) <> "0" And xProducto(ultimo, 5) = TipoVenta Then
        ''primero busco si existe otro regalo con la misma promocion
        Dim Y As Integer
        Dim existe As Boolean
        existe = False
            While Y < aux
                If val(xProducto(ultimo, 3)) > 0 And xProducto(Y, 5) = TipoVenta And xProducto(Y, 0) <> Codigo And xProducto(Y, 7) <> "0" And xProducto(Y, 26) = xProducto(ultimo, 26) Then  '' si no es el mismo producto y si es un regalo  y es la misma venta y es el mismo codigo de promocion
                    existe = True
                End If
                Y = Y + 1
            Wend
            
            If existe = False Then
                Y = 0
                While Y < aux
                                    
                    xProducto(Y, 26) = Replace(xProducto(Y, 26), Replace(xProducto(ultimo, 26), ";", ""), "")
                    'Al ejecutar la sentencia de arriba, el resultado final es (;) por lo que se agrego la
                    'siguiente linea a fin de eliminar ese caracter, ya que ese valor es utiliza posteriormente
                    'en una validacion de recarga de listado de pedidos.
                    If xProducto(Y, 26) = ";" Then
                        xProducto(Y, 26) = ""
                    End If
                    Y = Y + 1
                Wend
            End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''
     'xProducto.DeleteRows (ultimo)
    
    'Actualizacion mlevano 26/09/2012
    'Valido para la opcion de Guia de Remision
    If strCodigoTipoVenta = 10 Then
        ultimo = 0
        i = 0
        'n = 0
        Dim n As Integer
        While i < xProductoLote.Count(1)
            If xProductoLote(i, 0) = Codigo Then
                xProductoLote(i, 0) = ""
                n = n + 1
            End If
            i = i + 1
        Wend
    End If
    'Termina
    
    Set EliminaProducto = xProducto
End Function

'Modificado el 26/09/2012 mlevano
Public Function EliminaProductoNroLote(Codigo As String, TipoVenta As TipoVenta, NroLote As String, FracEnabled As String) As XArrayDB
    
    Dim ultimo As Integer ' declara variable contador
    Dim ultimo2 As Integer ' declara variable contador
    If xProducto.Count(1) <= 0 Then Exit Function
    Dim i As Integer
    Dim objProducto As New clsProducto
    Dim canti As Integer
    Dim prec As Double
    Dim aux As Integer
    Dim Fraccionable As String
    Dim cantidadFraccionable As Integer
    
    'Obteniendo la posicion en grdpedido
    aux = xProducto.Count(1)
    i = 0
    While i < aux
        If xProducto(i, 0) = Codigo And xProducto(i, 5) = TipoVenta Then
            ultimo = i
        End If
        i = i + 1
    Wend
    
    'Obteniendo la cantidad fraccionable del producto
    cantidadFraccionable = objProducto.DevFraccionProducto(Codigo)
    
    'Obteniendo la cantidad total en fracciones y sus precios
    i = 0
    While i < xProductoLote.Count(1)
        If xProductoLote(i, 0) = Codigo Then
            If xProductoLote(i, 2) = "F" Then
                canti = canti + xProductoLote(i, 3)
                prec = prec + xProductoLote(i, 4)
            Else
                canti = canti + (xProductoLote(i, 3) * cantidadFraccionable)
                prec = prec + xProductoLote(i, 4)
            End If
        End If
        i = i + 1
    Wend
    
    'Obteniendo resta: total - producto a borrar
    i = 0
    While i < xProductoLote.Count(1)
        If xProductoLote(i, 0) = Codigo And xProductoLote(i, 22) = NroLote Then
        ultimo2 = i
            If xProductoLote(ultimo2, 2) = "U" Then
                canti = canti - (xProductoLote(ultimo2, 3) * cantidadFraccionable)
            Else
                canti = canti - xProductoLote(ultimo2, 3)
            End If
        End If
        i = i + 1
    Wend
    
    'Obteniendo si es F o U, y si el numero se convierte en U o se mantiene en F
    If canti Mod cantidadFraccionable = 0 Then
        canti = canti / cantidadFraccionable
        Fraccionable = "U"
    Else
        canti = canti
        If FracEnabled = "1" Then
            Fraccionable = "F"
        Else
            Fraccionable = "U"
        End If
    End If
        
    'Devolviendo valores para la emision de la guia
    xProducto(ultimo, 3) = canti
    xProducto(ultimo, 4) = xProducto(ultimo, 4) - xProductoLote(ultimo2, 4)
    xProducto(ultimo, 2) = Fraccionable
    If xProducto(ultimo, 3) <= 0 Then
        xProducto.DeleteRows (ultimo)
    End If
        
    xProductoLote.DeleteRows (ultimo2)
    'xProductoLote(ultimo2, 0) = ""
    Set EliminaProductoNroLote = xProducto
    'xProductoLote.DeleteRows (ultimo)
    'Set EliminaProductoNroLote = xProductoLote
End Function

Public Property Get Totales() As Variant
    Dim i As Integer
    Dim dImporte As Variant
    Dim rs As OracleInProcServer.oraDynaset
    Dim objFormaPago As New clsFormaPago
    Dim dblRedondeo As Double
    Dim dblDonacion As Double
    Dim dblCoPago As Double
    Dim dlbg As Double
    Dim g As Integer
    Dim bolEfectivo As Boolean
    
    Dim dImporteAux As Double
    
    
    dImporte = Array("0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0")
    dblRedondeo = 0
    dblDonacion = 0
    dblCoPago = 0
    dlbg = 0
    For g = 0 To objVenta.Servicios.UpperBound(1)
        dlbg = dlbg + val(objVenta.Servicios(g, 8))
    Next g
    
    If dlbg > 0 Then dImporte(0) = dImporte(0) + CDbl(dlbg)
        
        While i <= xProducto.UpperBound(1)
        'While i <= xProducto.UpperBound(1)
            
         dImporteAux = dImporteAux + CDbl(xProducto(i, 4))
        
        '*****************************************************************************************'
        '   Evalua el precio menor para poder utilizarlo y sacar el copago 30/12/2008
        '*****************************************************************************************'
        
        
        '** FUE COMENTADO PORQUE PABLO DIJO Y COORDINO CON PSALAZAR  08/01/2009
        
''''''        If objVenta.CodModalidadVenta = "002" Then
''''''
''''''                If objVenta.PrecioMenorDeducible = "1" Then
''''''
''''''                    If CDbl(xProducto(i, 4)) > CDbl(xProducto(i, 25)) Then
''''''
''''''                         dImporte(0) = dImporte(0) + CDbl(xProducto(i, 25))
''''''
''''''                      ElseIf CDbl(xProducto(i, 4)) <= CDbl(xProducto(i, 25)) Then
''''''
''''''                         dImporte(0) = dImporte(0) + CDbl(xProducto(i, 4))
''''''
''''''                    End If
''''''
''''''                  ElseIf objVenta.PrecioMenorDeducible = "0" Then
''''''
''''''                    dImporte(0) = dImporte(0) + CDbl(xProducto(i, 4))
''''''
''''''                End If
''''''
''''''           Else
''''''
''''''              dImporte(0) = dImporte(0) + CDbl(xProducto(i, 4))
''''''
''''''         End If
         
         dImporte(0) = dImporte(0) + CDbl(xProducto(i, 4))
         
         dImporte(12) = dImporte(12) + (CDbl(xProducto(i, 25)) - CDbl(xProducto(i, 4)))
          dImporte(13) = dImporte(13) + CDbl(xProducto(i, 28))
         If dImporte(12) < 0 Then
            dImporte(12) = 0
         End If
        '*****************************************************************************************'
        '*****************************************************************************************'
        
        If xProducto(i, 6) = "1" Then
            dImporte(3) = dImporte(3) + 1
        Else
            dImporte(4) = dImporte(4) + 1
        End If
        
        'flgIntPctBeneficiario
        'me.ImpCopagoxRango(me.CodigoConvenio,
        If PctBeneficiario > 0 Then
            If flgIntPctBeneficiario = "0" Then
            
                '*****************************************************************************************'
                '** Evalua el precio menor para poder utilizarlo y sacar el copago 30/12/2008 **'
                '*****************************************************************************************'
                'If PctBeneficiarioOriginal = "" Then PctBeneficiarioOriginal = PctBeneficiario
                
                If objVenta.CodModalidadVenta = "002" Then
                
                      If objVenta.PrecioMenorDeducible = "1" Then
                            
                          dblCoPago = dblCoPago + Round(CDbl(xProducto(i, 25)) * PctBeneficiario / 100, 2)
                          If dImporte(8) <> 0 Then
                            Me.PctBeneficiarioReal = dImporte(0) / dImporte(8) * 100 'Arturo escate 02/02/2010
                          End If
                       ElseIf (objVenta.PrecioMenorDeducible = "0" Or objVenta.PrecioMenorDeducible = "") Then
                           
                          dblCoPago = dblCoPago + Round(CDbl(xProducto(i, 4)) * PctBeneficiario / 100, 2)
                           Me.PctBeneficiarioReal = PctBeneficiario 'Arturo escate 02/02/2010
                      End If
                  Else
                      dblCoPago = dblCoPago + Round(CDbl(xProducto(i, 4)) * PctBeneficiario / 100, 2)
                End If
                '*****************************************************************************************'
                '*****************************************************************************************'
               'Me.PctBeneficiario = Round((dblCoPago / dImporte(0)), 4) * 100
            Else
                
                dblCoPago = Me.ImpPctBeneficiario
                
                Me.PctBeneficiario = Round((Me.ImpPctBeneficiario / dImporte(0)), 4) * 100
                Dim ValorNew As Double
                 ValorNew = Me.ImpCopagoxRango(Me.CodigoConvenio, CDbl(dImporte(0)))
                 If ValorNew = -1 Then
                    dblCoPago = Me.ImpPctBeneficiario
                 Else
                    dblCoPago = ValorNew
                 End If
                    
                 
                Me.PctBeneficiario = Round((dblCoPago / dImporte(0)), 4) * 100
                Me.PctBeneficiarioReal = PctBeneficiario 'Arturo escate 02/02/2010
            End If
        
        End If
        
        i = i + 1
    Wend
        
        If Not dImporte(0) = 0 Then
            Me.PctBeneficiarioReal = dblCoPago / dImporte(0) * 100 'Arturo escate 02/02/2010
        End If

        'dImporte(2) = Format(Round(0, 1), "###,##0.00") ' Redondeo
        dImporte(2) = Round(0, 1) ' Redondeo
        dImporte(5) = i
        
        'dImporte(8) = Round(dImporte(0) * (PctBeneficiario / 100), 2)
        dImporte(8) = dblCoPago
        
        NewPctBeneficario = dImporte(8)
        
    'Forma pago de pago
    i = 0
    bolEfectivo = False
    While i <= xFormaPago.UpperBound(1)
        
        'If xFormaPago(i, 4) > 0 Then dImporte(6) = Format(dImporte(6) + xFormaPago(i, 4), "###,##0.00")
        
        '*****************************************************************************************'
        '** Evalua el precio menor para poder utilizarlo y sacar el copago 30/12/2008 **'
        '*****************************************************************************************'
        
        'If objVenta.CodModalidadVenta = "002" Then
        
         '       If objVenta.PrecioMenorDeducible = "1" Then
                   'If CDbl(xProducto(i, 4)) > CDbl(xProducto(i, 25)) Then
                   
         '           If xFormaPago(i, 25) > 0 Then dImporte(6) = dImporte(6) + CDbl(xFormaPago(i, 4))
                       
                    'ElseIf CDbl(xProducto(i, 4)) <= CDbl(xProducto(i, 25)) Then
                       
                    '   If xFormaPago(i, 4) > 0 Then dImporte(6) = dImporte(6) + CDbl(xFormaPago(i, 4))
                       
                   'End If
                   
          '      ElseIf objVenta.PrecioMenorDeducible = "0" Then
                           
          '           If xFormaPago(i, 4) > 0 Then dImporte(6) = dImporte(6) + CDbl(xFormaPago(i, 4))
                   
           '     End If
                
            'Else
            
           '   If xFormaPago(i, 4) > 0 Then dImporte(6) = dImporte(6) + CDbl(xFormaPago(i, 4))
                
        'End If
        
        If xFormaPago(i, 0) = "001" And bolEfectivo = False Then
            bolEfectivo = True
        End If
        
         If xFormaPago(i, 4) > 0 And xFormaPago(i, 0) <> "011" Then dImporte(6) = dImporte(6) + CDbl(xFormaPago(i, 4)) '** 05/01/2009
           
        '*****************************************************************************************'
        '*****************************************************************************************'
        
        i = i + 1
    Wend
    
    If dImporte(6) > 0 Then
        '###esta llamada se hace 5 veces por llamado a sub_totales
        Set rs = objFormaPago.ListaDonacion
        While Not rs.EOF
            i = 0
            While i <= xFormaPago.UpperBound(1)
                If rs("COD_FORMA_PAGO").Value = xFormaPago(i, 0) And _
                    rs("COD_HIJO").Value = xFormaPago(i, 2) Then
                    
                    '*****************************************************************************************'
                    '** Evalua el precio menor para poder utilizarlo y sacar el copago 30/12/2008 **'
                    '*****************************************************************************************'
                    
                    'If objVenta.CodModalidadVenta = "002" Then
                    
                     '       If objVenta.PrecioMenorDeducible = "1" Then
                     '               If CDbl(xProducto(i, 4)) > CDbl(xProducto(i, 25)) Then
                                    
                     '                   dImporte(9) = dImporte(9) + CDbl(xFormaPago(i, 4))
                                        
                     '                ElseIf CDbl(xProducto(i, 4)) <= CDbl(xProducto(i, 25)) Then
                                        
                     '                   dImporte(9) = dImporte(9) + CDbl(xFormaPago(i, 4))
                                        
                     '               End If
                               
                     '       ElseIf objVenta.PrecioMenorDeducible = "0" Then
                            
                     '            dImporte(9) = dImporte(9) + CDbl(xFormaPago(i, 4))
                              
                      '      End If
                            
                      ' Else
                       
                      '      dImporte(9) = dImporte(9) + CDbl(xFormaPago(i, 4))
                            
                     'End If
                     
                      dImporte(9) = dImporte(9) + CDbl(xFormaPago(i, 4))  '** 05/01/2009
                     
                    '*****************************************************************************************'
                    
                End If
                i = i + 1
            Wend
            rs.MoveNext
        Wend
    End If
    
    
    
    
    If dImporte(6) > 0 Then dImporte(7) = Round(dImporte(6) - dImporte(0) + dImporte(9), 2)
    
    If dImporte(8) > 0 Then dImporte(7) = Round(dImporte(6) - dImporte(8) + dImporte(9), 2)
    
    If dImporte(7) > 0 Then
        dblRedondeo = val("0.0" + right(Format(dImporte(7), "##0.00"), 1))
    Else
        dblRedondeo = -1 * val("0.0" + right(Format(dImporte(7), "##0.00"), 1))
    End If
    
    If dblRedondeo > 0.05 Then
        dblRedondeo = 0.1 - dblRedondeo
    Else
        'If dblRedondeo < 0 Then dblRedondeo = -1 * dblRedondeo
        dblRedondeo = -1 * dblRedondeo
    End If

    
'    Else
'
'    End If
    'dblRedondeo = 0
    
    If bolEfectivo Then
        
        Debug.Print val(dImporte(0))
        dImporte(1) = fnRedondeo(val(dImporte(0))) 'dblRedondeo
    Else
        dImporte(1) = 0
    End If
    
    
    If dImporte(6) > 0 Then
        '###esta llamada se hace 5 veces por llamado a sub_totales
        Set rs = objFormaPago.ListaRedondeo
        If Not rs.EOF Then AgregaFormaPago rs("COD_FORMA_PAGO").Value, _
                                            rs("DES_FORMA_PAGO").Value, _
                                            rs("COD_HIJO").Value, _
                                            rs("DES_HIJO").Value, _
                                            -1 * Abs(val(dImporte(1))), , , , , , , , , , , , , , , , , , , , , , , , , , , _
                                            rs("FLG_REDONDEO").Value

    End If
    
    Set objFormaPago = Nothing
    Set rs = Nothing
    
    'dImporte(2) = dImporte(7) - dblRedondeo  borrado por jmelgar
    'dImporte(2) = dImporte(7) + dImporte(1)   'agregado por jmelgar
    
    dImporte(2) = dImporte(7) - dblRedondeo
    strMontoTotalTotal = dImporte(0)
    
    'If objVenta.PrecioMenorDeducible = "1" Then
        
        'dImporte(0) = Val(dImporteAux) - Val(dblCoPago)
        
    'End If
    
    dImporte(10) = objUsuario.TipoCambio
    
    dImporte(11) = Round(dImporte(0) / IIf(objUsuario.TipoCambio = 0, 1, objUsuario.TipoCambio), 2)
    
    
    
    Totales = dImporte 'Devuelve la variable de la clase en forma de XArray
End Property

'*************************************************************************************
' Asigna los parametros de un convenio
'*************************************************************************************
Public Function AgregaConvenio(Codigo As String, Nombre As String) As XArrayDB
    strNombreConvenio = Nombre
    strCodigoConvenio = Codigo
    xConvenio(0, 0) = Codigo
    xConvenio(0, 1) = Nombre
End Function


'*************************************************************************************
' Propiedad que devuelve las caracteristicas del convenio
'*************************************************************************************
Public Property Get Convenio() As XArrayDB
    Set Convenio = xProducto 'Devuelve la variable de la clase en forma de XArray
End Property


'*************************************************************************************
' Propiedad que devuelve las Nombre del convenio
'*************************************************************************************
Public Property Get NombreConvenio() As String
    NombreConvenio = strNombreConvenio 'Devuelve el nombre del convenio
End Property


'*************************************************************************************
' Propiedad que asigna las Nombre del convenio
'*************************************************************************************
Public Property Let NombreConvenio(ByVal lNombreConvenio As String)
    strNombreConvenio = lNombreConvenio
End Property


'*************************************************************************************
' Propiedad que devuelve las Codigo del convenio
'*************************************************************************************
Public Property Get CodigoConvenio() As String
    CodigoConvenio = strCodigoConvenio 'Devuelve el nombre del convenio
End Property


'*************************************************************************************
' Propiedad que asigna las Codigo Convenio
'*************************************************************************************
Public Property Let CodigoConvenio(ByVal lCodigoConvenio As String)
    strCodigoConvenio = lCodigoConvenio
End Property

'*************************************************************************************
' Propiedad que devuelve las Codigo del Beneficiario
'*************************************************************************************
Public Property Get CodigoBeneficiario() As String
    CodigoBeneficiario = strCodigoBeneficiario 'Devuelve el nombre del convenio
End Property

'*************************************************************************************
' Propiedad que asigna las Codigo Beneficiario
'*************************************************************************************
Public Property Let CodigoBeneficiario(ByVal lCodigoBeneficiario As String)
    strCodigoBeneficiario = lCodigoBeneficiario
End Property

'*************************************************************************************
' Propiedad que devuelve el numero de receta
'*************************************************************************************
Public Property Get DetalleReceta() As String
    DetalleReceta = strNumeroReceta
End Property

Public Property Let DetalleReceta(ByVal lCodigoBeneficiario As String)
    strNumeroReceta = lCodigoBeneficiario
End Property

'*************************************************************************************
' Propiedad que devuelve el fecha de la receta
'*************************************************************************************
Public Property Get FechaReceta() As String
    FechaReceta = strFechaReceta
End Property
Public Property Let FechaReceta(ByVal lCodigoBeneficiario As String)
    strFechaReceta = lCodigoBeneficiario
End Property

'*************************************************************************************
' Propiedad que devuelve el Local de la Receta
'*************************************************************************************
Public Property Get LocalReceta() As String
    LocalReceta = strLocalReceta
End Property

Public Property Let LocalReceta(ByVal lCodigoBeneficiario As String)
    strLocalReceta = lCodigoBeneficiario
End Property





'*************************************************************************************
' Propiedad que devuelve las Codigo del Beneficiario
'*************************************************************************************
Public Property Get NombreBeneficiario() As String
    NombreBeneficiario = strNombreBeneficiario 'Devuelve el nombre del convenio
End Property

'*************************************************************************************
' Propiedad que asigna las Codigo Beneficiario
'*************************************************************************************
Public Property Let NombreBeneficiario(ByVal lNombreBeneficiario As String)
    strNombreBeneficiario = lNombreBeneficiario
End Property

'*************************************************************************************
' Propiedad que devuelve las Codigo del Tipo Venta
'*************************************************************************************
Public Property Get CodigoTipoVenta() As TipoVenta
    CodigoTipoVenta = strCodigoTipoVenta 'Devuelve el nombre del convenio
End Property

'*************************************************************************************
' Propiedad que asigna las Codigo Tipo Venta
'*************************************************************************************
Public Property Let CodigoTipoVenta(ByVal lCodigoTipoVenta As TipoVenta)
    strCodigoTipoVenta = lCodigoTipoVenta
End Property


Public Property Get CodModalidadVenta() As String
    CodModalidadVenta = strCodModalidadVenta 'Devuelve el nombre del convenio
End Property

'*************************************************************************************
' Propiedad que asigna las Codigo Tipo Venta
'*************************************************************************************
Public Property Let CodModalidadVenta(ByVal lCodModalidadVenta As String)
    strCodModalidadVenta = lCodModalidadVenta
End Property


'*************************************************************************************
' Propiedad que devuelve las Codigo del Beneficiario
'*************************************************************************************
Public Property Get NombreTipoVenta() As String
    Select Case strCodigoTipoVenta
    Case 1
        NombreTipoVenta = "Venta Regular"
    Case 2
        NombreTipoVenta = "Venta Convenio"
    Case 3
        NombreTipoVenta = "Venta Mayorista"
    Case 4
        NombreTipoVenta = "Cobro por Responsabilidad"
    Case 5
        NombreTipoVenta = "Canje Promociones"
    Case 6
        NombreTipoVenta = "Servicios"
    Case 7
        NombreTipoVenta = "Cobranza Credito"
    Case 8
        NombreTipoVenta = "Cotizaciones"
    Case 9
        NombreTipoVenta = "Facturación"
    Case 10
        NombreTipoVenta = "Guia de Remisión"
    Case 17
        NombreTipoVenta = "Recetario Magistral"
    Case 12
        NombreTipoVenta = "Pedido DLV"
    Case 19
        NombreTipoVenta = "Cajero Corresponsal"
        
    End Select
End Property

Sub OrdenarArregloFormaPago()
    Dim aux As Integer
    Dim pos As Integer
    If xFormaPago.Count(1) <= 0 Then Exit Sub
    
    Dim xFormaPagoT As New XArrayDB
    xFormaPagoT.ReDim 0, -1, 0, 35 '33 ECASTILLO 13.03.2020
    
    Dim i As Integer
    aux = xFormaPago.Count(1)
    i = 0
    While i < aux
        If xFormaPago(i, 0) = "008" Then
            pos = i
            xFormaPagoT.AppendRows
            xFormaPagoT(0, 0) = xFormaPago(i, 0)
            xFormaPagoT(0, 1) = xFormaPago(i, 1)
            xFormaPagoT(0, 2) = xFormaPago(i, 2)
            xFormaPagoT(0, 3) = xFormaPago(i, 3)
            xFormaPagoT(0, 4) = xFormaPago(i, 4)
            xFormaPagoT(0, 5) = xFormaPago(i, 5)
            xFormaPagoT(0, 6) = xFormaPago(i, 6)
            xFormaPagoT(0, 7) = xFormaPago(i, 7)
            xFormaPagoT(0, 9) = xFormaPago(i, 9)
            xFormaPagoT(0, 10) = xFormaPago(i, 10)
            xFormaPagoT(0, 11) = xFormaPago(i, 11)
            xFormaPagoT(0, 12) = xFormaPago(i, 12)
            xFormaPagoT(0, 13) = xFormaPago(i, 13)
            xFormaPagoT(0, 14) = xFormaPago(i, 14)
            xFormaPagoT(0, 15) = xFormaPago(i, 15)
            xFormaPagoT(0, 16) = xFormaPago(i, 16)
            xFormaPagoT(0, 17) = xFormaPago(i, 17)
            xFormaPagoT(0, 19) = xFormaPago(i, 19)
            xFormaPagoT(0, 21) = xFormaPago(i, 21)
            xFormaPagoT(0, 22) = xFormaPago(i, 22)
            xFormaPagoT(0, 25) = xFormaPago(i, 25)
            xFormaPagoT(0, 29) = xFormaPago(i, 29)
            xFormaPagoT(0, 31) = xFormaPago(i, 31) 'ECASTILLO 29.04.2020
            xFormaPagoT(0, 32) = xFormaPago(i, 32)
            xFormaPagoT(0, 33) = xFormaPago(i, 33)
            xFormaPagoT(0, 34) = xFormaPago(i, 34) 'ECASTILLO 13.03.2020
            xFormaPagoT(0, 35) = xFormaPago(i, 35) 'ECASTILLO 07.10.2020
            GoTo aqui
        End If
        i = i + 1
    Wend
aqui:
    For i = pos To aux - 1
        If i >= aux - 1 Then GoTo ACA:
        xFormaPago(i, 0) = xFormaPago(i + 1, 0)
        xFormaPago(i, 1) = xFormaPago(i + 1, 1)
        xFormaPago(i, 2) = xFormaPago(i + 1, 2)
        xFormaPago(i, 3) = xFormaPago(i + 1, 3)
        xFormaPago(i, 4) = xFormaPago(i + 1, 4)
        xFormaPago(i, 5) = xFormaPago(i + 1, 5)
        xFormaPago(i, 6) = xFormaPago(i + 1, 6)
        xFormaPago(i, 7) = xFormaPago(i + 1, 7)
        xFormaPago(i, 9) = xFormaPago(i + 1, 9)
        xFormaPago(i, 10) = xFormaPago(i + 1, 10)
        xFormaPago(i, 11) = xFormaPago(i + 1, 11)
        xFormaPago(i, 12) = xFormaPago(i + 1, 12)
        xFormaPago(i, 13) = xFormaPago(i + 1, 13)
        xFormaPago(i, 14) = xFormaPago(i + 1, 14)
        xFormaPago(i, 15) = xFormaPago(i + 1, 15)
        xFormaPago(i, 16) = xFormaPago(i + 1, 16)
        xFormaPago(i, 17) = xFormaPago(i + 1, 17)
        xFormaPago(i, 19) = xFormaPago(i + 1, 19)
        xFormaPago(i, 21) = xFormaPago(i + 1, 21)
        xFormaPago(i, 22) = xFormaPago(i + 1, 22)
        xFormaPago(i, 25) = xFormaPago(i + 1, 25)
        xFormaPago(i, 29) = xFormaPago(i + 1, 29)
        xFormaPago(i, 31) = xFormaPago(i + 1, 31) 'ECASTILLO 29.04.2020
        xFormaPago(i, 32) = xFormaPago(i + 1, 32)
        xFormaPago(i, 33) = xFormaPago(i + 1, 33)
        xFormaPago(i, 34) = xFormaPago(i + 1, 34) 'ECASTILLO 13.03.2020
        xFormaPago(i, 35) = xFormaPago(i + 1, 35) 'ECASTILLO 07.10.2020
    Next
ACA:
    If xFormaPagoT.Count(1) > 0 Then
        xFormaPago(aux - 1, 0) = xFormaPagoT(0, 0)
        xFormaPago(aux - 1, 1) = xFormaPagoT(0, 1)
        xFormaPago(aux - 1, 2) = xFormaPagoT(0, 2)
        xFormaPago(aux - 1, 3) = xFormaPagoT(0, 3)
        xFormaPago(aux - 1, 4) = xFormaPagoT(0, 4)
        xFormaPago(aux - 1, 5) = xFormaPagoT(0, 5)
        xFormaPago(aux - 1, 6) = xFormaPagoT(0, 6)
        xFormaPago(aux - 1, 7) = xFormaPagoT(0, 7)
        xFormaPago(aux - 1, 9) = xFormaPagoT(0, 9)
        xFormaPago(aux - 1, 10) = xFormaPagoT(0, 10)
        xFormaPago(aux - 1, 11) = xFormaPagoT(0, 11)
        xFormaPago(aux - 1, 12) = xFormaPagoT(0, 12)
        xFormaPago(aux - 1, 13) = xFormaPagoT(0, 13)
        xFormaPago(aux - 1, 14) = xFormaPagoT(0, 14)
        xFormaPago(aux - 1, 15) = xFormaPagoT(0, 15)
        xFormaPago(aux - 1, 16) = xFormaPagoT(0, 16)
        xFormaPago(aux - 1, 17) = xFormaPagoT(0, 17)
        xFormaPago(aux - 1, 19) = xFormaPagoT(0, 19)
        xFormaPago(aux - 1, 21) = xFormaPagoT(0, 21)
        xFormaPago(aux - 1, 22) = xFormaPagoT(0, 22)
        xFormaPago(aux - 1, 25) = xFormaPagoT(0, 25)
        xFormaPago(aux - 1, 29) = xFormaPagoT(0, 29)
        xFormaPago(aux - 1, 31) = xFormaPagoT(0, 31) 'ECASTILLO 29.04.2020
        xFormaPago(aux - 1, 32) = xFormaPagoT(0, 32)
        xFormaPago(aux - 1, 33) = xFormaPagoT(0, 33)
        xFormaPago(aux - 1, 34) = xFormaPagoT(0, 34) 'ECASTILLO 13.03.2020
        xFormaPago(aux - 1, 35) = xFormaPagoT(0, 35) 'ECASTILLO 07.10.2020
    End If
End Sub

Function AgregaFormaPago(ByVal CodigoFormaPago As String, _
                         ByVal DescripFormaPago As String, _
                         ByVal CodigoFormaPagoHijo As String, _
                         ByVal DescripFormaPagoHijo As String, _
                         ByVal ImporteCobrado As Double, _
                         Optional CodigoTarjeta As String = "", _
                         Optional CodigoMoneda As String = "", _
                         Optional CodigoDocumentoPago As String = "", _
                         Optional CodigoBanco As String = "", _
                         Optional CodigoDonacion As String = "", _
                         Optional CodigoCuentaCorriente As String = "", _
                         Optional ImporteTipoCambio As Double = 0, _
                         Optional NumeroTarjeta As String = "", _
                         Optional NumeroCuotas As String = "", _
                         Optional FechaVencimiento As String = "", _
                         Optional FlagCuotaNormal As String = "", _
                         Optional NumeroDocumentoDescuento As String = "", _
                         Optional NumeroMovimiento As String = "", _
                         Optional FlagVoucherManual As String = "", _
                         Optional FechaMovimiento As String = "", _
                         Optional NumeroAutorizacion As String = "", _
                         Optional FechaDocumento As String = "", _
                         Optional NumeroNotaCredito As String = "", _
                         Optional NumeroCheque As String = "", _
                         Optional NombCli As String = "", Optional DniCli As String = "", Optional CodProd_Vale As String = "", Optional CodBtl As String = "", Optional NumDocRef As String = "", Optional RetEfect As Double = 0, Optional ObsCheque As String = "", Optional flgRedondeo As String = "0", Optional ImporteDol As String = "", Optional NombTitular As String = "", Optional NumVale As String = "") As XArrayDB 'ECASTILLO 13.03.2020
                         
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If xFormaPago.Count(1) < 0 Then Exit Function
    'I.ECASTILLO 17.12.2020
    Dim flg_2e_reserva
    Dim noAgrega2Forma As Boolean
    Dim ii As Integer
    Dim yy As Integer
    flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
    If flg_2e_reserva = "0" Then
    Else
        If (xFormaPago.Count(1) >= 1 And (CodigoFormaPago <> "008" And CodigoFormaPagoHijo <> "027")) Then
            noAgrega2Forma = True
        End If
        If xFormaPago.Count(1) = 1 Then
            If (xFormaPago(0, 0) = "008" And xFormaPago(0, 2) = "027") Or (xFormaPago(0, 0) = "011" And xFormaPago(0, 2) = "001") Then noAgrega2Forma = False
        End If
        If (xFormaPago.Count(1) = 2) Then
            While ii < xFormaPago.Count(1)
                If xFormaPago(ii, 0) = "008" And xFormaPago(ii, 2) = "027" Then
                    noAgrega2Forma = False
                End If
                ii = ii + 1
            Wend
        End If
        '
        If xFormaPago.Count(1) > 0 Then
            yy = 0
            While yy < xFormaPago.Count(1)
                If (xFormaPago(yy, 0) = "002" Or _
                   xFormaPago(yy, 0) = "001") And _
                   (CodigoFormaPago <> "011" And CodigoFormaPago <> "008") Then
                   noAgrega2Forma = True
                   GoTo continuaFlujoFP
                Else
                   noAgrega2Forma = False
                End If
                yy = yy + 1
            Wend
        End If
    End If
    'F.ECASTILLO 17.12.2020
continuaFlujoFP:
    Dim i As Integer
    Dim encontro As Boolean
    aux = xFormaPago.Count(1)
    While i < aux
        If xFormaPago(i, 0) = CodigoFormaPago And xFormaPago(i, 2) = CodigoFormaPagoHijo And xFormaPago(i, 12) = NumeroTarjeta Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xFormaPago.Count(1) 'Llena con la ultima posicion disponible
        End If
        i = i + 1
    Wend
    If flg_2e_reserva = "0" Then
        If encontro = False Then
            xFormaPago.AppendRows
        End If
    Else
        If encontro = False And noAgrega2Forma = False Then
            xFormaPago.AppendRows
        End If
    End If
j:
    'I.ECASTILLO 17.12.2020
    If (noAgrega2Forma = True And encontro = False) Then Exit Function
    'F.ECASTILLO 17.12.2020
    '------'
    'Dim busca$
    'busca = xFormaPago.Find(0, 2, CodigoFormaPagoHijo)
    
    'If busca = -1 Then
    
        If xFormaPago.Count(1) = 0 Then ultimo = 0: xFormaPago.AppendRows
            xFormaPago(ultimo, 0) = CodigoFormaPago
            xFormaPago(ultimo, 1) = DescripFormaPago
            xFormaPago(ultimo, 2) = CodigoFormaPagoHijo
            xFormaPago(ultimo, 3) = DescripFormaPagoHijo
            xFormaPago(ultimo, 4) = ImporteCobrado
            xFormaPago(ultimo, 5) = CodigoTarjeta
            xFormaPago(ultimo, 6) = CodigoMoneda
            xFormaPago(ultimo, 7) = CodigoDocumentoPago
            xFormaPago(ultimo, 8) = CodigoBanco
            xFormaPago(ultimo, 9) = CodigoDonacion
            xFormaPago(ultimo, 10) = CodigoCuentaCorriente
            xFormaPago(ultimo, 11) = ImporteTipoCambio
            xFormaPago(ultimo, 12) = NumeroTarjeta
            xFormaPago(ultimo, 13) = NumeroCuotas
            xFormaPago(ultimo, 14) = FechaVencimiento
            xFormaPago(ultimo, 15) = FlagCuotaNormal
            xFormaPago(ultimo, 16) = NumeroDocumentoDescuento
            xFormaPago(ultimo, 17) = NumeroMovimiento
            xFormaPago(ultimo, 18) = FlagVoucherManual
            xFormaPago(ultimo, 19) = FechaMovimiento
            xFormaPago(ultimo, 20) = NumeroAutorizacion
            xFormaPago(ultimo, 21) = FechaDocumento
            xFormaPago(ultimo, 22) = NumeroNotaCredito
            xFormaPago(ultimo, 23) = NumeroCheque
            xFormaPago(ultimo, 24) = NombCli
            xFormaPago(ultimo, 25) = DniCli
            xFormaPago(ultimo, 26) = CodProd_Vale
            xFormaPago(ultimo, 27) = CodBtl
            xFormaPago(ultimo, 28) = NumDocRef
            xFormaPago(ultimo, 29) = RetEfect
            xFormaPago(ultimo, 30) = ObsCheque
            xFormaPago(ultimo, 31) = flgRedondeo
            xFormaPago(ultimo, 32) = ImporteDol
            xFormaPago(ultimo, 33) = NombTitular
            xFormaPago(ultimo, 34) = NumVale
            'xFormaPago(ultimo, 35) = "" 'ECASTILLO 07.10.2020
     '    Else
         
      '      xFormaPago(ultimo, 4) = ImporteCobrado
            
     'End If
     
'Set AgregaFormaPago = xProducto
Set AgregaFormaPago = xFormaPago
 End Function

'I.ECASTILLO 07.10.2020 - MODIFICA IMPORTE DE TARJETA, SOLO SI ES DISTINTO AL IMPORTE A PAGAR
'                       - PRUEBA
Public Function updTarjetaImporte(ByVal NuevoImporte As String, _
                                  Optional CodigoFormaPago As String, _
                                  Optional CodigoHijo As String)
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If xFormaPago.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    aux = xFormaPago.Count(1)
    While i < aux
        If xFormaPago(i, 35) = 0 _
           And xFormaPago(i, 0) = CodigoFormaPago _
           And xFormaPago(i, 4) <> NuevoImporte _
           Then
            ultimo = i
            encontro = True
        End If
        If encontro = True Then
            xFormaPago(ultimo, 4) = NuevoImporte
            xFormaPago(ultimo, 35) = 1
        End If
        i = i + 1
    Wend
End Function
'F.ECASTILLO 07.10.2020

Public Property Get TotalFormaPago() As Double
    Dim i As Integer
    Dim dPagado As Double
    dPagado = 0
    
    While i <= xFormaPago.UpperBound(1)
        dPagado = dPagado + val(xFormaPago(i, 4))
        i = i + 1
    Wend
    TotalFormaPago = dPagado 'Devuelve la variable de la clase en forma de XArray
End Property

Function AgregaMaquina(ByVal CodigoTipoDoc As String, _
                       ByVal CodigoMaquina As String, _
                       ByVal NumeroActual As String, _
                       ByVal FlagEncola As String, _
                       Optional MaquinaAlternativa As String = "", _
                       Optional Ticketera As String = "", _
                       Optional CodFormato As String = "") As XArrayDB
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If xMaquina.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    aux = xMaquina.Count(1)
    While i < aux
        If xMaquina(i, 0) = CodigoTipoDoc And xMaquina(i, 1) = CodigoMaquina Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xMaquina.Count(1) 'Llena con la ultima posicion disponible
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xMaquina.AppendRows
    End If
j:
        If xMaquina.Count(1) = 0 Then ultimo = 0: xMaquina.AppendRows
            xMaquina(ultimo, 0) = CodigoTipoDoc
            xMaquina(ultimo, 1) = CodigoMaquina
            xMaquina(ultimo, 2) = NumeroActual
            xMaquina(ultimo, 4) = MaquinaAlternativa
            If val(FlagEncola) = 1 Then
                xMaquina(ultimo, 3) = "Imprime en la maquina"
            Else
                xMaquina(ultimo, 3) = "Saca el correlativo de la Maquina"
            End If
            xMaquina(ultimo, 5) = FlagEncola
            xMaquina(ultimo, 6) = Ticketera
            xMaquina(ultimo, 7) = CodFormato
            
Set AgregaMaquina = xMaquina

End Function


Function AgregaCobroResponsabilidad(CodigoUsuario As String, _
                         NombreUsuario As String, _
                         ImporteCobrado As Double, _
                         FormaCalculo As Integer _
                        ) As XArrayDB
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If xCobroResponsabilidad.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    Dim ImpProrreteo As Double
    Dim ImpAcumulado As Double
    aux = xCobroResponsabilidad.Count(1)
    
    While i < aux
        If xCobroResponsabilidad(i, 0) = CodigoUsuario Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xCobroResponsabilidad.Count(1) 'Llena con la ultima posicion disponible
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xCobroResponsabilidad.AppendRows
    End If
    
    
j:
    If xCobroResponsabilidad.Count(1) = 0 Then ultimo = 0: xCobroResponsabilidad.AppendRows
        xCobroResponsabilidad(ultimo, 0) = CodigoUsuario
        xCobroResponsabilidad(ultimo, 1) = NombreUsuario
        xCobroResponsabilidad(ultimo, 2) = ImporteCobrado
        
    If FormaCalculo = 1 Then
        ImpProrreteo = Round(ImporteCobrado / (xCobroResponsabilidad.UpperBound(1) + 1), 2)
        For i = 0 To xCobroResponsabilidad.UpperBound(1)
           ImpAcumulado = ImpAcumulado + ImpProrreteo
           xCobroResponsabilidad(i, 2) = ImpProrreteo
        Next i
        
        If ImporteCobrado <> ImpAcumulado Then
            xCobroResponsabilidad(ultimo, 2) = ImpProrreteo + (ImporteCobrado - Round(ImpAcumulado, 2))
        End If
    End If
        
Set AgregaCobroResponsabilidad = xCobroResponsabilidad
End Function

Public Property Get TotalCobroResp() As Double
    Dim i As Integer
    Dim dImporte As Double
    dImporte = 0
    
    While i <= xCobroResponsabilidad.UpperBound(1)
        dImporte = dImporte + val(xCobroResponsabilidad(i, 2))
        i = i + 1
    Wend
    TotalCobroResp = dImporte 'Devuelve la variable de la clase en forma de XArray
End Property


Public Property Get TipoCliente() As String
    TipoCliente = strTipoCliente
End Property

Public Property Let TipoCliente(ByVal lValor As String)
    strTipoCliente = lValor
End Property

Public Property Get Ruc() As String
    Ruc = strRuc
End Property

Public Property Let Ruc(ByVal lValor As String)
    strRuc = lValor
End Property

Public Property Get RazonSocial() As String
    RazonSocial = strRazonSocial
End Property

Public Property Let RazonSocial(ByVal lValor As String)
    strRazonSocial = lValor
End Property

Public Property Get NumeroDocumentoID() As String
    NumeroDocumentoID = strNumeroDocumentoID
End Property

Public Property Let NumeroDocumentoID(ByVal lValor As String)
    strNumeroDocumentoID = lValor
End Property



Public Property Get NombreCliente() As String
    NombreCliente = strNombreCliente
End Property

Public Property Let NombreCliente(ByVal lValor As String)
    strNombreCliente = lValor
End Property

Public Property Get ApellidoCliente() As String
    ApellidoCliente = strApellidoCliente
End Property

Public Property Let ApellidoCliente(ByVal lValor As String)
    strApellidoCliente = lValor
End Property

Public Property Get DireccionCliente() As String
    DireccionCliente = strDireccionCliente
End Property

Public Property Let DireccionCliente(ByVal lValor As String)
    strDireccionCliente = lValor
End Property
Public Property Get DireccionClienteSocial() As String
    DireccionClienteSocial = strDireccionClienteSocial
End Property

Public Property Let DireccionClienteSocial(ByVal lValor As String)
    strDireccionClienteSocial = lValor
End Property

Public Property Get CodigoCliente() As String
    CodigoCliente = strCodigoCliente
End Property

Public Property Let CodigoCliente(ByVal lValor As String)
    strCodigoCliente = lValor
End Property

Public Property Get CodigoClienteDLV() As String
    CodigoClienteDLV = strCodigoClienteDLV
End Property

Public Property Let CodigoClienteDLV(ByVal lValor As String)
    strCodigoClienteDLV = lValor
End Property


Public Property Get CodigoDocumentoVenta() As String
    CodigoDocumentoVenta = strCodigoDocumentoVenta
End Property

Public Property Let CodigoDocumentoVenta(ByVal lValor As String)
    strCodigoDocumentoVenta = lValor
End Property

Public Property Let UbigeoEntrega(ByVal lValor As String)
    strUbigeoEntrega = lValor
End Property

Public Property Get UbigeoEntrega() As String
    UbigeoEntrega = strUbigeoEntrega
End Property
Public Property Let LocalAtencion(ByVal lValor As String)
    strLocalAtencion = lValor
End Property

Public Property Get LocalAtencion() As String
    LocalAtencion = strLocalAtencion
End Property

Public Property Let LocalDespacho(ByVal lValor As String)
    strLocalDespacho = lValor
End Property

Public Property Get LocalDespacho() As String
    LocalDespacho = strLocalDespacho
End Property



Public Static Property Get NumElementosCobro() As Byte
NumElementosCobro = byteNumElementosCobro
End Property

Public Static Property Let NumElementosCobro(ByVal lbyteNumElementosCobro As Byte)
byteNumElementosCobro = lbyteNumElementosCobro
End Property



Public Function Grabar(ByVal oradb As OraDatabase, ByRef pNumProforma As String, ByVal ArchivoVital As String) As String
Dim i As Integer
On Error GoTo Control
'If FormaPago.Count(1) <= 0 And CodigoConvenio = "" Then MsgBox "Ingrese forma de Pago", vbCritical, App.ProductName

byteNumElementosProd = IIf(Producto.UpperBound(1) < 0, 1, Producto.UpperBound(1)) + 1
byteNumElementosFormaPago = IIf(FormaPago.UpperBound(1) < 0, 1, FormaPago.UpperBound(1)) + 1
byteNumElementosCobro = IIf(CobroResponsabilidad.UpperBound(1) < 0, 1, CobroResponsabilidad.UpperBound(1)) + 1
byteNumElementosRM = IIf(RecetarioMagistral.UpperBound(1) < 0, 1, RecetarioMagistral.UpperBound(1)) + 1
byteNumElementosDatoAdic = IIf(DatosAdicional.UpperBound(1) < 0, 0, DatosAdicional.UpperBound(1)) + 1
    

'I.ECASTILLO 17.12.2020 | validar que lat tenga 1 solo guion, sino es erroneo
Dim countG As Long
Dim nVecesG As Long
For countG = 1 To Len(Latitud)
    If Mid(Latitud, countG, 1) = "-" Then nVecesG = nVecesG + 1
Next countG
If nVecesG <= 1 Then
Else
    Latitud = "": Longitud = ""
End If
'F.ECASTILLO 17.12.2020
    
    

    For i = oradb.Parameters.Count - 1 To 0 Step -1

        oradb.Parameters.Remove i

    Next

oradb.Parameters.Add "A_COD_CIA", objUsuario.CodigoEmpresa, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Compañia
oradb.Parameters.Add "A_COD_LOCAL", objUsuario.CodigoLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Codigo de local que emite la proforma
oradb.Parameters.Add "A_COD_MAQUINA", objUsuario.NombrePC, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Codigo de Maquina que genera la Proforma
oradb.Parameters.Add "A_COD_CLIENTE", Me.CodigoCliente, ORAPARM_INPUT, ORATYPE_VARCHAR2 ' Codigo del cliente a quien se le emite la proforma
oradb.Parameters.Add "A_COD_TIP_DOC_VTA", Me.CodigoDocumentoVenta, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Codigo del documento al cual se va a converitr la proforma
oradb.Parameters.Add "A_COD_EMPRESA", Me.CodEmpresa, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Codigo de la empresa a la cual se emitira la proforma
oradb.Parameters.Add "A_COD_CONVENIO", Me.CodigoConvenio, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Codigo del convenio al cual se emite la proforma
''Arturo Escate 03/02/2010 para corregir el problema de lo co-pagos de edelnor
If Not Me.PctBeneficiario = Round(Me.PctBeneficiarioReal, 0) Then
    oradb.Parameters.Add "A_PCT_BENEFICIARIO", Me.PctBeneficiarioReal, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Codigo del convenio al cual se emite la proforma
Else
    oradb.Parameters.Add "A_PCT_BENEFICIARIO", Me.PctBeneficiario, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Codigo del convenio al cual se emite la proforma
End If

oradb.Parameters.Add "A_COD_MEDICO", Me.CodMedico, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Codigo del medico que emite la receta
oradb.Parameters.Add "A_IMP_TIP_CAMBIO", objUsuario.TipoCambio, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Importe del tipo de cambio

'**************************************************************************************************************************************************************************************************************************************'
' se puso esta validacion para la asignacion de la meta de delivery 19/10/2007 Cristhian Rueda                                                                                                                                         '
'
' frm_VTA_Previa.pblnAsignoMeta = False => No se asigna meta de venta por parte del supervisor DLV                                                                                                                                     '
' frm_VTA_Previa.pblnAsignoMeta = True =>  Se asigna meta de venta por parte del supervisor DLV                                                                                                                                        '
'**************************************************************************************************************************************************************************************************************************************'
If frm_VTA_Previa.pblnAsignoMeta = False Then
    oradb.Parameters.Add "A_COD_USUARIO", objUsuario.Codigo, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Codigo del Usuario que emite la proforma
  Else
    oradb.Parameters.Add "A_COD_USUARIO", Trim(frm_VTA_Previa.pstrCodigoUsuario), ORAPARM_INPUT, ORATYPE_VARCHAR2 'Codigo del Usuario que emite la proforma
End If
'**************************************************************************************************************************************************************************************************************************************'
oradb.Parameters.Add "A_COD_UBIGEO", Me.UbigeoEntrega, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Codigo del Ubigeo

oradb.Parameters.Add "A_COD_TIPO_VENTA", objUsuario.CodTipoVenta, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Codigo del Ubigeo

'##oradb.Parameters.Add "A_COD_LOCAL_REF", Me.LocalAtencion, ORAPARM_INPUT, ORATYPE_VARCHAR2 'CODIGO DE LOCAL QUE LE VE A ATENDER
oradb.Parameters.Add "A_COD_LOCAL_REF", Me.LocalDespacho, ORAPARM_INPUT, ORATYPE_VARCHAR2  'CODIGO DE LOCAL QUE LE VE A ATENDER

oradb.Parameters.AddTable "A_CAD_COD_CAMPO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDatoAdic, 200
oradb.Parameters.AddTable "A_CAD_VAL_CAMPO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDatoAdic, 200
oradb.Parameters.Add "A_CAD_COD_DIAGNOSTICO", Me.DetalleReceta, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_FCH_RECETA", Me.FechaReceta, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_LOCAL_RECETA", Me.LocalReceta, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_CLIENTE_DLV", Me.CodigoClienteDLV, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_FLG_TRANSF_LOCAL", Me.FlgTransfLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2 'flag de tranferencia de local
'I.ECASTILLO 17.12.2020
Dim flg_ruteoA_cnv
flg_ruteoA_cnv = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRACNV") '1 => ACTIVO, 0 => INACTIVO
If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
    GoTo cnvNoRuteaAuto
End If
Dim flg_2e_reserva
flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV3") '1 => ACTIVO, 0 => INACTIVO
If flg_2e_reserva = "0" Then
cnvNoRuteaAuto:
    oradb.Parameters.Add "A_FLG_ENTREGA_LOCAL", Me.FlgEntregaLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2 'flag de que se entregara en en local
Else
    oradb.Parameters.Add "A_FLG_ENTREGA_LOCAL", Me.bk_chkRET, ORAPARM_INPUT, ORATYPE_VARCHAR2 'flag de que se entregara en en local
    oradb.Parameters.Add "A_DELIVERY_TIME", Me.bk_deliveryTime, ORAPARM_INPUT, ORATYPE_VARCHAR2
End If
'oradb.Parameters.Add "A_FLG_ENTREGA_LOCAL", Me.FlgEntregaLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2 'flag de que se entregara en en local
'F.ECASTILLO 17.12.2020
oradb.Parameters.Add "A_FLG_RESERVA_STK", Me.FlgReservaStk, ORAPARM_INPUT, ORATYPE_VARCHAR2 ' flag de reserva de stock
oradb.Parameters.Add "A_FLG_ENTREGA_TERCERO", Me.EntregaTercero, ORAPARM_INPUT, ORATYPE_VARCHAR2  ' flag de entrega a terceros
'flg_2e_reserva = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "FLGRSVDLV2") '1 => ACTIVO, 0 => INACTIVO
If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
    GoTo cnvNoRuteaAuto2
End If
If flg_2e_reserva = "0" Then
cnvNoRuteaAuto2:
    oradb.Parameters.Add "A_FLG_FECHA_PACTADA", Me.FlgPactado, ORAPARM_INPUT, ORATYPE_VARCHAR2  ' flag de fecha pactada "Por Cristhian rueda 07/11/2007"
Else
    oradb.Parameters.Add "A_FLG_FECHA_PACTADA", Me.bk_flgPactado, ORAPARM_INPUT, ORATYPE_VARCHAR2  ' flag de fecha pactada "Por Cristhian rueda 07/11/2007"
End If
oradb.Parameters.Add "A_COD_DIRECCION_CLI", Me.CodDireccionCli, ORAPARM_INPUT, ORATYPE_VARCHAR2 'flag de direccion del cliente
oradb.Parameters.Add "A_COD_MOTORIZADO", Me.CodMotorizado, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Codigo del motorizado al que se asigno
oradb.Parameters.Add "A_FLG_URGENTE", Me.FlgUrgente, ORAPARM_INPUT, ORATYPE_VARCHAR2 'flag urgente
oradb.Parameters.Add "A_DES_AUX_CLI_TLF", Me.DesAuxCliTlf, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_DES_AUX_CLI_NOMBRE", Me.NombreClienteDLV, ORAPARM_INPUT, ORATYPE_VARCHAR2
''oradb.Parameters.Add "A_DES_AUX_CLI_NOMBRE", Out_NombreCliente, ORAPARM_INPUT, ORATYPE_VARCHAR2
Dim strSuf As String
Dim valI As Integer
valI = InStr(1, Me.DireccionClienteDLV, objVenta.bk_SufijoDir)
If valI > 0 Then Me.DireccionClienteDLV = Trim(Replace(Me.DireccionClienteDLV, objVenta.bk_SufijoDir, "", valI, 1))
oradb.Parameters.Add "A_DES_AUX_CLI_DIRECC", Me.DireccionClienteDLV, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_DES_AUX_CLI_REF", Me.DesReferenciaCli, ORAPARM_INPUT, ORATYPE_VARCHAR2
'I.ECASTILLO 17.12.2020
Dim fchPactada
Dim segundaHora
If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
    GoTo cnvNoRuteaAuto3
End If
If flg_2e_reserva = "0" Then
cnvNoRuteaAuto3:
    oradb.Parameters.Add "A_FCH_HORA_PACT_ENTR", Me.FchHoraPactEntr, ORAPARM_INPUT, ORATYPE_VARCHAR2
    oradb.Parameters.Add "A_HORA_SEGUNDA_PACT_ENTR", Null, ORAPARM_INPUT, ORATYPE_VARCHAR2
Else
    fchPactada = Format(objVenta.bk_FechaCapacidad, "dd/mm/yyyy ") & Format(objVenta.bk_HoraCapacidad, "hh:mm:ss")
    segundaHora = Format(objVenta.bk_HoraCapacidad2, "hh:mm:ss")
    oradb.Parameters.Add "A_FCH_HORA_PACT_ENTR", fchPactada, ORAPARM_INPUT, ORATYPE_VARCHAR2 'flag de que se entregara en en local
    oradb.Parameters.Add "A_HORA_SEGUNDA_PACT_ENTR", segundaHora, ORAPARM_INPUT, ORATYPE_VARCHAR2 'flag de que se entregara en en local
End If
'oradb.Parameters.Add "A_FCH_HORA_PACT_ENTR", Me.FchHoraPactEntr, ORAPARM_INPUT, ORATYPE_VARCHAR2
'F.ECASTILLO 17.12.2020
oradb.Parameters.Add "A_FCH_HORA_PACT_RECOG", Me.FchHoraPactRecog, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_DES_AUX_RECOGE_TLF", Me.DesAuxRecogeTlf, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_DES_AUX_RECOGE_NOMBRE", Me.DesAuxRecogeNombre, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_DES_AUX_RECOGE_DIRECC", Me.DesAuxRecogeDirecc, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_DES_AUX_RECOGE_REF", Me.DesAuxRecogeRef, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_DES_AUX_MOTORIZADO", Me.DesAuxMotorizado, ORAPARM_INPUT, ORATYPE_VARCHAR2


If objVenta.Out_Tipo = "1" Then
    oradb.Parameters.Add "A_NUM_RUC_EMPRESA", Out_NumeroId, ORAPARM_INPUT, ORATYPE_VARCHAR2
    oradb.Parameters.Add "A_DES_RAZON_SOCIAL", Out_NombreCliente, ORAPARM_INPUT, ORATYPE_VARCHAR2
    'Linea agregada por DJARA para guardar la direccion social para la facturacion
    oradb.Parameters.Add "A_DES_DIREC_SOCIAL", Out_Direccion, ORAPARM_INPUT, ORATYPE_VARCHAR2
Else
    oradb.Parameters.Add "A_NUM_RUC_EMPRESA", Out_NumeroId, ORAPARM_INPUT, ORATYPE_VARCHAR2
    oradb.Parameters.Add "A_DES_RAZON_SOCIAL", "", ORAPARM_INPUT, ORATYPE_VARCHAR2
    'Linea agregada por DJARA
    oradb.Parameters.Add "A_DES_DIREC_SOCIAL", Out_Direccion, ORAPARM_INPUT, ORATYPE_VARCHAR2
End If



'oradb.Parameters.Add "A_OBS_NOTA_LOCAL", ObsNotaLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_OBS_NOTA_LOCAL", ObservacionLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_OBS_NOTA_VERIFICACION", ObsNotaVerificacion, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_OBS_NOTA_RUTEO", ObsNotaRuteo, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_OBS_NOTA_MOTORIZADO", ObsNotaMotorizado, ORAPARM_INPUT, ORATYPE_VARCHAR2

oradb.Parameters.Add "A_NUM_PROFORMA_ORIG", "", ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_FLG_COMMIT", "1", ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_NOM_TITULAR", NomTitular, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_NUM_DNI", NumDNI, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_MOTORIZADO_CONV", CodMotorizadoConv, ORAPARM_INPUT, ORATYPE_VARCHAR2

oradb.Parameters.AddTable "A_CAD_COD_USU_RESP", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosCobro, 200
oradb.Parameters.AddTable "A_CAD_IMP_USU_RESP", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosCobro, 200

'FORMA DE PAGO
oradb.Parameters.AddTable "A_CAD_COD_FORM_PAGO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_COD_FPAG_HIJO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_COD_MONEDA_PAGO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_TIP_CAMBIO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_COD_TIP_TARJETA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_NUM_TARJETA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_FLG_TIP_CUOTA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_NUM_CUOTA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_FCH_VENCIMIENTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_MTO_RETIRO_EFEC", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_COD_DOC_PAGO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_NUM_DOC_PAGO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_NUM_DOC_NOTA_CRED", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_FCH_DOC_NOTA_CRED", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_COD_DOC_IDENT", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_NUM_DOC_IDENT", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_COD_BANCO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_NUM_MOVIMIENTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_FCH_MOVIMIENTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_CTA_CTE_BTL", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_COD_DONACION", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_IMPORTE", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
'Se agrego campo Nombre Titular para la compra con tarjeta de credito en DLV'
'Fecha 10/10/2007 Por Cristhian Rueda'
oradb.Parameters.AddTable "A_CAD_NOM_TITULAR", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_NUM_VALE", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200

'DETALLE DE PRODUCTOS DE PROF
oradb.Parameters.AddTable "A_CAD_TIP_VENTA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200
oradb.Parameters.AddTable "A_CAD_COD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200
oradb.Parameters.AddTable "A_CAD_CTD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200
oradb.Parameters.AddTable "A_CAD_FLG_CTD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200
oradb.Parameters.AddTable "A_CAD_SUBT_NETO_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200

'Recetario Magistral'
'Graba Proforma'
oradb.Parameters.AddTable "A_CAD2_TIP_VENTA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200
oradb.Parameters.AddTable "A_CAD2_COD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200
oradb.Parameters.AddTable "A_CAD2_COD_PRODUCTO_INS", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200
oradb.Parameters.AddTable "A_CAD2_CTD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200
oradb.Parameters.AddTable "A_CAD2_FLG_CTD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200
oradb.Parameters.AddTable "A_CAD2_PRC_UNIT_INSUMO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200
oradb.Parameters.AddTable "A_CAD2_PCT_BASE", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200
oradb.Parameters.AddTable "A_CAD2_CTD_BASE", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200
oradb.Parameters.AddTable "A_CAD2_UND_MEDIDA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200

oradb.Parameters.Add "A_NUM_PROFORMA_GENERADA", "", ORAPARM_OUTPUT, ORATYPE_VARCHAR2
'oradb.Parameters.Add "A_NUM_RECETA",ME., ORAPARM_INPUT, ORATYPE_VARCHAR2

'-- Se creo el parametro para lo que es el registro del numero de autorizacion para la formna de pago con tarjeta --'
'-- 10/01/2008 Por Cristhian Rueda                                                                                --'
oradb.Parameters.AddTable "A_CAD_NUM_AUTORIZACION", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_FLG_PROMOCION", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200



oradb.Parameters.Add "A_CARI_ID", "", ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.AddTable "A_CAD_COD_PROMOCION", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200
oradb.Parameters.AddTable "A_CAD_COD_PROMOCION_DCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200

oradb.Parameters.Add "A_NUM_ARCHIVO", ArchivoVital, ORAPARM_INPUT, ORATYPE_VARCHAR2 '-- Parámetro donde guardar el número del archivo de plan vital en la tabla de proformas para relacionarlas 27/01/2010



' SE AGREGAN LOS VALORES DEL MAPA
oradb.Parameters.Add "A_LATITUD", Latitud, ORAPARM_INPUT, ORATYPE_VARCHAR2 'LATITUD

oradb.Parameters.Add "A_LONGITUD", Longitud, ORAPARM_INPUT, ORATYPE_VARCHAR2 'LONGITUD

' FIN  VALORES MAPAS
oradb.Parameters.Add "A_TIPO_DOCUMENTO_PTS", strTipoDocumento_Pts, ORAPARM_INPUT, ORATYPE_VARCHAR2 '-- tipo documento para otorgar puntos
oradb.Parameters.Add "A_NUM_DOCUMENTO_PTS", strNumDocumento_Pts, ORAPARM_INPUT, ORATYPE_VARCHAR2 '-- documento al cual le otorga puntos
'I.ECASTILLO 27.10.2020
' 17.12.2020 | se comenta pero volver a usar
oradb.Parameters.Add "A_FLG_RESERVA_CAPACIDAD", objVenta.flgReservoCap, ORAPARM_INPUT, ORATYPE_VARCHAR2
'F.ECASTILLO 27.10.2020

oradb.Parameters.Add "A_COD_LOCAL_CALL_CENTER", objUsuario.CodLocalCallCenter, ORAPARM_INPUT, ORATYPE_VARCHAR2  '-- GNIBIN 20210210 Proyecto MultiMarca
oradb.Parameters.Add "A_DELIVERY_TYPE", bk_ServiceType, ORAPARM_INPUT, ORATYPE_VARCHAR2 '--ECASTILLO 17.12.2020
If gstrIndRAv3 = "2" Then
oradb.Parameters.Add "A_IND_RESERVE_CAPACIDAD", gstrFlagReservaCap, ORAPARM_INPUT, ORATYPE_VARCHAR2 '--ECASTILLO 14.07.2021
oradb.Parameters.Add "A_START_HOUR", Me.bk_starHour, ORAPARM_INPUT, ORATYPE_VARCHAR2 '--ECASTILLO 14.07.2021
oradb.Parameters.Add "A_END_HOUR", Me.bk_endHour, ORAPARM_INPUT, ORATYPE_VARCHAR2 '--ECASTILLO 14.07.2021
End If
Dim CodUsuResp As OracleInProcServer.OraParamArray
Dim ImpUsuResp As OracleInProcServer.OraParamArray
Dim CodFormPago As OracleInProcServer.OraParamArray
Dim CodMonedaPago As OracleInProcServer.OraParamArray
Dim TipCambio As OracleInProcServer.OraParamArray
Dim CodTipTarjeta As OracleInProcServer.OraParamArray
Dim NumTarjeta As OracleInProcServer.OraParamArray
Dim FlgTipCuota As OracleInProcServer.OraParamArray
Dim NumCuota As OracleInProcServer.OraParamArray
Dim FchVencimiento As OracleInProcServer.OraParamArray
Dim NumAutorizacion As OracleInProcServer.OraParamArray
Dim CodDocPago As OracleInProcServer.OraParamArray
Dim NumDocPago As OracleInProcServer.OraParamArray
Dim NumDocNotaCred As OracleInProcServer.OraParamArray
Dim FchDocNotaCred As OracleInProcServer.OraParamArray
Dim CodDocIdent As OracleInProcServer.OraParamArray
Dim NumDocIdent As OracleInProcServer.OraParamArray
Dim CodBanco As OracleInProcServer.OraParamArray
Dim NumMovimiento As OracleInProcServer.OraParamArray
Dim FchMovimiento As OracleInProcServer.OraParamArray
Dim CtaCteBtl As OracleInProcServer.OraParamArray
Dim CodDonacion As OracleInProcServer.OraParamArray
Dim Importe As OracleInProcServer.OraParamArray
Dim TipVenta As OracleInProcServer.OraParamArray
Dim codProducto As OracleInProcServer.OraParamArray
Dim CtdProducto As OracleInProcServer.OraParamArray
Dim FlgCtdProducto As OracleInProcServer.OraParamArray
Dim SubtNetoProducto As OracleInProcServer.OraParamArray
Dim CodFPagHijo As OracleInProcServer.OraParamArray
Dim MtoRetiroEfec As OracleInProcServer.OraParamArray
'Cambio 10/10/2007 Por Cristhian Rueda'
Dim NombTitular As OracleInProcServer.OraParamArray
Dim NumVale As OracleInProcServer.OraParamArray 'ECASTILLO 13.03.2020
'''''''''''''''''''''''
''Recetario Magistral''
Dim Tip2Venta As OracleInProcServer.OraParamArray
Dim Cod2Producto As OracleInProcServer.OraParamArray
Dim Cod2ProductoIns As OracleInProcServer.OraParamArray
Dim Ctd2Producto As OracleInProcServer.OraParamArray
Dim Flg2CtdProducto As OracleInProcServer.OraParamArray
Dim Prc2UnitInsumo As OracleInProcServer.OraParamArray
Dim Ctd2PctBase As OracleInProcServer.OraParamArray
Dim Ctd2Base As OracleInProcServer.OraParamArray
Dim UndMedida As OracleInProcServer.OraParamArray

Dim CadenaCodigoPromocion As OracleInProcServer.OraParamArray
Dim CadenaCodigoPromocionDcto As OracleInProcServer.OraParamArray

'''''''''''''''''''''
''''Datos Adicionales
Dim CodCampo As OracleInProcServer.OraParamArray
Dim ValCampo As OracleInProcServer.OraParamArray

Dim FlgPromocion As OracleInProcServer.OraParamArray ''Arturo Escate

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ '
Set CodUsuResp = oradb.Parameters("A_CAD_COD_USU_RESP")
Set ImpUsuResp = oradb.Parameters("A_CAD_IMP_USU_RESP")

Set CodFormPago = oradb.Parameters("A_CAD_COD_FORM_PAGO")
Set CodFPagHijo = oradb.Parameters("A_CAD_COD_FPAG_HIJO")
Set CodMonedaPago = oradb.Parameters("A_CAD_COD_MONEDA_PAGO")
Set TipCambio = oradb.Parameters("A_CAD_TIP_CAMBIO")
Set CodTipTarjeta = oradb.Parameters("A_CAD_COD_TIP_TARJETA")
Set NumTarjeta = oradb.Parameters("A_CAD_NUM_TARJETA")
Set FlgTipCuota = oradb.Parameters("A_CAD_FLG_TIP_CUOTA")
Set NumCuota = oradb.Parameters("A_CAD_NUM_CUOTA")
Set FchVencimiento = oradb.Parameters("A_CAD_FCH_VENCIMIENTO")
Set NumAutorizacion = oradb.Parameters("A_CAD_NUM_AUTORIZACION")   'Agregado 11/01/2008 Por Cristhian R.


Set MtoRetiroEfec = oradb.Parameters("A_CAD_MTO_RETIRO_EFEC")
Set CodDocPago = oradb.Parameters("A_CAD_COD_DOC_PAGO")
Set NumDocPago = oradb.Parameters("A_CAD_NUM_DOC_PAGO")
Set NumDocNotaCred = oradb.Parameters("A_NUM_DOC_NOTA_CRED")
Set FchDocNotaCred = oradb.Parameters("A_FCH_DOC_NOTA_CRED")
Set CodDocIdent = oradb.Parameters("A_CAD_COD_DOC_IDENT")
Set NumDocIdent = oradb.Parameters("A_CAD_NUM_DOC_IDENT")
Set CodBanco = oradb.Parameters("A_CAD_COD_BANCO")
Set NumMovimiento = oradb.Parameters("A_CAD_NUM_MOVIMIENTO")
Set FchMovimiento = oradb.Parameters("A_CAD_FCH_MOVIMIENTO")
Set CtaCteBtl = oradb.Parameters("A_CAD_CTA_CTE_BTL")
Set CodDonacion = oradb.Parameters("A_CAD_COD_DONACION")
Set Importe = oradb.Parameters("A_CAD_IMPORTE")
Set TipVenta = oradb.Parameters("A_CAD_TIP_VENTA")
Set codProducto = oradb.Parameters("A_CAD_COD_PRODUCTO")
Set CtdProducto = oradb.Parameters("A_CAD_CTD_PRODUCTO")
Set FlgCtdProducto = oradb.Parameters("A_CAD_FLG_CTD_PRODUCTO")
Set SubtNetoProducto = oradb.Parameters("A_CAD_SUBT_NETO_PRODUCTO")
Set NombTitular = oradb.Parameters("A_CAD_NOM_TITULAR")
Set NumVale = oradb.Parameters("A_CAD_NUM_VALE")
Set CadenaCodigoPromocion = oradb.Parameters("A_CAD_COD_PROMOCION")
Set CadenaCodigoPromocionDcto = oradb.Parameters("A_CAD_COD_PROMOCION_DCTO")



''Recetario Magistral''
Set Tip2Venta = oradb.Parameters("A_CAD2_TIP_VENTA")
Set Cod2Producto = oradb.Parameters("A_CAD2_COD_PRODUCTO")
Set Cod2ProductoIns = oradb.Parameters("A_CAD2_COD_PRODUCTO_INS")
Set Ctd2Producto = oradb.Parameters("A_CAD2_CTD_PRODUCTO")
Set Flg2CtdProducto = oradb.Parameters("A_CAD2_FLG_CTD_PRODUCTO")
Set Prc2UnitInsumo = oradb.Parameters("A_CAD2_PRC_UNIT_INSUMO")
Set Ctd2PctBase = oradb.Parameters("A_CAD2_PCT_BASE")
Set Ctd2Base = oradb.Parameters("A_CAD2_CTD_BASE")
Set UndMedida = oradb.Parameters("A_CAD2_UND_MEDIDA")


''Datos Adicionales
Set CodCampo = oradb.Parameters("A_CAD_COD_CAMPO")
Set ValCampo = oradb.Parameters("A_CAD_VAL_CAMPO")


Set FlgPromocion = oradb.Parameters("A_CAD_FLG_PROMOCION")

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ '

'-------------------------DATO ADICIONAL-------------------------------

    For i = 0 To DatosAdicional.UpperBound(1)
        CodCampo(i) = DatosAdicional(i, 0)
        ValCampo(i) = DatosAdicional(i, 5)
    Next


'------------------------COBRO POR RESPONSABILIDAD----------------------------
    
    
    For i = 0 To CobroResponsabilidad.UpperBound(1)
        CodUsuResp(i) = CobroResponsabilidad(i, 0)
        ImpUsuResp(i) = CobroResponsabilidad(i, 2)
    Next

'------------------------FORMA DE PAGO----------------------------------------
    
    
    
    
    
        For i = 0 To FormaPago.UpperBound(1)
            CodFormPago(i) = FormaPago(i, 0)
            CodFPagHijo(i) = FormaPago(i, 2)
            CodMonedaPago(i) = FormaPago(i, 6)
            TipCambio(i) = FormaPago(i, 11)
            CodTipTarjeta(i) = FormaPago(i, 5)
            NumTarjeta(i) = FormaPago(i, 12)
            FlgTipCuota(i) = FormaPago(i, 15)
            NumCuota(i) = FormaPago(i, 13)
            FchVencimiento(i) = FormaPago(i, 14)
            CodDocPago(i) = FormaPago(i, 7)
            NumDocPago(i) = FormaPago(i, 16)
            NumDocNotaCred(i) = FormaPago(i, 22)
            FchDocNotaCred(i) = FormaPago(i, 21)
            CodDocIdent(i) = ""
            NumDocIdent(i) = FormaPago(i, 25)
            CodBanco(i) = FormaPago(i, 8)
            NumMovimiento(i) = FormaPago(i, 17)
            FchMovimiento(i) = FormaPago(i, 19)
            CtaCteBtl(i) = FormaPago(i, 10)
            CodDonacion(i) = FormaPago(i, 9)
            Importe(i) = FormaPago(i, 4)
            MtoRetiroEfec(i) = FormaPago(i, 29)      'cambiar por el verdadero
            NombTitular(i) = FormaPago(i, 33)
            NumVale(i) = FormaPago(i, 34) 'ECASTILLO 13.03.2020
            NumAutorizacion(i) = FormaPago(i, 20)    'Numero de Autorización Agregado 11/01/2007 Por Cristhian R.
        Next

'------------------------PARA EL DETALLE DE PRODUCTOS DE PROF-----------------
    
    
    
    
    For i = 0 To Producto.UpperBound(1)
        TipVenta(i) = Producto(i, 5)
        codProducto(i) = Producto(i, 0)
        CtdProducto(i) = Producto(i, 3)
        FlgCtdProducto(i) = Producto(i, 6)
        'PrcUnitProducto(i) = Producto(i, 4)
        'PrcNetoProducto(i) = Producto(i, 4)
        SubtNetoProducto(i) = Producto(i, 4)
        FlgPromocion(i) = Producto(i, 7)
        CadenaCodigoPromocion(i) = Producto(i, 26)
        CadenaCodigoPromocionDcto(i) = Producto(i, 27)
    Next

'------------------------PARA EL CASO DE RECETARIO MAGISTRAL------------------


    
    For i = 0 To RecetarioMagistral.UpperBound(1)
        Cod2Producto(i) = RecetarioMagistral(i, 0)
        Cod2ProductoIns(i) = RecetarioMagistral(i, 1)
        Ctd2Producto(i) = RecetarioMagistral(i, 3)
        Flg2CtdProducto(i) = "0"
        Prc2UnitInsumo(i) = RecetarioMagistral(i, 9)
        Tip2Venta(i) = RecetarioMagistral(i, 5)
        Ctd2PctBase(i) = RecetarioMagistral(i, 7)
        Ctd2Base(i) = RecetarioMagistral(i, 6)
        UndMedida(i) = RecetarioMagistral(i, 8)
    Next





    If flg_ruteoA_cnv <> "1" And objVenta.ptmModalidad = Venta_Convenio Then
        GoTo cnvNoRuteaAuto4
    End If
    If flg_2e_reserva = "0" Then
cnvNoRuteaAuto4:
        oradb.ExecuteSQL " BEGIN BTLPROD.PKG_PROFORMA_V3.SP_GRABA (:A_COD_CIA                  ,:A_COD_LOCAL               ,:A_COD_MAQUINA             ," _
                                                         & ":A_COD_CLIENTE               ,:A_COD_TIP_DOC_VTA         ,:A_COD_EMPRESA             ," _
                                                         & ":A_COD_CONVENIO              ,:A_PCT_BENEFICIARIO        ,:A_COD_MEDICO              ," _
                                                         & ":A_IMP_TIP_CAMBIO            ,:A_COD_USUARIO             ,:A_COD_UBIGEO              ,:A_COD_TIPO_VENTA , :A_COD_LOCAL_REF             ," _
                                                         & ":A_CAD_COD_CAMPO             ,:A_CAD_VAL_CAMPO           ,:A_CAD_COD_DIAGNOSTICO     ,:A_FCH_RECETA              ,:A_COD_LOCAL_RECETA , :A_COD_CLIENTE_DLV ,:A_FLG_TRANSF_LOCAL        ,:A_FLG_ENTREGA_LOCAL       ," _
                                                         & ":A_FLG_RESERVA_STK           ,:A_COD_DIRECCION_CLI       ,:A_COD_MOTORIZADO          ," _
                                                         & ":A_FLG_URGENTE               ,:A_DES_AUX_CLI_TLF         ,:A_DES_AUX_CLI_NOMBRE      ," _
                                                         & ":A_DES_AUX_CLI_DIRECC        ,:A_DES_AUX_CLI_REF         ,:A_FCH_HORA_PACT_ENTR      ,:A_FCH_HORA_PACT_RECOG     ," _
                                                         & ":A_DES_AUX_RECOGE_TLF        ,:A_DES_AUX_RECOGE_NOMBRE   ,:A_DES_AUX_RECOGE_DIRECC   ," _
                                                         & ":A_DES_AUX_RECOGE_REF        ,:A_DES_AUX_MOTORIZADO      ,:A_NUM_RUC_EMPRESA           ,:A_DES_RAZON_SOCIAL        ," _
                                                         & ":A_OBS_NOTA_LOCAL            ,:A_OBS_NOTA_VERIFICACION   ,:A_OBS_NOTA_RUTEO            ,:A_OBS_NOTA_MOTORIZADO     ," _
                                                         & ":A_CAD_COD_USU_RESP          ," _
                                                         & ":A_CAD_IMP_USU_RESP          ,:A_CAD_COD_FORM_PAGO       ,:A_CAD_COD_FPAG_HIJO       ," _
                                                         & ":A_CAD_COD_MONEDA_PAGO       ,:A_CAD_TIP_CAMBIO          ,:A_CAD_COD_TIP_TARJETA     ," _
                                                         & ":A_CAD_NUM_TARJETA           ,:A_CAD_FLG_TIP_CUOTA       ,:A_CAD_NUM_CUOTA           ," _
                                                         & ":A_CAD_FCH_VENCIMIENTO       ,:A_CAD_MTO_RETIRO_EFEC     ,:A_CAD_COD_DOC_PAGO        ," _
                                                         & ":A_CAD_NUM_DOC_PAGO          ,:A_NUM_DOC_NOTA_CRED       ,:A_FCH_DOC_NOTA_CRED       ," _
                                                         & ":A_CAD_COD_DOC_IDENT         ,:A_CAD_NUM_DOC_IDENT       ,:A_CAD_NOM_TITULAR         ,:A_CAD_COD_BANCO           ," _
                                                         & ":A_CAD_NUM_MOVIMIENTO        ,:A_CAD_FCH_MOVIMIENTO      ,:A_CAD_CTA_CTE_BTL         ," _
                                                         & ":A_CAD_COD_DONACION          ,:A_CAD_IMPORTE             ,:A_CAD_TIP_VENTA           ," _
                                                         & ":A_CAD_COD_PRODUCTO          ,:A_CAD_CTD_PRODUCTO        ,:A_CAD_FLG_CTD_PRODUCTO    ," _
                                                         & ":A_CAD_SUBT_NETO_PRODUCTO    ,:A_CAD2_TIP_VENTA          ,:A_CAD2_COD_PRODUCTO       ," _
                                                         & ":A_CAD2_COD_PRODUCTO_INS     ,:A_CAD2_CTD_PRODUCTO       ,:A_CAD2_FLG_CTD_PRODUCTO   ," _
                                                         & ":A_CAD2_PRC_UNIT_INSUMO      ,:A_CAD2_PCT_BASE           ,:A_CAD2_CTD_BASE           ," _
                                                         & ":A_CAD2_UND_MEDIDA           ,:A_NUM_PROFORMA_GENERADA   ,:A_NUM_PROFORMA_ORIG       ,:A_FLG_COMMIT ,:A_NOM_TITULAR, :A_NUM_DNI, :A_FLG_ENTREGA_TERCERO, :A_FLG_FECHA_PACTADA, :A_COD_MOTORIZADO_CONV, NULL,:A_CAD_NUM_AUTORIZACION, :A_CAD_FLG_PROMOCION, :A_DES_DIREC_SOCIAL, :A_CARI_ID, :A_CAD_COD_PROMOCION,  :A_CAD_COD_PROMOCION_DCTO, null, null, null, :A_NUM_ARCHIVO, :A_LATITUD, :A_LONGITUD, :A_TIPO_DOCUMENTO_PTS, :A_NUM_DOCUMENTO_PTS, :A_CAD_NUM_VALE, :A_FLG_RESERVA_CAPACIDAD, :A_COD_LOCAL_CALL_CENTER) ; END;"
                                                         
    Else
        If gstrIndRAv3 = "0" Or gstrIndRAv3 = "" Then
gstrIndRAv3Other:
            oradb.ExecuteSQL " BEGIN BTLPROD.PKG_PROFORMA_V3.SP_GRABA (:A_COD_CIA                  ,:A_COD_LOCAL               ,:A_COD_MAQUINA             ," _
                                                             & ":A_COD_CLIENTE               ,:A_COD_TIP_DOC_VTA         ,:A_COD_EMPRESA             ," _
                                                             & ":A_COD_CONVENIO              ,:A_PCT_BENEFICIARIO        ,:A_COD_MEDICO              ," _
                                                             & ":A_IMP_TIP_CAMBIO            ,:A_COD_USUARIO             ,:A_COD_UBIGEO              ,:A_COD_TIPO_VENTA , :A_COD_LOCAL_REF             ," _
                                                             & ":A_CAD_COD_CAMPO             ,:A_CAD_VAL_CAMPO           ,:A_CAD_COD_DIAGNOSTICO     ,:A_FCH_RECETA              ,:A_COD_LOCAL_RECETA , :A_COD_CLIENTE_DLV ,:A_FLG_TRANSF_LOCAL        ,:A_FLG_ENTREGA_LOCAL       ," _
                                                             & ":A_FLG_RESERVA_STK           ,:A_COD_DIRECCION_CLI       ,:A_COD_MOTORIZADO          ," _
                                                             & ":A_FLG_URGENTE               ,:A_DES_AUX_CLI_TLF         ,:A_DES_AUX_CLI_NOMBRE      ," _
                                                             & ":A_DES_AUX_CLI_DIRECC        ,:A_DES_AUX_CLI_REF         ,:A_FCH_HORA_PACT_ENTR      ,:A_FCH_HORA_PACT_RECOG     ," _
                                                             & ":A_DES_AUX_RECOGE_TLF        ,:A_DES_AUX_RECOGE_NOMBRE   ,:A_DES_AUX_RECOGE_DIRECC   ," _
                                                             & ":A_DES_AUX_RECOGE_REF        ,:A_DES_AUX_MOTORIZADO      ,:A_NUM_RUC_EMPRESA           ,:A_DES_RAZON_SOCIAL        ," _
                                                             & ":A_OBS_NOTA_LOCAL            ,:A_OBS_NOTA_VERIFICACION   ,:A_OBS_NOTA_RUTEO            ,:A_OBS_NOTA_MOTORIZADO     ," _
                                                             & ":A_CAD_COD_USU_RESP          ," _
                                                             & ":A_CAD_IMP_USU_RESP          ,:A_CAD_COD_FORM_PAGO       ,:A_CAD_COD_FPAG_HIJO       ," _
                                                             & ":A_CAD_COD_MONEDA_PAGO       ,:A_CAD_TIP_CAMBIO          ,:A_CAD_COD_TIP_TARJETA     ," _
                                                             & ":A_CAD_NUM_TARJETA           ,:A_CAD_FLG_TIP_CUOTA       ,:A_CAD_NUM_CUOTA           ," _
                                                             & ":A_CAD_FCH_VENCIMIENTO       ,:A_CAD_MTO_RETIRO_EFEC     ,:A_CAD_COD_DOC_PAGO        ," _
                                                             & ":A_CAD_NUM_DOC_PAGO          ,:A_NUM_DOC_NOTA_CRED       ,:A_FCH_DOC_NOTA_CRED       ," _
                                                             & ":A_CAD_COD_DOC_IDENT         ,:A_CAD_NUM_DOC_IDENT       ,:A_CAD_NOM_TITULAR         ,:A_CAD_COD_BANCO           ," _
                                                             & ":A_CAD_NUM_MOVIMIENTO        ,:A_CAD_FCH_MOVIMIENTO      ,:A_CAD_CTA_CTE_BTL         ," _
                                                             & ":A_CAD_COD_DONACION          ,:A_CAD_IMPORTE             ,:A_CAD_TIP_VENTA           ," _
                                                             & ":A_CAD_COD_PRODUCTO          ,:A_CAD_CTD_PRODUCTO        ,:A_CAD_FLG_CTD_PRODUCTO    ," _
                                                             & ":A_CAD_SUBT_NETO_PRODUCTO    ,:A_CAD2_TIP_VENTA          ,:A_CAD2_COD_PRODUCTO       ," _
                                                             & ":A_CAD2_COD_PRODUCTO_INS     ,:A_CAD2_CTD_PRODUCTO       ,:A_CAD2_FLG_CTD_PRODUCTO   ," _
                                                             & ":A_CAD2_PRC_UNIT_INSUMO      ,:A_CAD2_PCT_BASE           ,:A_CAD2_CTD_BASE           ," _
                                                             & ":A_CAD2_UND_MEDIDA           ,:A_NUM_PROFORMA_GENERADA   ,:A_NUM_PROFORMA_ORIG       ,:A_FLG_COMMIT ,:A_NOM_TITULAR, :A_NUM_DNI, :A_FLG_ENTREGA_TERCERO, :A_FLG_FECHA_PACTADA, :A_COD_MOTORIZADO_CONV, NULL,:A_CAD_NUM_AUTORIZACION, :A_CAD_FLG_PROMOCION, :A_DES_DIREC_SOCIAL, :A_CARI_ID, :A_CAD_COD_PROMOCION,  :A_CAD_COD_PROMOCION_DCTO, null, null, null, :A_NUM_ARCHIVO, :A_LATITUD, :A_LONGITUD, :A_TIPO_DOCUMENTO_PTS, :A_NUM_DOCUMENTO_PTS, :A_CAD_NUM_VALE, :A_FLG_RESERVA_CAPACIDAD, :A_COD_LOCAL_CALL_CENTER, :A_DELIVERY_TYPE, :A_HORA_SEGUNDA_PACT_ENTR, :A_DELIVERY_TIME) ; END;"
        ElseIf gstrIndRAv3 = "2" Then
            oradb.ExecuteSQL " BEGIN BTLPROD.PKG_PROFORMA_V3.SP_GRABA (:A_COD_CIA                  ,:A_COD_LOCAL               ,:A_COD_MAQUINA             ," _
                                                             & ":A_COD_CLIENTE               ,:A_COD_TIP_DOC_VTA         ,:A_COD_EMPRESA             ," _
                                                             & ":A_COD_CONVENIO              ,:A_PCT_BENEFICIARIO        ,:A_COD_MEDICO              ," _
                                                             & ":A_IMP_TIP_CAMBIO            ,:A_COD_USUARIO             ,:A_COD_UBIGEO              ,:A_COD_TIPO_VENTA , :A_COD_LOCAL_REF             ," _
                                                             & ":A_CAD_COD_CAMPO             ,:A_CAD_VAL_CAMPO           ,:A_CAD_COD_DIAGNOSTICO     ,:A_FCH_RECETA              ,:A_COD_LOCAL_RECETA , :A_COD_CLIENTE_DLV ,:A_FLG_TRANSF_LOCAL        ,:A_FLG_ENTREGA_LOCAL       ," _
                                                             & ":A_FLG_RESERVA_STK           ,:A_COD_DIRECCION_CLI       ,:A_COD_MOTORIZADO          ," _
                                                             & ":A_FLG_URGENTE               ,:A_DES_AUX_CLI_TLF         ,:A_DES_AUX_CLI_NOMBRE      ," _
                                                             & ":A_DES_AUX_CLI_DIRECC        ,:A_DES_AUX_CLI_REF         ,:A_FCH_HORA_PACT_ENTR      ,:A_FCH_HORA_PACT_RECOG     ," _
                                                             & ":A_DES_AUX_RECOGE_TLF        ,:A_DES_AUX_RECOGE_NOMBRE   ,:A_DES_AUX_RECOGE_DIRECC   ," _
                                                             & ":A_DES_AUX_RECOGE_REF        ,:A_DES_AUX_MOTORIZADO      ,:A_NUM_RUC_EMPRESA           ,:A_DES_RAZON_SOCIAL        ," _
                                                             & ":A_OBS_NOTA_LOCAL            ,:A_OBS_NOTA_VERIFICACION   ,:A_OBS_NOTA_RUTEO            ,:A_OBS_NOTA_MOTORIZADO     ," _
                                                             & ":A_CAD_COD_USU_RESP          ," _
                                                             & ":A_CAD_IMP_USU_RESP          ,:A_CAD_COD_FORM_PAGO       ,:A_CAD_COD_FPAG_HIJO       ," _
                                                             & ":A_CAD_COD_MONEDA_PAGO       ,:A_CAD_TIP_CAMBIO          ,:A_CAD_COD_TIP_TARJETA     ," _
                                                             & ":A_CAD_NUM_TARJETA           ,:A_CAD_FLG_TIP_CUOTA       ,:A_CAD_NUM_CUOTA           ," _
                                                             & ":A_CAD_FCH_VENCIMIENTO       ,:A_CAD_MTO_RETIRO_EFEC     ,:A_CAD_COD_DOC_PAGO        ," _
                                                             & ":A_CAD_NUM_DOC_PAGO          ,:A_NUM_DOC_NOTA_CRED       ,:A_FCH_DOC_NOTA_CRED       ," _
                                                             & ":A_CAD_COD_DOC_IDENT         ,:A_CAD_NUM_DOC_IDENT       ,:A_CAD_NOM_TITULAR         ,:A_CAD_COD_BANCO           ," _
                                                             & ":A_CAD_NUM_MOVIMIENTO        ,:A_CAD_FCH_MOVIMIENTO      ,:A_CAD_CTA_CTE_BTL         ," _
                                                             & ":A_CAD_COD_DONACION          ,:A_CAD_IMPORTE             ,:A_CAD_TIP_VENTA           ," _
                                                             & ":A_CAD_COD_PRODUCTO          ,:A_CAD_CTD_PRODUCTO        ,:A_CAD_FLG_CTD_PRODUCTO    ," _
                                                             & ":A_CAD_SUBT_NETO_PRODUCTO    ,:A_CAD2_TIP_VENTA          ,:A_CAD2_COD_PRODUCTO       ," _
                                                             & ":A_CAD2_COD_PRODUCTO_INS     ,:A_CAD2_CTD_PRODUCTO       ,:A_CAD2_FLG_CTD_PRODUCTO   ," _
                                                             & ":A_CAD2_PRC_UNIT_INSUMO      ,:A_CAD2_PCT_BASE           ,:A_CAD2_CTD_BASE           ," _
                                                             & ":A_CAD2_UND_MEDIDA           ,:A_NUM_PROFORMA_GENERADA   ,:A_NUM_PROFORMA_ORIG       ,:A_FLG_COMMIT ,:A_NOM_TITULAR, :A_NUM_DNI, :A_FLG_ENTREGA_TERCERO, :A_FLG_FECHA_PACTADA, :A_COD_MOTORIZADO_CONV, NULL,:A_CAD_NUM_AUTORIZACION, :A_CAD_FLG_PROMOCION, :A_DES_DIREC_SOCIAL, :A_CARI_ID, :A_CAD_COD_PROMOCION,  :A_CAD_COD_PROMOCION_DCTO, null, null, null, :A_NUM_ARCHIVO, :A_LATITUD, :A_LONGITUD, :A_TIPO_DOCUMENTO_PTS, :A_NUM_DOCUMENTO_PTS, :A_CAD_NUM_VALE, :A_FLG_RESERVA_CAPACIDAD, :A_COD_LOCAL_CALL_CENTER, :A_DELIVERY_TYPE, :A_HORA_SEGUNDA_PACT_ENTR, :A_DELIVERY_TIME, :A_IND_RESERVE_CAPACIDAD, :A_START_HOUR, :A_END_HOUR) ; END;"
        Else
            GoTo gstrIndRAv3Other
        End If
    End If
    
'   oradb.ExecuteSQL " BEGIN BTLPROD.PKG_PROFORMA_V3.SP_GRABA (:A_COD_CIA                  ,:A_COD_LOCAL               ,:A_COD_MAQUINA             ," _
'                                                         & ":A_COD_CLIENTE               ,:A_COD_TIP_DOC_VTA         ,:A_COD_EMPRESA             ," _
'                                                         & ":A_COD_CONVENIO              ,:A_PCT_BENEFICIARIO        ,:A_COD_MEDICO              ," _
'                                                         & ":A_IMP_TIP_CAMBIO            ,:A_COD_USUARIO             ,:A_COD_UBIGEO              ,:A_COD_TIPO_VENTA , :A_COD_LOCAL_REF             ," _
'                                                         & ":A_CAD_COD_CAMPO             ,:A_CAD_VAL_CAMPO           ,:A_CAD_COD_DIAGNOSTICO     ,:A_FCH_RECETA              ,:A_COD_LOCAL_RECETA , :A_COD_CLIENTE_DLV ,:A_FLG_TRANSF_LOCAL        ,:A_FLG_ENTREGA_LOCAL       ," _
'                                                         & ":A_FLG_RESERVA_STK           ,:A_COD_DIRECCION_CLI       ,:A_COD_MOTORIZADO          ," _
'                                                         & ":A_FLG_URGENTE               ,:A_DES_AUX_CLI_TLF         ,:A_DES_AUX_CLI_NOMBRE      ," _
'                                                         & ":A_DES_AUX_CLI_DIRECC        ,:A_DES_AUX_CLI_REF         ,:A_FCH_HORA_PACT_ENTR      ,:A_FCH_HORA_PACT_RECOG     ," _
'                                                         & ":A_DES_AUX_RECOGE_TLF        ,:A_DES_AUX_RECOGE_NOMBRE   ,:A_DES_AUX_RECOGE_DIRECC   ," _
'                                                         & ":A_DES_AUX_RECOGE_REF        ,:A_DES_AUX_MOTORIZADO      ,:A_NUM_RUC_EMPRESA           ,:A_DES_RAZON_SOCIAL        ," _
'                                                         & ":A_OBS_NOTA_LOCAL            ,:A_OBS_NOTA_VERIFICACION   ,:A_OBS_NOTA_RUTEO            ,:A_OBS_NOTA_MOTORIZADO     ," _
'                                                         & ":A_CAD_COD_USU_RESP          ," _
'                                                         & ":A_CAD_IMP_USU_RESP          ,:A_CAD_COD_FORM_PAGO       ,:A_CAD_COD_FPAG_HIJO       ," _
'                                                         & ":A_CAD_COD_MONEDA_PAGO       ,:A_CAD_TIP_CAMBIO          ,:A_CAD_COD_TIP_TARJETA     ," _
'                                                         & ":A_CAD_NUM_TARJETA           ,:A_CAD_FLG_TIP_CUOTA       ,:A_CAD_NUM_CUOTA           ," _
'                                                         & ":A_CAD_FCH_VENCIMIENTO       ,:A_CAD_MTO_RETIRO_EFEC     ,:A_CAD_COD_DOC_PAGO        ," _
'                                                         & ":A_CAD_NUM_DOC_PAGO          ,:A_NUM_DOC_NOTA_CRED       ,:A_FCH_DOC_NOTA_CRED       ," _
'                                                         & ":A_CAD_COD_DOC_IDENT         ,:A_CAD_NUM_DOC_IDENT       ,:A_CAD_NOM_TITULAR         ,:A_CAD_COD_BANCO           ," _
'                                                         & ":A_CAD_NUM_MOVIMIENTO        ,:A_CAD_FCH_MOVIMIENTO      ,:A_CAD_CTA_CTE_BTL         ," _
'                                                         & ":A_CAD_COD_DONACION          ,:A_CAD_IMPORTE             ,:A_CAD_TIP_VENTA           ," _
'                                                         & ":A_CAD_COD_PRODUCTO          ,:A_CAD_CTD_PRODUCTO        ,:A_CAD_FLG_CTD_PRODUCTO    ," _
'                                                         & ":A_CAD_SUBT_NETO_PRODUCTO    ,:A_CAD2_TIP_VENTA          ,:A_CAD2_COD_PRODUCTO       ," _
'                                                         & ":A_CAD2_COD_PRODUCTO_INS     ,:A_CAD2_CTD_PRODUCTO       ,:A_CAD2_FLG_CTD_PRODUCTO   ," _
'                                                         & ":A_CAD2_PRC_UNIT_INSUMO      ,:A_CAD2_PCT_BASE           ,:A_CAD2_CTD_BASE           ," _
'                                                         & ":A_CAD2_UND_MEDIDA           ,:A_NUM_PROFORMA_GENERADA   ,:A_NUM_PROFORMA_ORIG       ,:A_FLG_COMMIT ,:A_NOM_TITULAR, :A_NUM_DNI, :A_FLG_ENTREGA_TERCERO, :A_FLG_FECHA_PACTADA, :A_COD_MOTORIZADO_CONV, NULL,:A_CAD_NUM_AUTORIZACION, :A_CAD_FLG_PROMOCION, :A_DES_DIREC_SOCIAL, :A_CARI_ID, :A_CAD_COD_PROMOCION,  :A_CAD_COD_PROMOCION_DCTO, null, null, null, :A_NUM_ARCHIVO, :A_LATITUD, :A_LONGITUD, :A_TIPO_DOCUMENTO_PTS, :A_NUM_DOCUMENTO_PTS, :A_CAD_NUM_VALE, :A_FLG_RESERVA_CAPACIDAD, :A_COD_LOCAL_CALL_CENTER, :A_DELIVERY_TYPE, :A_HORA_SEGUNDA_PACT_ENTR, :A_DELIVERY_TIME) ; END;"
'                                                         '& ":A_CAD2_UND_MEDIDA           ,:A_NUM_PROFORMA_GENERADA   ,:A_NUM_PROFORMA_ORIG       ,:A_FLG_COMMIT ,:A_NOM_TITULAR, :A_NUM_DNI, :A_FLG_ENTREGA_TERCERO, :A_FLG_FECHA_PACTADA, :A_COD_MOTORIZADO_CONV, NULL,:A_CAD_NUM_AUTORIZACION, :A_CAD_FLG_PROMOCION, :A_DES_DIREC_SOCIAL, :A_CARI_ID, :A_CAD_COD_PROMOCION,  :A_CAD_COD_PROMOCION_DCTO, null, null, null, :A_NUM_ARCHIVO, :A_LATITUD, :A_LONGITUD, :A_TIPO_DOCUMENTO_PTS, :A_NUM_DOCUMENTO_PTS, :A_CAD_NUM_VALE) ; END;"
                                                         ' 17.12.2020 | se comenta pero volver a usar | se vuelve agregar nuevo parametro
                                                         




Set Ctd2PctBase = oradb.Parameters("A_CAD2_PCT_BASE")
Set Ctd2Base = oradb.Parameters("A_CAD2_CTD_BASE")
Set UndMedida = oradb.Parameters("A_CAD2_UND_MEDIDA")
Set Ctd2Producto = oradb.Parameters("A_CAD2_CTD_PRODUCTO")


    
    pNumProforma = Trim(oradb.Parameters("A_NUM_PROFORMA_GENERADA").Value)

        




Exit Function

Control:

    Err.Raise Err.Number, "clsVenta.Grabar", Err.Description

End Function



Public Static Property Get NumElementosFormaPago() As Byte
NumElementosFormaPago = byteNumElementosFormaPago

End Property

Public Static Property Let NumElementosFormaPago(ByVal lbyteNumElementosFormaPago As Byte)
byteNumElementosFormaPago = lbyteNumElementosFormaPago
End Property

Public Static Property Get NumElementosProd() As Byte
NumElementosProd = byteNumElementosProd
End Property

Public Static Property Let NumElementosProd(ByVal lbyteNumElementosProd As Byte)
byteNumElementosProd = lbyteNumElementosProd
End Property

Public Static Property Get NumElementosRM() As Byte
NumElementosRM = byteNumElementosRM
End Property

Public Static Property Let NumElementosRM(ByVal lbyteNumElementosRM As Byte)
byteNumElementosRM = lbyteNumElementosRM
End Property

'*************************************************************************************
' Metodos para la parte de Convenios
'*************************************************************************************

Public Property Get flgBeneficiarios() As String
    flgBeneficiarios = strflgBeneficiarios
End Property

Public Property Let flgBeneficiarios(ByVal lflgBeneficiarios As String)
    strflgBeneficiarios = lflgBeneficiarios
End Property

Public Property Get PctBeneficiario() As Double
    PctBeneficiario = dblPctBeneficiario
End Property

Public Property Let PctBeneficiario(ByVal ldblPctBeneficiario As Double)
    dblPctBeneficiario = ldblPctBeneficiario
End Property

Public Property Get ImpPctBeneficiario() As Double
    ImpPctBeneficiario = ImpdblPctBeneficiario
End Property

Public Property Let ImpPctBeneficiario(ByVal ldblPctBeneficiario As Double)
    ImpdblPctBeneficiario = ldblPctBeneficiario
End Property
Public Property Get flgPctBeneficiario() As Integer
    flgPctBeneficiario = flgIntPctBeneficiario
End Property

Public Property Let flgPctBeneficiario(ByVal ldblPctBeneficiario As Integer)
    flgIntPctBeneficiario = ldblPctBeneficiario
End Property
Public Property Get FlgRepartidor() As String
    FlgRepartidor = strFlgRepartidor
End Property

Public Property Let FlgRepartidor(ByVal lstrFlgRepartidor As String)
    strFlgRepartidor = lstrFlgRepartidor
End Property


Public Property Get FlgMedico() As String
    FlgMedico = strFlgMedico
End Property

Public Property Let FlgMedico(ByVal lstrFlgMedico As String)
    strFlgMedico = lstrFlgMedico
End Property

Public Property Get FlgPolitica() As String
    FlgPolitica = strFlgPolitica
End Property

Public Property Let FlgPolitica(ByVal lstrFlgPolitica As String)
    strFlgPolitica = lstrFlgPolitica
End Property

'''Public Property Get CodConvenio() As String
'''    CodConvenio = strCodConvenio
'''End Property
'''
'''Public Property Let CodConvenio(ByVal lstrCodConvenio As String)
'''    strCodConvenio = lstrCodConvenio
'''End Property



'''''Public Property Get CodBeneficiario() As String
'''''    CodBeneficiario = strCodBeneficiario
'''''End Property
'''''
'''''Public Property Let CodBeneficiario(ByVal lstrCodBeneficiario As String)
'''''    strCodBeneficiario = lstrCodBeneficiario
'''''End Property

'''Public Property Get NomConvenio() As String
'''    NomConvenio = strNomConvenio
'''End Property
'''
'''Public Property Let NomConvenio(ByVal lstrNomConvenio As String)
'''    strNomConvenio = lstrNomConvenio
'''End Property

'''Public Property Get NomBeneficiario() As String
'''    NomBeneficiario = strNomBeneficiario
'''End Property
'''
'''Public Property Let NomBeneficiario(ByVal lstrNomBeneficiario As String)
'''    strNomBeneficiario = lstrNomBeneficiario
'''End Property


Public Property Get CodEmpresa() As String
    CodEmpresa = strCodEmpresa
End Property

Public Property Let CodEmpresa(ByVal lstrCodEmpresa As String)
    strCodEmpresa = lstrCodEmpresa
End Property

Public Property Get CodMedico() As String
    CodMedico = strCodMedico
End Property

Public Property Let CodMedico(ByVal lstrCodMedico As String)
    strCodMedico = lstrCodMedico
End Property




Public Property Get DatosAdicional() As XArrayDB
    Set DatosAdicional = xDatoAdicional
End Property

Public Property Set DatosAdicional(ByVal lstrDatosAdicionales As XArrayDB)
    Set xDatoAdicional = lstrDatosAdicionales
End Property

Public Property Get Diagnostico() As XArrayDB
    Set Diagnostico = xDiagnostico
End Property

Public Property Set Diagnostico(ByVal lstrDiagnostico As XArrayDB)
    Set xDiagnostico = lstrDiagnostico
End Property

Public Function AgregaLineaDatoAdicional(ByVal Codigo As String, ByVal Descripcion As String, ByVal Tipo As String, ByVal max As Integer, ByVal min As Integer, ByVal Valor As Variant) As XArrayDB
Dim ultimo As Integer
Dim aux As Integer
    

    
    If xDatoAdicional.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    aux = xDatoAdicional.Count(1)
    While i < aux
        If xDatoAdicional(i, 0) = Codigo Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xDatoAdicional.Count(1)
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xDatoAdicional.AppendRows
    End If
j:
    If xDatoAdicional.Count(1) = 0 Then ultimo = 0: xDatoAdicional.AppendRows

    xDatoAdicional(ultimo, 0) = Codigo
    xDatoAdicional(ultimo, 1) = Descripcion
    xDatoAdicional(ultimo, 2) = Tipo
    xDatoAdicional(ultimo, 3) = max
    xDatoAdicional(ultimo, 4) = min
    xDatoAdicional(ultimo, 5) = Valor
    
    
  Set DatosAdicional = xDatoAdicional
        
End Function

Public Function AgregaDiagnostico(ByVal Codigo As String, ByVal Descripcion As String) As XArrayDB
Dim ultimo As Integer
Dim aux As Integer



    If xDiagnostico.Count(1) < 0 Then Exit Function

    Dim i As Integer
    Dim encontro As Boolean
    aux = xDiagnostico.Count(1)
    While i < aux
        If xDiagnostico(i, 0) = Codigo Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xDiagnostico.Count(1)
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xDiagnostico.AppendRows
    End If
j:
    If xDiagnostico.Count(1) = 0 Then ultimo = 0: xDiagnostico.AppendRows

    xDiagnostico(ultimo, 0) = Codigo
    xDiagnostico(ultimo, 1) = Descripcion

End Function


'*************************************************************************************
' Metodos para la parte de Cliente
'*************************************************************************************


''''''Public Property Get CodCliente() As String
''''''    CodCliente = strCodCliente
''''''End Property
''''''
''''''Public Property Let CodCliente(ByVal lstrCodCliente As String)
''''''    strCodCliente = lstrCodCliente
''''''End Property

'*************************************************************************************
' Metodos para la parte de Documento
'*************************************************************************************


'''Public Property Get CodTipDocVta() As String
'''    CodTipDocVta = strCodTipDocVta
'''End Property
'''
'''Public Property Let CodTipDocVta(ByVal lstrCodTipDocVta As String)
'''    strCodTipDocVta = lstrCodTipDocVta
'''End Property

'--- Documento referncia para cuando es nota de credito ---'
Public Property Let CodDocRef(ByVal vstrCodDocRef As String)
    strCodDocRef = vstrCodDocRef
End Property

Public Property Get CodDocRef() As String
    CodDocRef = strCodDocRef
End Property

Public Property Let NumDocRef(ByVal vstrNumDocRef As String)
    strNumDocRef = vstrNumDocRef
End Property

Public Property Get NumDocRef() As String
    NumDocRef = strNumDocRef
End Property

Public Property Let BtlRef(ByVal vstrBtlRef As String)
    strBtlRef = vstrBtlRef
End Property

Public Property Get BtlRef() As String
    BtlRef = strBtlRef
End Property

Public Property Let FlgNcrDev(ByVal vstrFlgNcrDev As String)
    strFlgNcrDev = vstrFlgNcrDev
End Property

Public Property Get FlgNcrDev() As String
    FlgNcrDev = strFlgNcrDev
End Property

Public Property Let CodConcepto(ByVal vstrCodConcepto As String)
    strConcepto = vstrCodConcepto
End Property

Public Property Get CodConcepto() As String
    CodConcepto = strConcepto
End Property

Public Property Let CodSubConcepto(ByVal vstrCodSubConcepto As String)
    strSubConcepto = vstrCodSubConcepto
End Property

Public Property Get CodSubConcepto() As String
    CodSubConcepto = strSubConcepto
End Property

'-----------------------------------------------------------'

'*************************************************************************************
' Metodos para la parte de Delivery
'*************************************************************************************

Public Static Property Get FlgTransfLocal() As String
FlgTransfLocal = strFlgTransfLocal
End Property

Public Static Property Let FlgTransfLocal(ByVal lstrFlgTransfLocal As String)
    strFlgTransfLocal = lstrFlgTransfLocal
End Property

Public Static Property Get FlgEntregaLocal() As String
    FlgEntregaLocal = strFlgEntregaLocal
End Property

Public Static Property Let FlgEntregaLocal(ByVal lstrFlgEntregaLocal As String)
strFlgEntregaLocal = lstrFlgEntregaLocal
End Property

'Public Static Property Get FlgPedUrgente() As String
'    FlgPedUrgente = strFlgPedUrgente
'End Property
'
'Public Static Property Let FlgPedUrgente(ByVal lstrFlgPedUrgente As String)
'    strFlgPedUrgente = lstrFlgPedUrgente
'End Property


Public Static Property Get FlgReservaStk() As String
FlgReservaStk = strFlgReservaStk
End Property

Public Static Property Let FlgReservaStk(ByVal lstrFlgReservaStk As String)
strFlgReservaStk = lstrFlgReservaStk
End Property


Public Static Property Get CodDireccionCli() As String
CodDireccionCli = strCodDireccionCli
End Property

Public Static Property Let CodDireccionCli(ByVal lstrCodDireccionCli As String)
strCodDireccionCli = lstrCodDireccionCli
End Property

Public Static Property Get CodMotorizado() As String
CodMotorizado = strCodMotorizado
End Property

Public Static Property Let CodMotorizado(ByVal lstrCodMotorizado As String)
strCodMotorizado = lstrCodMotorizado
End Property

Public Static Property Get FlgUrgente() As String
FlgUrgente = strFlgUrgente
End Property

Public Static Property Let FlgUrgente(ByVal lstrFlgUrgente As String)
    strFlgUrgente = lstrFlgUrgente
End Property

Public Static Property Get FlgPactado() As String
    FlgPactado = strFlgPactado
End Property

Public Static Property Let FlgPactado(ByVal lstrFlgPactado As String)
    strFlgPactado = lstrFlgPactado
End Property

Public Static Property Get DesAuxCliTlf() As String
    DesAuxCliTlf = strDesAuxCliTlf
End Property

Public Static Property Let DesAuxCliTlf(ByVal lstrDesAuxCliTlf As String)
    strDesAuxCliTlf = lstrDesAuxCliTlf
End Property

Public Static Property Get DesAuxCliNombre() As String
DesAuxCliNombre = strDesAuxCliNombre
End Property

Public Static Property Let DesAuxCliNombre(ByVal lstrDesAuxCliNombre As String)
strDesAuxCliNombre = lstrDesAuxCliNombre
End Property

Public Static Property Get EntregaTercero() As String
EntregaTercero = flgEntregaTercero
End Property

Public Static Property Let EntregaTercero(ByVal lstrflgEntregaTercero As String)
flgEntregaTercero = lstrflgEntregaTercero
End Property

Public Static Property Get DesAuxCliDirecc() As String
    DesAuxCliDirecc = strDesAuxCliDirecc
End Property

Public Static Property Let DesAuxCliDirecc(ByVal lstrDesAuxCliDirecc As String)
    strDesAuxCliDirecc = lstrDesAuxCliDirecc
End Property



Public Static Property Get FchHoraPactEntr() As String
FchHoraPactEntr = datFchHoraPactEntr
End Property

Public Static Property Let FchHoraPactEntr(ByVal ldatFchHoraPactEntr As String)
datFchHoraPactEntr = ldatFchHoraPactEntr
End Property

Public Static Property Get HoraPactEntr() As String
HoraPactEntr = strHoraPactEntr
End Property

Public Static Property Let HoraPactEntr(ByVal lstrHoraPactEntrada As String)
strHoraPactEntr = lstrHoraPactEntrada
End Property




Public Static Property Get FchHoraPactRecog() As String
    FchHoraPactRecog = datFchHoraPactRecog
End Property

Public Static Property Let FchHoraPactRecog(ByVal ldatFchHoraPactRecog As String)
    datFchHoraPactRecog = ldatFchHoraPactRecog
End Property

Public Static Property Get HoraPactRecog() As String
    HoraPactRecog = strHoraPactRecogo
End Property

Public Static Property Let HoraPactRecog(ByVal lstrHoraPactRecogo As String)
    strHoraPactRecogo = lstrHoraPactRecogo
End Property



Public Static Property Get DesAuxRecogeTlf() As String
DesAuxRecogeTlf = strDesAuxRecogeTlf
End Property

Public Static Property Let DesAuxRecogeTlf(ByVal lstrDesAuxRecogeTlf As String)
strDesAuxRecogeTlf = lstrDesAuxRecogeTlf
End Property

Public Static Property Get DesAuxRecogeNombre() As String
DesAuxRecogeNombre = strDesAuxRecogeNombre
End Property

Public Static Property Let DesAuxRecogeNombre(ByVal lstrDesAuxRecogeNombre As String)
strDesAuxRecogeNombre = lstrDesAuxRecogeNombre
End Property


Public Static Property Get DesAuxRecogeDirecc() As String
DesAuxRecogeDirecc = strDesAuxRecogeDirecc
End Property

Public Static Property Let DesAuxRecogeDirecc(ByVal lstrDesAuxRecogeDirecc As String)
strDesAuxRecogeDirecc = lstrDesAuxRecogeDirecc
End Property


Public Static Property Get DesAuxRecogeRef() As String
DesAuxRecogeRef = strDesAuxRecogeRef
End Property

Public Static Property Let DesAuxRecogeRef(ByVal lstrDesAuxRecogeRef As String)
strDesAuxRecogeRef = lstrDesAuxRecogeRef
End Property

Public Static Property Get DesAuxMotorizado() As String
DesAuxMotorizado = strDesAuxMotorizado
End Property

Public Static Property Let DesAuxMotorizado(ByVal lstrDesAuxMotorizado As String)
strDesAuxMotorizado = lstrDesAuxMotorizado
End Property

Public Static Property Get ObsNotaVerificacion() As String
ObsNotaVerificacion = strObsNotaVerificacion
End Property

Public Static Property Let ObsNotaVerificacion(ByVal lstrObsNotaVerificacion As String)
strObsNotaVerificacion = lstrObsNotaVerificacion
End Property

Public Static Property Get ObsNotaRuteo() As String
ObsNotaRuteo = strObsNotaRuteo
End Property

Public Static Property Let ObsNotaRuteo(ByVal lstrObsNotaRuteo As String)
strObsNotaRuteo = lstrObsNotaRuteo
End Property

Public Static Property Get ObsNotaLocal() As String
ObsNotaLocal = strObsNotaLocal
End Property

Public Static Property Let ObsNotaLocal(ByVal lstrObsNotaLocal As String)
strObsNotaLocal = lstrObsNotaLocal
End Property


'=============================================================================================================================
'=============================================================================================================================
'Indice |Fecha          | Autor             | Cambio realizado
'-----------------------------------------------------------------------------------------------------------------------------
'1      | 12/11/2009     | Arturo Escate     | Se agrego el campo de observación  al documento
'-----------------------------------------------------------------------------------------------------------------------------
'=============================================================================================================================

Public Function GrabarDoc(ByVal oradb As OraDatabase, _
                          ByRef oTipDoc As OraParamArray, _
                          ByRef oNumDoc As OraParamArray, _
                          ByRef oTipDocCo As OraParamArray, _
                          ByRef oNumDocCo As OraParamArray, _
                          ByRef RetPromoMensaje As String, _
                          Optional flg_NC As Boolean, _
                          Optional strcodConv As String, _
                          Optional dblPctBeneficarioAux As Double, _
                          Optional dblPctBeneficiariOrig, _
                          Optional FlgAnulados As Boolean) As Boolean
                          
                          
'=============================================================================================================================
'=============================================================================================================================
'Validaciones de Visual
'=============================================================================================================================
'If strcodConv <> "" Then
'    'If CodigoConvenio = "" And flg_NC = False And ObjVenta.ptmModalidad <> Cobro_Responsabilidad Then MsgBox "Ingrese forma de Pago", vbCritical, App.ProductName: GrabarDoc = False: frm_VTA_FormaPago.Show: Exit Function
'Else
'    If FormaPago.Count(1) <= 1 And CodigoConvenio = "" And flg_NC = False And objVenta.ptmModalidad <> Cobro_Responsabilidad Then MsgBox "Ingrese forma de Pago", vbCritical, App.ProductName: GrabarDoc = False: frm_VTA_FormaPago.Show: Exit Function
'End If
'If Producto.Count(1) <= 0 Then MsgBox "Debe de seleccionar algun producto", vbCritical, App.ProductName: GrabarDoc = False: Exit Function
''If Not Me.CodigoTipoVenta = Venta_Convenio And Me.TipoCliente = "" And flg_NC = False Then MsgBox "Tiene que cambiar el tipo de documento de venta", vbCritical, App.ProductName: GrabarDoc = False: Exit Function
''=====================================================fin de validaciones=====================================================
''Autor : Juan Arturo Escate Espichan
''Fecha : 18/08/2008
Dim SecuenciaGrabacion As String
SecuenciaGrabacion = objUsuario.NombrePC & Format(Date, "yyyymmyy") & Format(time, "hhmmss")

FlgAnulados = False
If Me.debeRegularizar = True Then
    MsgBox "Tiene documentos de contingencia que aun no se han regularizado", vbCritical, App.ProductName
    Exit Function
End If

If Servicios.UpperBound(1) > 0 And Not objVenta.CodigoTipoVenta = Servicio Then
    MsgBox "Debe de cambiar la modalidad de venta a Servicios", vbCritical, App.ProductName
    Exit Function
End If

''Validacion de la venta mayorista, si el cliente es atendido por la BTLCERO o por el local
Dim strMensajeMayo As String
If objVenta.CodigoTipoVenta = Venta_Mayorista Then
    strMensajeMayo = Me.ValidaVentaMayorista(objVenta.CodigoCliente, objUsuario.CodigoLocal, Totales(0))
    If Not strMensajeMayo = "" Then
        MsgBox strMensajeMayo, vbCritical, App.ProductName
        Exit Function
    End If
End If
''Con esta funcion valido que la venta por lotes tengan todos los datos
If ValidaVentaLotes = False Then
    Exit Function
End If

Dim i As Integer
Dim strCajaAbierta As String

'''''''''''''''''''On Error GoTo Control
'If PctBeneficiario > 0 Then
'            If flgIntPctBeneficiario = "0" Then 'POR PORCENTAJE
'            Else 'POR IMPORTE
'                If Me.ImpPctBeneficiario > Totales(0) Then
'                    MsgBox "El Importe de Co-Pago No pude ser mayor al monto", vbCritical, App.ProductName
'
'                    Exit Function
'                End If
'
'            End If
'
'End If

byteNumElementosProd = IIf(Producto.UpperBound(1) < 0, 0, Producto.UpperBound(1)) + 1
byteNumElementosFormaPago = IIf(FormaPago.UpperBound(1) < 0, 0, FormaPago.UpperBound(1)) + 1
byteNumElementosCobro = IIf(CobroResponsabilidad.UpperBound(1) < 0, 0, CobroResponsabilidad.UpperBound(1)) + 1
byteNumElementosRM = IIf(RecetarioMagistral.UpperBound(1) < 0, 0, RecetarioMagistral.UpperBound(1)) + 1
byteNumElementosDatoAdic = IIf(DatosAdicional.UpperBound(1) < 0, 0, DatosAdicional.UpperBound(1)) + 1
byteNumElementosServ = IIf(Servicios.UpperBound(1) < 0, 0, Servicios.UpperBound(1)) + 1
byteNumElementosEspVal = IIf(EspeciesValoradas.UpperBound(1) < 0, 0, EspeciesValoradas.UpperBound(1)) + 1



    For i = oradb.Parameters.Count - 1 To 0 Step -1

        oradb.Parameters.Remove i

    Next







oradb.Parameters.Add "A_COD_CIA", objUsuario.CodigoEmpresa, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_LOCAL", objUsuario.CodigoLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_MAQUINA", objUsuario.NombrePC, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_TIPO_VENTA", objUsuario.CodTipoVenta, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_CLIENTE", Trim(CodigoCliente), ORAPARM_INPUT, ORATYPE_VARCHAR2

oradb.Parameters.Add "A_COD_TIP_DOC_VTA", CodigoDocumentoVenta, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_EMPRESA", CodEmpresa, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_CONVENIO", Trim(CodigoConvenio), ORAPARM_INPUT, ORATYPE_VARCHAR2
'oradb.Parameters.Add "A_PCT_BENEFICIARIO", PctBeneficiario, ORAPARM_INPUT, ORATYPE_VARCHAR2

oradb.Parameters.Add "A_COD_MEDICO", CodMedico, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_REPARTIDOR", CodMotorizado, ORAPARM_INPUT, ORATYPE_VARCHAR2

oradb.Parameters.Add "A_IMP_TIP_CAMBIO", objUsuario.TipoCambio, ORAPARM_INPUT, ORATYPE_VARCHAR2


oradb.Parameters.Add "A_NUM_PEDIDO_PADRE", NumPedidoPadre, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_CONCEPTO", CodConcepto, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_SUBCONCEPTO_FACT", CodSubConcepto, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_ORIGEN_DOCUMENTO", "", ORAPARM_INPUT, ORATYPE_VARCHAR2

oradb.Parameters.Add "A_COD_USUARIO", objUsuario.Codigo, ORAPARM_INPUT, ORATYPE_VARCHAR2

oradb.Parameters.Add "A_COD_LOCAL_REFER", BtlRef, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_TIPO_DOC_REFER", CodDocRef, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_NUM_DOCUMENTO_REFER", NumDocRef, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_FLG_NCR_DEVOLUCION", FlgNcrDev, ORAPARM_INPUT, ORATYPE_VARCHAR2

oradb.Parameters.AddTable "A_CAD_COD_CAMPO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDatoAdic, 200
oradb.Parameters.AddTable "A_CAD_VAL_CAMPO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDatoAdic, 200

oradb.Parameters.Add "A_FLG_ORIGEN_DELIV", "", ORAPARM_INPUT, ORATYPE_VARCHAR2

oradb.Parameters.Add "A_FLG_TRANSF_LOCAL", "", ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_FLG_ENTREGA_LOCAL", "", ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_FLG_RESERVA_STK", "", ORAPARM_INPUT, ORATYPE_VARCHAR2


If TipoCliente = "1" Then
    oradb.Parameters.Add "A_NUM_RUC_EMPRESA", Out_NumeroId, ORAPARM_INPUT, ORATYPE_VARCHAR2
    oradb.Parameters.Add "A_DES_RAZON_SOCIAL", Out_NombreCliente, ORAPARM_INPUT, ORATYPE_VARCHAR2
    'Agregado por DJara `12/09/2008
    oradb.Parameters.Add "A_DES_AUX_CLI_DIRECC", Out_Direccion, ORAPARM_INPUT, ORATYPE_VARCHAR2
    '*******************************
Else
    oradb.Parameters.Add "A_NUM_RUC_EMPRESA", objVenta.NumeroDocumentoID, ORAPARM_INPUT, ORATYPE_VARCHAR2
    oradb.Parameters.Add "A_DES_RAZON_SOCIAL", objVenta.NombreCliente, ORAPARM_INPUT, ORATYPE_VARCHAR2
    'Agregado por DJara `12/09/2008
    If objVenta.CodigoTipoVenta = Venta_Convenio Then
        oradb.Parameters.Add "A_DES_AUX_CLI_DIRECC", Out_Direccion, ORAPARM_INPUT, ORATYPE_VARCHAR2
    Else
        oradb.Parameters.Add "A_DES_AUX_CLI_DIRECC", objVenta.DireccionCliente, ORAPARM_INPUT, ORATYPE_VARCHAR2
    End If
    '*******************************
End If

oradb.Parameters.Add "A_COD_DIRECCION_CLI", CodDireccionCli, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_FLG_URGENTE", "", ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_DES_AUX_CLI_TLF", DesAuxCliTlf, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_DES_AUX_CLI_NOMBRE", DesAuxCliNombre, ORAPARM_INPUT, ORATYPE_VARCHAR2
'Comentado por DJara 12/09/2008
'oradb.Parameters.Add "A_DES_AUX_CLI_DIRECC", DesAuxCliDirecc, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_UBIGEO", Ubigeo, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_MOTIVO_COBRO", CodMotivoCobro, ORAPARM_INPUT, ORATYPE_VARCHAR2

oradb.Parameters.AddTable "A_CAD_COD_USU_RESP", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosCobro, 200
oradb.Parameters.AddTable "A_CAD_IMP_USU_RESP", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosCobro, 200
oradb.Parameters.AddTable "A_CAD_COD_FORM_PAGO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_COD_FPAG_HIJO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_COD_MONEDA_PAGO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_TIP_CAMBIO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_COD_TIP_TARJETA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_NUM_TARJETA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_NUM_AUTORIZACION", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_FLG_TIP_CUOTA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_NUM_CUOTA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_FCH_VENCIMIENTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_MTO_RETIRO_EFEC", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_COD_DOC_PAGO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_NUM_DOC_PAGO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_NUM_DOC_NOTA_CRED", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_FCH_DOC_NOTA_CRED", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_COD_DOC_IDENT", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_NUM_DOC_IDENT", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_COD_BANCO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_NUM_MOVIMIENTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_FCH_MOVIMIENTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_CTA_CTE_BTL", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_COD_DONACION", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_IMPORTE", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200
oradb.Parameters.AddTable "A_CAD_TIP_VENTA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200
oradb.Parameters.AddTable "A_CAD_COD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200
oradb.Parameters.AddTable "A_CAD_CTD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200
oradb.Parameters.AddTable "A_CAD_FLG_CTD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200
oradb.Parameters.AddTable "A_CAD_SUBT_NETO_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200
oradb.Parameters.AddTable "A_CAD_FLG_PROMOCION", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200
oradb.Parameters.AddTable "A_CAD_NRO_LOTE", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200
oradb.Parameters.AddTable "A_CAD_FCH_VMTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200
''Para Recetario Magistral''
'Graba Documento'
oradb.Parameters.AddTable "A_CAD2_TIP_VENTA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200
oradb.Parameters.AddTable "A_CAD2_COD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200
oradb.Parameters.AddTable "A_CAD2_COD_PRODUCTO_INS", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200
oradb.Parameters.AddTable "A_CAD2_CTD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200
oradb.Parameters.AddTable "A_CAD2_FLG_CTD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200
oradb.Parameters.AddTable "A_CAD2_PRC_UNIT_INSUMO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200

oradb.Parameters.AddTable "A_CAD2_PCT_BASE", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200
oradb.Parameters.AddTable "A_CAD2_CTD_BASE", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200
oradb.Parameters.AddTable "A_CAD2_UND_MEDIDA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosRM, 200

''Para Servicios''
oradb.Parameters.AddTable "A_CAD3_COD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosServ, 200
oradb.Parameters.AddTable "A_CAD3_COD_TIP_SERVICIO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosServ, 200
oradb.Parameters.AddTable "A_CAD3_COD_SERVICIO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosServ, 200
oradb.Parameters.AddTable "A_CAD3_IMP_NETO_SERVICIO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosServ, 200
'oradb.Parameters.AddTable "A_CAD3_NUM_IDENT_SERV", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosServ, 200
oradb.Parameters.AddTable "A_CAD3_NUM_VOUCHER", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosServ, 200
oradb.Parameters.AddTable "A_CAD3_NUM_RECIBO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosServ, 200
oradb.Parameters.AddTable "A_CAD3_NUM_SUMINISTRO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosServ, 200
oradb.Parameters.AddTable "A_CAD3_FLG_TIPO_PAGO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosServ, 200

''Para especies valoradas
oradb.Parameters.AddTable "A_ARR_COD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosEspVal, 200
oradb.Parameters.AddTable "A_ARR_NUM_ESPECIE_VAL", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosEspVal, 200
oradb.Parameters.AddTable "A_ARR_NUM_PLACA", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosEspVal, 200


oradb.Parameters.AddTable "A_TIPDOC_GENERAD", ORAPARM_OUTPUT, ORATYPE_VARCHAR2, 100, 200
oradb.Parameters.AddTable "A_NUMDOC_GENERAD", ORAPARM_OUTPUT, ORATYPE_VARCHAR2, 100, 200
oradb.Parameters.AddTable "A_TIPDOC_GENERAD_CO", ORAPARM_OUTPUT, ORATYPE_VARCHAR2, 100, 200
oradb.Parameters.AddTable "A_NUMDOC_GENERAD_CO", ORAPARM_OUTPUT, ORATYPE_VARCHAR2, 100, 200

oradb.Parameters.Add "A_CAD_COD_DIAGNOSTICO", Me.DetalleReceta, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_FCH_RECETA", Me.FechaReceta, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_LOCAL_RECETA", Me.LocalReceta, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_NUM_RECETA", "", ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_DEV_EFECTIVO_NC", DevEfectivoNC, ORAPARM_INPUT, ORATYPE_VARCHAR2

oradb.Parameters.Add "A_PCT_BENEFICIARIO_ORIG", PctBeneficiario, ORAPARM_INPUT, ORATYPE_VARCHAR2

oradb.Parameters.AddTable "A_CAD_PROMOCION", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200
oradb.Parameters.AddTable "A_CAD_PROMOCION_DCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200

oradb.Parameters.Add "A_DES_OBSERVACION", Observacion, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_MOTIVO_DEV", strMotivoNotaCredito, ORAPARM_INPUT, ORATYPE_VARCHAR2


oradb.Parameters.Add "A_EXISTE_DNI_RENIEC", objVenta.vExisteDNI_RENIEC, ORAPARM_INPUT, ORATYPE_VARCHAR2

If (CodigoDocumentoVenta = TipoDocTKB Or CodigoDocumentoVenta = TipoDocTKF) And SerieTKT = "" Then
    SerieTKT = objMaquina.Serie_Ticketera(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, TipoDocTKB, objUsuario.NombrePC)
    
    'Modificado para ticketera en Red o Local 23/08/2012 MLEVANO
    If Not SerieTKT = "" Then
        Call objDocumento.LocalORed(objUsuario.CodigoEmpresa, objUsuario.NombrePC, TipoDocTKB)
    Else
        'parche jrazuri 18/07/2008
        If SerieTKT = "" Then
                Err.Raise -100, "", "No se tiene número de Ticket para ticketera"
        End If
    End If
End If


oradb.Parameters.Add "A_COD_SERIE_REL", SerieTKT, ORAPARM_INPUT, ORATYPE_VARCHAR2

''Arturo Escate 25/05/2015 para puntos
oradb.Parameters.AddTable "A_CAD_PUNTOS", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosProd, 200
    Dim VtaTotal  As Double
    Dim vALpUNTOALL As Double
    For i = 0 To Producto.UpperBound(1)
        VtaTotal = VtaTotal + val(Producto(i, 4))
        vALpUNTOALL = vALpUNTOALL + val(Producto(i, 28))
    Next

oradb.Parameters.Add "A_VENTA_TOTAL", VtaTotal, ORAPARM_INPUT, ORATYPE_VARCHAR2

oradb.Parameters.Add "A_PUNTOS_REDIME", MontoRedime, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_PUNTOS_INICIAL", PuntosTarjetaMonedero, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_PUNTOS_ACUMULADO", vALpUNTOALL, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_NUM_TARJ_PUNTOS", NumeroTarjetaMonedero, ORAPARM_INPUT, ORATYPE_VARCHAR2

''''''''''''''''''''''''''''''''''''''''''''''''''''''

oradb.Parameters.Add "RET_PROMO_MENSAJE", "", ORAPARM_OUTPUT

Dim CadPuntos As OracleInProcServer.OraParamArray
Dim CadPuntosRedime As OracleInProcServer.OraParamArray

Dim CodCampo As OracleInProcServer.OraParamArray
Dim ValCampo As OracleInProcServer.OraParamArray
Dim CodUsuResp As OracleInProcServer.OraParamArray
Dim ImpUsuResp As OracleInProcServer.OraParamArray
Dim CodFormPago As OracleInProcServer.OraParamArray
Dim CodFPagHijo As OracleInProcServer.OraParamArray
Dim CodMonedaPago As OracleInProcServer.OraParamArray
Dim TipCambio As OracleInProcServer.OraParamArray
Dim CodTipTarjeta As OracleInProcServer.OraParamArray
Dim NumTarjeta As OracleInProcServer.OraParamArray
Dim NumAutorizacion As OracleInProcServer.OraParamArray 'Agregado Arturo Escate
Dim FlgTipCuota As OracleInProcServer.OraParamArray
Dim NumCuota As OracleInProcServer.OraParamArray
Dim FchVencimiento As OracleInProcServer.OraParamArray
Dim MtoRetiroEfec As OracleInProcServer.OraParamArray
Dim CodDocPago As OracleInProcServer.OraParamArray
Dim NumDocPago As OracleInProcServer.OraParamArray
Dim NumDocNotaCred As OracleInProcServer.OraParamArray
Dim FchDocNotaCred As OracleInProcServer.OraParamArray
Dim CodDocIdent As OracleInProcServer.OraParamArray
Dim NumDocIdent As OracleInProcServer.OraParamArray
Dim CodBanco As OracleInProcServer.OraParamArray
Dim NumMovimiento As OracleInProcServer.OraParamArray
Dim FchMovimiento As OracleInProcServer.OraParamArray
Dim CtaCteBtl As OracleInProcServer.OraParamArray
Dim CodDonacion As OracleInProcServer.OraParamArray
Dim Importe As OracleInProcServer.OraParamArray
Dim TipVenta As OracleInProcServer.OraParamArray
Dim codProducto As OracleInProcServer.OraParamArray
Dim CtdProducto As OracleInProcServer.OraParamArray
Dim FlgCtdProducto As OracleInProcServer.OraParamArray
Dim SubtNetoProducto As OracleInProcServer.OraParamArray
Dim FlgPromocion As OracleInProcServer.OraParamArray
Dim NroLote As OracleInProcServer.OraParamArray
Dim FchVmto As OracleInProcServer.OraParamArray

Dim CodigoProductoEV As OracleInProcServer.OraParamArray
Dim NumeroEV As OracleInProcServer.OraParamArray
Dim PlacaEV As OracleInProcServer.OraParamArray

Dim Tip2Venta As OracleInProcServer.OraParamArray
Dim Cod2Producto As OracleInProcServer.OraParamArray
Dim Cod2ProductoIns As OracleInProcServer.OraParamArray
Dim Ctd2Producto As OracleInProcServer.OraParamArray
Dim Flg2CtdProducto As OracleInProcServer.OraParamArray
Dim Prc2UnitInsumo As OracleInProcServer.OraParamArray
Dim Ctd2PctBase As OracleInProcServer.OraParamArray
Dim Ctd2Base As OracleInProcServer.OraParamArray
Dim UndMedida As OracleInProcServer.OraParamArray

Dim CodProductoServicio As OracleInProcServer.OraParamArray
Dim CodTipServicio As OracleInProcServer.OraParamArray
Dim CodServicio As OracleInProcServer.OraParamArray
Dim ImpNetoServicio As OracleInProcServer.OraParamArray
Dim NumVoucher As OracleInProcServer.OraParamArray
Dim NumRecibo As OracleInProcServer.OraParamArray
Dim NumSuministro As OracleInProcServer.OraParamArray
Dim FlgTipoPago As OracleInProcServer.OraParamArray

'Dim SubtNetoProductoPP As OracleInProcServer.OraParamArray

Dim dblDet As Double
Dim dblNewPctBenef As Double
Dim CadenaPromocion As OracleInProcServer.OraParamArray
Dim CadenaPromocionDcto As OracleInProcServer.OraParamArray
dblDet = 0: dblNewPctBenef = 0
''Dim oTipDoc As OracleInProcServer.OraParamArray
''Dim oNumDoc As OracleInProcServer.OraParamArray
''Dim oTipDocCo As OracleInProcServer.OraParamArray
''Dim oNumDocCo As OracleInProcServer.OraParamArray


Set CodCampo = oradb.Parameters("A_CAD_COD_CAMPO")
Set ValCampo = oradb.Parameters("A_CAD_VAL_CAMPO")
Set CodUsuResp = oradb.Parameters("A_CAD_COD_USU_RESP")
Set ImpUsuResp = oradb.Parameters("A_CAD_IMP_USU_RESP")
Set CodFormPago = oradb.Parameters("A_CAD_COD_FORM_PAGO")
Set CodFPagHijo = oradb.Parameters("A_CAD_COD_FPAG_HIJO")
Set CodMonedaPago = oradb.Parameters("A_CAD_COD_MONEDA_PAGO")
Set TipCambio = oradb.Parameters("A_CAD_TIP_CAMBIO")
Set CodTipTarjeta = oradb.Parameters("A_CAD_COD_TIP_TARJETA")
Set NumTarjeta = oradb.Parameters("A_CAD_NUM_TARJETA")
Set NumAutorizacion = oradb.Parameters("A_CAD_NUM_AUTORIZACION")  ''agregado arturo escate
Set FlgTipCuota = oradb.Parameters("A_CAD_FLG_TIP_CUOTA")
Set NumCuota = oradb.Parameters("A_CAD_NUM_CUOTA")
Set FchVencimiento = oradb.Parameters("A_CAD_FCH_VENCIMIENTO")
Set MtoRetiroEfec = oradb.Parameters("A_CAD_MTO_RETIRO_EFEC")
Set CodDocPago = oradb.Parameters("A_CAD_COD_DOC_PAGO")
Set NumDocPago = oradb.Parameters("A_CAD_NUM_DOC_PAGO")
Set NumDocNotaCred = oradb.Parameters("A_NUM_DOC_NOTA_CRED")
Set FchDocNotaCred = oradb.Parameters("A_FCH_DOC_NOTA_CRED")
Set CodDocIdent = oradb.Parameters("A_CAD_COD_DOC_IDENT")
Set NumDocIdent = oradb.Parameters("A_CAD_NUM_DOC_IDENT")
Set CodBanco = oradb.Parameters("A_CAD_COD_BANCO")
Set NumMovimiento = oradb.Parameters("A_CAD_NUM_MOVIMIENTO")
Set FchMovimiento = oradb.Parameters("A_CAD_FCH_MOVIMIENTO")
Set CtaCteBtl = oradb.Parameters("A_CAD_CTA_CTE_BTL")
Set CodDonacion = oradb.Parameters("A_CAD_COD_DONACION")
Set Importe = oradb.Parameters("A_CAD_IMPORTE")
Set TipVenta = oradb.Parameters("A_CAD_TIP_VENTA")
Set codProducto = oradb.Parameters("A_CAD_COD_PRODUCTO")
Set CtdProducto = oradb.Parameters("A_CAD_CTD_PRODUCTO")
Set FlgCtdProducto = oradb.Parameters("A_CAD_FLG_CTD_PRODUCTO")
Set SubtNetoProducto = oradb.Parameters("A_CAD_SUBT_NETO_PRODUCTO")
Set FlgPromocion = oradb.Parameters("A_CAD_FLG_PROMOCION")
Set NroLote = oradb.Parameters("A_CAD_NRO_LOTE")
Set FchVmto = oradb.Parameters("A_CAD_FCH_VMTO")
Set CadenaPromocion = oradb.Parameters("A_CAD_PROMOCION")
Set CadenaPromocionDcto = oradb.Parameters("A_CAD_PROMOCION_DCTO")

'Set SubtNetoProductoPP = oradb.Parameters("A_CAD_SUBT_NETO_PRODUCTO")

Set CadPuntos = oradb.Parameters("A_CAD_PUNTOS")
'Set CadPuntosRedime = oradb.Parameters("A_CAD_PUNTOS_REDIME")
'Recetario Magistral'
'Graba Documento'
Set Tip2Venta = oradb.Parameters("A_CAD2_TIP_VENTA")
Set Cod2Producto = oradb.Parameters("A_CAD2_COD_PRODUCTO")
Set Cod2ProductoIns = oradb.Parameters("A_CAD2_COD_PRODUCTO_INS")
Set Ctd2Producto = oradb.Parameters("A_CAD2_CTD_PRODUCTO")
Set Flg2CtdProducto = oradb.Parameters("A_CAD2_FLG_CTD_PRODUCTO")
Set Prc2UnitInsumo = oradb.Parameters("A_CAD2_PRC_UNIT_INSUMO")
Set Ctd2PctBase = oradb.Parameters("A_CAD2_PCT_BASE")
Set Ctd2Base = oradb.Parameters("A_CAD2_CTD_BASE")
Set UndMedida = oradb.Parameters("A_CAD2_UND_MEDIDA")

Set CodigoProductoEV = oradb.Parameters("A_ARR_COD_PRODUCTO")
Set NumeroEV = oradb.Parameters("A_ARR_NUM_ESPECIE_VAL")
Set PlacaEV = oradb.Parameters("A_ARR_NUM_PLACA")


Set CodProductoServicio = oradb.Parameters("A_CAD3_COD_PRODUCTO")
Set CodTipServicio = oradb.Parameters("A_CAD3_COD_TIP_SERVICIO")
Set CodServicio = oradb.Parameters("A_CAD3_COD_SERVICIO")
Set ImpNetoServicio = oradb.Parameters("A_CAD3_IMP_NETO_SERVICIO")
'Set NumIdentServ = oradb.Parameters("A_CAD3_NUM_IDENT_SERV")
Set NumVoucher = oradb.Parameters("A_CAD3_NUM_VOUCHER")
Set NumRecibo = oradb.Parameters("A_CAD3_NUM_RECIBO")
Set NumSuministro = oradb.Parameters("A_CAD3_NUM_SUMINISTRO")
Set FlgTipoPago = oradb.Parameters("A_CAD3_FLG_TIPO_PAGO")


'-------------------------DATO ADICIONAL-------------------------------

    For i = 0 To DatosAdicional.UpperBound(1)
        CodCampo(i) = DatosAdicional(i, 0)
        ValCampo(i) = DatosAdicional(i, 5)
    Next

'------------------------Especies valoradad SOAT----------------------------

    For i = 0 To xEspeciesValoradas.UpperBound(1)
        CodigoProductoEV(i) = xEspeciesValoradas(i, 0)
        NumeroEV(i) = xEspeciesValoradas(i, 1)
        PlacaEV(i) = xEspeciesValoradas(i, 2)
    Next


'------------------------COBRO POR RESPONSABILIDAD----------------------------
    
    For i = 0 To CobroResponsabilidad.UpperBound(1)
        CodUsuResp(i) = CobroResponsabilidad(i, 0)
        ImpUsuResp(i) = CobroResponsabilidad(i, 2)
    Next

'------------------------FORMA DE PAGO----------------------------------------
    
    
    
    
    
    For i = 0 To FormaPago.UpperBound(1)
        CodFormPago(i) = FormaPago(i, 0)
        CodFPagHijo(i) = FormaPago(i, 2)
        CodMonedaPago(i) = FormaPago(i, 6)
        TipCambio(i) = FormaPago(i, 11)
        CodTipTarjeta(i) = FormaPago(i, 5)
        NumTarjeta(i) = FormaPago(i, 12)
        NumAutorizacion(i) = FormaPago(i, 20) ''agregado Arturo escate
        FlgTipCuota(i) = FormaPago(i, 15)
        NumCuota(i) = FormaPago(i, 13)
        FchVencimiento(i) = FormaPago(i, 14)
        MtoRetiroEfec(i) = FormaPago(i, 29)      'cambiar por el verdadero
        CodDocPago(i) = FormaPago(i, 7)
        NumDocPago(i) = FormaPago(i, 16)
        NumDocNotaCred(i) = FormaPago(i, 22)
        FchDocNotaCred(i) = FormaPago(i, 21)
        CodDocIdent(i) = ""
        NumDocIdent(i) = FormaPago(i, 25)
        CodBanco(i) = FormaPago(i, 8)
        NumMovimiento(i) = FormaPago(i, 17)
        FchMovimiento(i) = FormaPago(i, 19)
        CtaCteBtl(i) = FormaPago(i, 10)
        CodDonacion(i) = FormaPago(i, 9)
        Importe(i) = FormaPago(i, 4)
    Next

'------------------------PARA EL DETALLE DE PRODUCTOS DE PROF-----------------
    Dim dblPrecioCero As Double
    Dim NumeroIngresoCliente As Integer
    'Dim VtaTotal As Double
    NumeroIngresoCliente = 0
    VtaTotal = 0
    
    For i = 0 To Producto.UpperBound(1)
        TipVenta(i) = Producto(i, 5)
        codProducto(i) = Producto(i, 0)
        CtdProducto(i) = Producto(i, 3)
        FlgCtdProducto(i) = Producto(i, 6)
        
        '++ Se ejecuta cuando la modalidad es convenio y esta marcado como deducible en base al precio publico
        '++ 09/01/2009 Por Cristhian Rueda
        If objVenta.CodModalidadVenta = "002" Then
            If objVenta.PrecioMenorDeducible = "1" Then
                dblDet = dblDet + Producto(i, 4)
            End If
        End If
        
        SubtNetoProducto(i) = Producto(i, 4)
        
        FlgPromocion(i) = Producto(i, 7)
        NroLote(i) = Producto(i, 22)
        FchVmto(i) = Producto(i, 23)
        CadPuntos(i) = Producto(i, 28)
        
        'CadPuntosRedime(i) = Round(Producto(i, 4) / VtaTotal, 3)
        CadenaPromocion(i) = Producto(i, 26)
        NumeroIngresoCliente = NumeroIngresoCliente + fnSolicitaCliente(Producto(i, 26))
        CadenaPromocionDcto(i) = Producto(i, 27)
    Next
    If NumeroIngresoCliente > 0 And Me.DesAuxCliNombre = "" Then
        MsgBox "La promoción solicita que ingrese los datos del cliente", vbCritical, App.ProductName
        mdiPrincipal.cmdDocumento_Click
        Exit Function
    End If
    
    
'------------------------PARA EL CASO DE RECETARIO MAGISTRAL------------------

'    If objVenta.CodigoTipoVenta = Recetario Then
'        If RecetarioMagistral.UpperBound(1) = -1 Then MsgBox "Se Encuentra en Modalidad de RECETARIO MAGISTRAL " & Chr(13) & " No Podra hacer la Venta por no tener INSUMOS", vbCritical, "Recetario Magistral": Exit Function
'        'If Producto.UpperBound(1) >= 0 Then MsgBox "Se Encuentra en la Modalidad RECETARIO MAGISTRAL, " & Chr(13) & " Solo Podra hacer Venta de INSUMOS", vbCritical, "Recetario Magistral": Exit Function
'      Else
'        If RecetarioMagistral.UpperBound(1) <> -1 Then MsgBox "Se Encuentra en Modalidad " & Trim(objVenta.NombreTipoVenta) & "  " & Chr(13) & "  No Podra hacer la Venta por Tener INSUMOS", vbCritical, objVenta.NombreTipoVenta: Exit Function
'    End If
'-----------------------------------------------------------------------------
    
    
    For i = 0 To RecetarioMagistral.UpperBound(1)
        Cod2Producto(i) = RecetarioMagistral(i, 0)
        Cod2ProductoIns(i) = RecetarioMagistral(i, 1)
        Ctd2Producto(i) = RecetarioMagistral(i, 3)
        Flg2CtdProducto(i) = "0"
        Prc2UnitInsumo(i) = RecetarioMagistral(i, 9)
        '''Tip2Venta(i) = RecetarioMagistral(i, 5)
        Tip2Venta(i) = CodigoTipoVenta
        Ctd2PctBase(i) = RecetarioMagistral(i, 7)
        Ctd2Base(i) = RecetarioMagistral(i, 6)
        UndMedida(i) = RecetarioMagistral(i, 8)
    Next
Dim flgConfirmaVentaRecarga As Boolean
flgConfirmaVentaRecarga = False
'-----------------------PARA EL CASO DE COBRO DE SERVICIOS------------------------------


    For i = 0 To Servicios.UpperBound(1)
        CodProductoServicio(i) = Servicios(i, 9)
        CodTipServicio(i) = Servicios(i, 0)
''''''''''''Esto se comento en enero09
''''''''''''       If EsNavsat(CodTipServicio(i)) = True Then
''''''''''''        Dim strNumeroAprobacion As String
''''''''''''        Dim strNumeroTraceOriginal As String
''''''''''''        Dim strNumeroPin As String
''''''''''''        '
''''''''''''        Dim ExisteNavsat As Boolean
''''''''''''
''''''''''''
''''''''''''        ExisteNavsat = False
''''''''''''           If ValidaTransaccionNavsat(CStr(xServicio(i, 11)), CStr(Servicios(i, 9)), CDbl(strMontoTotalTotal), strNumeroAprobacion, strNumeroTraceOriginal, strNumeroPin) = True Then
''''''''''''                    Servicios(i, 10) = strNumeroAprobacion 'NUMRECIBO
''''''''''''                    Servicios(i, 4) = strNumeroTraceOriginal 'NUMVOUCHER
''''''''''''                    Servicios(i, 11) = strNumeroPin ''numero de pin para las llamadas
''''''''''''                    ExisteNavsat = True
''''''''''''            Else
''''''''''''                MsgBox "mensaje de error"
''''''''''''           End If
''''''''''''        End If
        CodServicio(i) = Servicios(i, 2)
        ImpNetoServicio(i) = Servicios(i, 6)
        NumRecibo(i) = Servicios(i, 10)
        NumSuministro(i) = Servicios(i, 11)
        'NumIdentServ(i) = Servicios(i, 4)
        NumVoucher(i) = Servicios(i, 4)
        FlgTipoPago(i) = Servicios(i, 12)
    Next

    '++ Se ejecuta cuando la modalidad es convenio y esta marcado como deducible en base al precio publico
    '++ 09/01/2009 Por Cristhian Rueda
    If objVenta.CodModalidadVenta = "002" Then
        If objVenta.PrecioMenorDeducible = "1" Then
            dblNewPctBenef = Round(val(100 * (dblPctBeneficarioAux) / dblDet), 2)
            oradb.Parameters.Add "A_PCT_BENEFICIARIO", dblNewPctBenef, ORAPARM_INPUT, ORATYPE_VARCHAR2
            
          Else
          
            oradb.Parameters.Add "A_PCT_BENEFICIARIO", PctBeneficiario, ORAPARM_INPUT, ORATYPE_VARCHAR2
        End If
        
      Else
        oradb.Parameters.Add "A_PCT_BENEFICIARIO", PctBeneficiario, ORAPARM_INPUT, ORATYPE_VARCHAR2
    End If

'0=CodigoPadre As String
'1=DescripcionPadre As String
'2=CodigoHijo As String
'3=DescripcionHijo As String
'4=NumOperacion As String
'5=TipoVenta As TipoVenta
'6=Monto As String
'7=TipoCambio As String
'8=Importe As String
'9=Producto As String
'10=NumeroRecibo As String
'11=NumeroSuministro As String
'12=FlagTipoPago As String

If xServicio.Count(1) > 0 Then
i = 0
    If Not gblNumeroTelefonico = "" Then
        If gblNumeroTelefonico = CStr(xServicio(i, 11)) Then
            MsgBox "Esta intentando enviar, dos veces una transaccion de NAVSAT, presiono CANCELAR y vuelva intentarlo", vbCritical, App.ProductName
            GrabarDoc = False
            Exit Function
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''
'''''''    Select Case gclsOracle.FN_Valor("CMR.PKG_PRODUCTO.FN_LAB", xServicio(i, 9))
'''''''        Case gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_RUC_BRIGHTSTART")
'''''''            Dim strErrorCodNew As String
'''''''            Dim strErrorDesNew As String
'''''''            Dim strSecConfirmacion As String
'''''''            ValidaTransaccionNew strMontoTotalTotal, CStr(xServicio(i, 11)), objUsuario.NombrePC, strErrorCodNew, strErrorDesNew, strSecConfirmacion
'''''''            If Not strErrorCodNew = "0" Then
'''''''            Dim strMsg As String
'''''''                strMsg = gclsOracle.FN_Valor("BTLPROD.PKG_SERVICIO.FN_DEV_ERROR", gclsOracle.FN_Valor("CMR.PKG_PRODUCTO.FN_LAB", xServicio(i, 9)), strErrorCodNew)
'''''''                If strMsg = "-1" Then
'''''''                    MsgBox strErrorDesNew, vbCritical, "Error consultando al proveedor"
'''''''                Else
'''''''                    MsgBox strMsg
'''''''                End If
'''''''                GrabarDoc = False
'''''''                Exit Function
'''''''            Else
'''''''                Servicios(i, 4) = strSecConfirmacion
'''''''                NumVoucher(i) = Servicios(i, 4)
'''''''            End If
'''''''            flgConfirmaVentaRecarga = False
'''''''        Case gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_RUC_NAVSAT")
'''''''            flgConfirmaVentaRecarga = True
'''''''        Case Else
'''''''            MsgBox "No se ha defnido el proveedor de recargas", vbCritical, App.ProductName
'''''''            GrabarDoc = False
'''''''            Exit Function
'''''''    End Select
    '''''''''''''''''''''''''''''''''''''''
            'GrabarDoc = False
            'Exit Function
    
End If


oradb.ExecuteSQL " BEGIN BTLPROD.PKG_DOCUMENTO.SP_GRABA(:RET_PROMO_MENSAJE,:A_COD_CIA,:A_COD_LOCAL,:A_COD_MAQUINA,:A_COD_TIPO_VENTA,:A_COD_CLIENTE,:A_COD_TIP_DOC_VTA,:A_COD_EMPRESA,:A_COD_CONVENIO,:A_PCT_BENEFICIARIO,:A_COD_MEDICO,:A_COD_REPARTIDOR," _
        & ":A_IMP_TIP_CAMBIO,:A_NUM_PEDIDO_PADRE,:A_COD_CONCEPTO,:A_COD_SUBCONCEPTO_FACT,:A_COD_ORIGEN_DOCUMENTO,:A_COD_USUARIO,:A_COD_LOCAL_REFER,:A_COD_TIPO_DOC_REFER," _
        & ":A_NUM_DOCUMENTO_REFER,:A_FLG_NCR_DEVOLUCION,:A_CAD_COD_CAMPO,:A_CAD_VAL_CAMPO,:A_FLG_ORIGEN_DELIV,:A_FLG_TRANSF_LOCAL,:A_FLG_ENTREGA_LOCAL,:A_FLG_RESERVA_STK,:A_NUM_RUC_EMPRESA,:A_DES_RAZON_SOCIAL,:A_COD_DIRECCION_CLI," _
        & ":A_FLG_URGENTE,:A_DES_AUX_CLI_TLF,:A_DES_AUX_CLI_NOMBRE,:A_DES_AUX_CLI_DIRECC,:A_COD_UBIGEO,:A_COD_MOTIVO_COBRO,:A_CAD_COD_USU_RESP,:A_CAD_IMP_USU_RESP,:A_CAD_COD_FORM_PAGO,:A_CAD_COD_FPAG_HIJO,:A_CAD_COD_MONEDA_PAGO," _
        & ":A_CAD_TIP_CAMBIO,:A_CAD_COD_TIP_TARJETA,:A_CAD_NUM_TARJETA,:A_CAD_FLG_TIP_CUOTA,:A_CAD_NUM_CUOTA,:A_CAD_FCH_VENCIMIENTO,:A_CAD_MTO_RETIRO_EFEC,:A_CAD_NUM_AUTORIZACION,:A_CAD_COD_DOC_PAGO,:A_CAD_NUM_DOC_PAGO," _
        & ":A_NUM_DOC_NOTA_CRED,:A_FCH_DOC_NOTA_CRED,:A_CAD_COD_DOC_IDENT,:A_CAD_NUM_DOC_IDENT,:A_CAD_COD_BANCO,:A_CAD_NUM_MOVIMIENTO,:A_CAD_FCH_MOVIMIENTO,:A_CAD_CTA_CTE_BTL," _
        & ":A_CAD_COD_DONACION,:A_CAD_IMPORTE,:A_CAD_TIP_VENTA,:A_CAD_COD_PRODUCTO,:A_CAD_CTD_PRODUCTO,:A_CAD_FLG_CTD_PRODUCTO,:A_CAD_SUBT_NETO_PRODUCTO,:A_CAD_FLG_PROMOCION,:A_CAD2_TIP_VENTA," _
        & ":A_CAD2_COD_PRODUCTO,:A_CAD2_COD_PRODUCTO_INS,:A_CAD2_CTD_PRODUCTO,:A_CAD2_FLG_CTD_PRODUCTO,:A_CAD2_PRC_UNIT_INSUMO,:A_CAD2_PCT_BASE,:A_CAD2_CTD_BASE,:A_CAD2_UND_MEDIDA, " _
        & ":A_CAD3_COD_PRODUCTO, :A_CAD3_COD_TIP_SERVICIO,:A_CAD3_COD_SERVICIO,:A_CAD3_IMP_NETO_SERVICIO,:A_CAD3_NUM_VOUCHER, :A_CAD3_NUM_RECIBO, :A_CAD3_NUM_SUMINISTRO, :A_CAD3_FLG_TIPO_PAGO, " _
        & ":A_ARR_COD_PRODUCTO , :A_ARR_NUM_ESPECIE_VAL,:A_ARR_NUM_PLACA, " _
        & ":A_TIPDOC_GENERAD, :A_NUMDOC_GENERAD, :A_TIPDOC_GENERAD_CO, :A_NUMDOC_GENERAD_CO, :A_CAD_COD_DIAGNOSTICO, :A_FCH_RECETA, :A_COD_LOCAL_RECETA,:A_NUM_RECETA, :A_DEV_EFECTIVO_NC, :A_COD_SERIE_REL, :A_CAD_NRO_LOTE, :A_CAD_FCH_VMTO, :A_PCT_BENEFICIARIO_ORIG, :A_CAD_PROMOCION,:A_CAD_PROMOCION_DCTO,:A_DES_OBSERVACION, " _
        & ":A_COD_MOTIVO_DEV,:A_EXISTE_DNI_RENIEC,:A_CAD_PUNTOS,:A_VENTA_TOTAL,:A_PUNTOS_REDIME,:A_PUNTOS_INICIAL,:A_PUNTOS_ACUMULADO,:A_NUM_TARJ_PUNTOS) ; END;"
        


''    oradb.ExecuteSQL " BEGIN BTLPROD.PKG_DOCUMENTO.SP_GRABA(:A_COD_CIA,:A_COD_LOCAL,:A_COD_MAQUINA,:A_COD_TIPO_VENTA,:A_COD_CLIENTE,:A_COD_TIP_DOC_VTA,:A_COD_EMPRESA,:A_COD_CONVENIO,:A_PCT_BENEFICIARIO,:A_COD_MEDICO," _
''        & ":A_IMP_TIP_CAMBIO,:A_NUM_PEDIDO_PADRE,:A_COD_CONCEPTO,:A_COD_SUBCONCEPTO_FACT,:A_COD_ORIGEN_DOCUMENTO,:A_COD_USUARIO,:A_COD_LOCAL_REFER,:A_COD_TIPO_DOC_REFER," _
''        & ":A_NUM_DOCUMENTO_REFER,:A_FLG_NCR_DEVOLUCION,:A_CAD_COD_CAMPO,:A_CAD_VAL_CAMPO,:A_FLG_ORIGEN_DELIV,:A_FLG_TRANSF_LOCAL,:A_FLG_ENTREGA_LOCAL,:A_FLG_RESERVA_STK,:A_NUM_RUC_EMPRESA,:A_DES_RAZON_SOCIAL,:A_COD_DIRECCION_CLI," _
''        & ":A_FLG_URGENTE,:A_DES_AUX_CLI_TLF,:A_DES_AUX_CLI_NOMBRE,:A_DES_AUX_CLI_DIRECC,:A_COD_UBIGEO,:A_COD_MOTIVO_COBRO,:A_CAD_COD_USU_RESP,:A_CAD_IMP_USU_RESP,:A_CAD_COD_FORM_PAGO,:A_CAD_COD_FPAG_HIJO,:A_CAD_COD_MONEDA_PAGO," _
''        & ":A_CAD_TIP_CAMBIO,:A_CAD_COD_TIP_TARJETA,:A_CAD_NUM_TARJETA,:A_CAD_FLG_TIP_CUOTA,:A_CAD_NUM_CUOTA,:A_CAD_FCH_VENCIMIENTO,:A_CAD_MTO_RETIRO_EFEC,:A_CAD_NUM_AUTORIZACION,:A_CAD_COD_DOC_PAGO,:A_CAD_NUM_DOC_PAGO," _
''        & ":A_NUM_DOC_NOTA_CRED,:A_FCH_DOC_NOTA_CRED,:A_CAD_COD_DOC_IDENT,:A_CAD_NUM_DOC_IDENT,:A_CAD_COD_BANCO,:A_CAD_NUM_MOVIMIENTO,:A_CAD_FCH_MOVIMIENTO,:A_CAD_CTA_CTE_BTL," _
''        & ":A_CAD_COD_DONACION,:A_CAD_IMPORTE,:A_CAD_TIP_VENTA,:A_CAD_COD_PRODUCTO,:A_CAD_CTD_PRODUCTO,:A_CAD_FLG_CTD_PRODUCTO,:A_CAD_SUBT_NETO_PRODUCTO,:A_CAD_FLG_PROMOCION,:A_CAD2_TIP_VENTA," _
''        & ":A_CAD2_COD_PRODUCTO,:A_CAD2_COD_PRODUCTO_INS,:A_CAD2_CTD_PRODUCTO,:A_CAD2_FLG_CTD_PRODUCTO,:A_CAD2_PRC_UNIT_INSUMO, " _
''        & ":A_CAD3_COD_PRODUCTO, :A_CAD3_COD_TIP_SERVICIO,:A_CAD3_COD_SERVICIO,:A_CAD3_IMP_NETO_SERVICIO,:A_CAD3_NUM_VOUCHER, :A_CAD3_NUM_RECIBO, :A_CAD3_NUM_SUMINISTRO, :A_CAD3_FLG_TIPO_PAGO, " _
''        & ":A_ARR_COD_PRODUCTO , :A_ARR_NUM_ESPECIE_VAL,:A_ARR_NUM_PLACA, " _
''        & ":A_TIPDOC_GENERAD, :A_NUMDOC_GENERAD, :A_TIPDOC_GENERAD_CO, :A_NUMDOC_GENERAD_CO) ; END;"
''






    Set oTipDoc = oradb.Parameters("A_TIPDOC_GENERAD")
    Set oNumDoc = oradb.Parameters("A_NUMDOC_GENERAD")
    Set oTipDocCo = oradb.Parameters("A_TIPDOC_GENERAD_CO")
    Set oNumDocCo = oradb.Parameters("A_NUMDOC_GENERAD_CO")
    
    RetPromoMensaje = Trim("" & oradb.Parameters("RET_PROMO_MENSAJE").Value)

GrabarDoc = True
  

''MEJORANDO NAVSAT
If xServicio.Count(1) > 0 Then
    gblNumeroTelefonico = CStr(xServicio(i, 11))
    If EsNavsat(xServicio(0, 0)) = True Then
            Dim strNumeroAprobacion As String
            Dim strNumeroTraceOriginal As String
            Dim strNumeroPin As String
            Dim TramaEnvioOut As String
            Dim TramaRepuestaOut As String
            Dim inNumDocumento As String
            Dim inTipoDocumento As String
            Dim inNumRecibo As String
            Dim inTipoRecibo As String
            Dim objServicio As New clsServicio
            Dim flgConfirmar As String
            Dim u As Integer
            i = 0
            u = 0
            flgConfirmar = "0"
            While u <= 3
                If oTipDoc(u) = "BOL" Or oTipDoc(u) = "TKB" Or oTipDoc(u) = "FAC" Or oTipDoc(u) = "TKF" Then
                    inNumDocumento = oNumDoc(u)
                    inTipoDocumento = oTipDoc(u)
                ElseIf oTipDoc(u) = "REC" Then
                    inNumRecibo = oNumDoc(u)
                    inTipoRecibo = oTipDoc(u)
                End If
                u = u + 1
            Wend
sale:

'?????????????????????????????????????????????????????
'?????????????????????????????????????????????????????
'?????????????????????????????????????????????????????
'?????????????????????????????????????????????????????
'?????????????????????????????????????????????????????
    Select Case gclsOracle.FN_Valor("CMR.PKG_PRODUCTO.FN_LAB", xServicio(i, 9))
        ''aca ejecuta la dll de brightstart
        Case gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_RUC_BRIGHTSTART")
            Dim strErrorCodNew As String
            Dim strErrorDesNew As String
            Dim strSecConfirmacion As String
            ValidaTransaccionNew strMontoTotalTotal, CStr(xServicio(i, 11)), objUsuario.NombrePC, strErrorCodNew, strErrorDesNew, strSecConfirmacion, CStr(Servicios(i, 9))
            If Not strErrorCodNew = "0" Then
            Dim strMsg As String
                strMsg = "" & gclsOracle.FN_Valor("BTLPROD.PKG_SERVICIO.FN_DEV_ERROR", gclsOracle.FN_Valor("CMR.PKG_PRODUCTO.FN_LAB", xServicio(i, 9)), strErrorCodNew)
                MsgBox IIf(strMsg = "", strErrorDesNew, strMsg), vbCritical, "Error consultando al proveedor"
                flgConfirmar = "0"
            Else
                strNumeroAprobacion = strSecConfirmacion 'NUMRECIBO
                strNumeroTraceOriginal = Servicios(i, 4) 'NUMVOUCHER
                strNumeroPin = CStr(xServicio(i, 11)) ''numero de pin para las llamadas
                flgConfirmar = "1"
            End If
            TramaRepuestaOut = "" & IIf(strMsg = "", strErrorDesNew, strMsg)
            TramaEnvioOut = strMontoTotalTotal & "/" & CStr(xServicio(i, 11)) & "/" & objUsuario.NombrePC & "/" & strErrorCodNew & "/" & "" & IIf(strMsg = "", strErrorDesNew, strMsg) & "/" & strSecConfirmacion
        ''aca ejecuta la dll de navsat
        Case gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_RUC_NAVSAT")
            If ValidaTransaccionNavsat(CStr(xServicio(i, 11)), CStr(Servicios(i, 9)), CDbl(strMontoTotalTotal), strNumeroAprobacion, strNumeroTraceOriginal, strNumeroPin, TramaEnvioOut, TramaRepuestaOut, SecuenciaGrabacion) = True Then
                    flgConfirmar = "1"
            Else
                    flgConfirmar = "0"
            End If
        Case Else
            MsgBox "No se ha defnido el proveedor de recargas", vbCritical, App.ProductName
            GrabarDoc = False
            Exit Function
    End Select



'?????????????????????????????????????????????????????
'?????????????????????????????????????????????????????
'?????????????????????????????????????????????????????
'?????????????????????????????????????????????????????
'?????????????????????????????????????????????????????
            Servicios(i, 10) = strNumeroAprobacion 'NUMRECIBO
            Servicios(i, 4) = strNumeroTraceOriginal 'NUMVOUCHER
            Servicios(i, 11) = strNumeroPin ''numero de pin para las llamadas
            
            Dim Mensaje As String
            Dim DocumentoAnulados As String
            Mensaje = objServicio.ConfirmaVenta(objUsuario.CodigoEmpresa, inTipoDocumento, objUsuario.CodigoLocal, inNumDocumento, strNumeroAprobacion, strNumeroTraceOriginal, strNumeroPin, TramaRepuestaOut, objUsuario.Codigo, TramaEnvioOut, objUsuario.CodigoLiquidacion, CDbl(strMontoTotalTotal), Servicios(i, 11), CStr(Servicios(i, 9)), SecuenciaGrabacion, flgConfirmar, inTipoRecibo, inNumRecibo, DocumentoAnulados)
            objServicio.logRecarga TramaRepuestaOut, objUsuario.Codigo, TramaEnvioOut, objUsuario.CodigoLiquidacion, CDbl(strMontoTotalTotal), CStr(xServicio(i, 11)), CStr(Servicios(i, 9)), objUsuario.CodigoEmpresa, inTipoDocumento, inNumDocumento, IIf(Mensaje = "", "OK", Mensaje), SecuenciaGrabacion
            If flgConfirmar = "0" Then
                If Not Mensaje = "" Then
                    MsgBox Mensaje, vbCritical, App.ProductName
                    FlgAnulados = True
                Else
                    MsgBox "ADVERTENCIA" & Chr(13) & "Se anularon los siguiente documentos" & Chr(13) & DocumentoAnulados & Chr(13) & "ESTOS DOCUMENTOS NO DEBEN ENTREGARSE AL CLIENTE.", vbCritical, App.ProductName
                    FlgAnulados = True
                End If
            Else
                FlgAnulados = False
            End If
            Set objServicio = Nothing

    End If
    End If
''DESMEJORANDO NAVSAT
Exit Function

''''''''''Control:
''''''''''    Dim strError As String
''''''''''    Dim strErrorNavSAT As String
''''''''''
''''''''''    strError = Err.Description
''''''''''
''''''''''    If ExisteNavsat = True Then
''''''''''        i = 0
''''''''''        If EsNavsat(xServicio(0, 0)) = True Then
''''''''''            If Not AnulaTransaccionNavsat(CStr(xServicio(i, 11)), CStr(Servicios(i, 9)), CDbl(strMontoTotalTotal), Servicios(i, 10), Servicios(i, 4)) Then
''''''''''            'ValidaTransaccionNavsat(CStr(xServicio(i, 11)), CStr(Servicios(i, 9)), CDbl(strMontoTotalTotal), strNumeroAprobacion, strNumeroTraceOriginal)
''''''''''                strErrorNavSAT = "Existe un error al anular la transaccion NAVSAT" 'Err.Description
''''''''''            End If
''''''''''        End If
''''''''''        'GrabarDoc = False
''''''''''        '            Err.Raise vbObjectError + 1, "clsVenta.GrabarDoc", "Error en transaccion de NAVSAT " & Err.Description
''''''''''    End If
        
        
'''''''''Salida:
'''''''''    GrabarDoc = False
'''''''''
'''''''''    strError = "Se produjo el siguiente error al registrar la transacción." & Chr(13) & _
'''''''''                strError & _
'''''''''                IIf(strErrorNavSAT = "", "", "Error NAVSAT" & Chr(13) & strErrorNavSAT)
'''''''''
'''''''''    Err.Raise vbObjectError + 1, "clsVenta.GrabarDoc", strError


End Function

Public Property Get CodMotivoCobro() As String
    CodMotivoCobro = strCodMotivoCobro
End Property

Public Property Let CodMotivoCobro(ByVal lstrCodMotivoCobro As String)
    strCodMotivoCobro = lstrCodMotivoCobro

End Property



Public Function ListaMotivoCobro(Optional ByVal vstrCodMotivoCobro As String = "") As OracleInProcServer.oraDynaset
    
    On Error GoTo Pase
    Set ListaMotivoCobro = gclsOracle.FN_Cursor("BTLPROD.PKG_COBRANZA.FN_LISTA_MOTIVO_COBRO", 0, vstrCodMotivoCobro)
    
    Exit Function
Pase:
    Err.Raise Err.Number, "clsVenta.ListaMotivoCobro", Err.Description
End Function


'*************************************************************************************
' Metodo para agregar un producto al arreglo xProductoCobro
'*************************************************************************************

''Se cambio por una nueva función que esta al final

''''Public Function AgregaProductoCobro(Codigo As String, _
''''                            Descripcion As String, _
''''                            Cantidad As Integer, _
''''                            FlagFraccion As String, _
''''                            Precio As Double, _
''''                            NewCantidad As Integer, _
''''                            NewFlgFraccion As String, _
''''                            NewPrecio As Double) As XArrayDB
''''    Dim ultimo As Integer ' declara variable contador
''''    Dim aux As Integer
''''    If xProductoCobro.Count(1) < 0 Then Exit Function
''''
''''    Dim i As Integer
''''    Dim encontro As Boolean
''''
''''
''''    aux = xProductoCobro.Count(1)
''''    While i < aux
''''        If xProductoCobro(i, 0) = Codigo Then
''''            ultimo = i
''''            encontro = True
''''GoTo j
''''        Else
''''            encontro = False
''''            ultimo = xProductoCobro.Count(1) 'Llena con la ultima posicion disponible
''''        End If
''''        i = i + 1
''''    Wend
''''    If encontro = False Then
''''        xProductoCobro.AppendRows
''''    End If
''''j:
''''    If xProductoCobro.Count(1) = 0 Then ultimo = 0: xProductoCobro.AppendRows
''''    xProductoCobro(ultimo, 0) = 0          ' Producto Seleccionado
''''    xProductoCobro(ultimo, 1) = Codigo          ' Asigna el codigo de producto
''''    xProductoCobro(ultimo, 2) = Descripcion     ' Asigna la descripcion del producto
''''    xProductoCobro(ultimo, 3) = FlagFraccion    ' Asigna si la cantidad es fraccion o Unidad
''''    xProductoCobro(ultimo, 4) = Cantidad        ' Asigna la cantidad seleccionada del producto
''''    xProductoCobro(ultimo, 5) = Precio          ' Asigna el tipo de Venta
''''    xProductoCobro(ultimo, 6) = NewCantidad     ' Asigna el tipo de Venta
''''    xProductoCobro(ultimo, 7) = NewFlgFraccion  ' Asigna si la cantidad es fraccion o Unidad
''''    xProductoCobro(ultimo, 8) = NewPrecio       'valor libre
''''    Set AgregaProductoCobro = xProductoCobro
''''
''''End Function


Public Sub CargarProductoCobro(ByVal pstrTipDoc As String, ByVal pstrNumDoc As String)
Dim oraProductoCobro As oraDynaset
Dim objDocumento As New clsDocumento
Dim strCodigo As String
Dim strDescripcion As String
Dim intCtdProducto As Integer
Dim intCtdProductoFrac As Integer
Dim intCtdFraccionamiento As Integer
Dim dblMtoSubtotal As Double
Dim dblPrcUnitNetoVta As Double
Dim intCantidad As Integer
Dim strFlagFracc As String

        On Error GoTo CtrlErr
        
        xProductoCobro.ReDim 0, -1, 0, 11
        ''Se cambio por una función nueva
        ''Set oraProductoCobro = objDocumento.ListaDetalle(objUsuario.CodigoEmpresa, pstrNumDoc, pstrTipDoc, objUsuario.CodigoLocal)
        Set oraProductoCobro = objDocumento.ListaDetalleCXR(objUsuario.CodigoEmpresa, pstrNumDoc, pstrTipDoc, objUsuario.CodigoLocal)
        
        If oraProductoCobro.RecordCount = 0 Then
          Set oraProductoCobro = objDocumento.Lista_DDocumentos(pstrNumDoc, pstrTipDoc)
        End If
        
        intCtdProducto = 0
        intCtdProductoFrac = 0
        intCtdFraccionamiento = 0

        dblMtoSubtotal = 0
        dblPrcUnitNetoVta = 0

        While Not oraProductoCobro.EOF
            intCantidad = 0
            
            strCodigo = oraProductoCobro("COD_PRODUCTO").Value
            strDescripcion = oraProductoCobro("DES_PRODUCTO").Value
            intCtdProducto = oraProductoCobro("CTD_PRODUCTO").Value
            intCtdProductoFrac = oraProductoCobro("CTD_PRODUCTO_FRAC").Value
            intCtdFraccionamiento = oraProductoCobro("CTD_FRACCIONAMIENTO").Value
            dblMtoSubtotal = oraProductoCobro("MTO_SUBTOTAL").Value
            dblPrcUnitNetoVta = oraProductoCobro("PRC_UNIT_NETO_VTA").Value
            
            If intCtdProducto > 0 And intCtdProductoFrac > 0 Then
                intCantidad = intCtdProducto * intCtdFraccionamiento + intCtdProductoFrac
                strFlagFracc = "F"
            End If
            
            If intCtdProducto > 0 And intCtdProductoFrac = 0 Then
                intCantidad = intCtdProducto
                strFlagFracc = "U"
            End If
        
            If intCtdProducto = 0 And intCtdProductoFrac > 0 Then
                intCantidad = intCtdProductoFrac
                strFlagFracc = "F"
            End If
        
            AgregaProductoCobro strCodigo, strDescripcion, intCtdProducto, intCtdProductoFrac, dblPrcUnitNetoVta, intCtdFraccionamiento, dblMtoSubtotal, 0, 0, 0
            oraProductoCobro.MoveNext
   

        Wend
        Exit Sub
        
CtrlErr:
    
    Err.Raise Err.Number, "clsVenta.CargarProductoCobro", Err.Description
End Sub


Public Property Get FlgTipoConvenio() As Variant
    FlgTipoConvenio = strFlgTipoConvenio
End Property

Public Property Let FlgTipoConvenio(ByVal vNewValue As Variant)
    strFlgTipoConvenio = vNewValue
End Property


Public Sub OrdenaDoc(ByVal oradb As OraDatabase, ByVal oTipDoc As OraParamArray, _
                                ByVal oNumDoc As OraParamArray, _
                                ByVal oTipDocCo As OraParamArray, _
                                ByVal oNumDocCo As OraParamArray, _
                                ByRef oAuxTipDoc As OraParamArray, _
                                ByRef oAuxNumDoc As OraParamArray)
                                
On Error GoTo Control
                                
                                
Dim i As Integer
                                

For i = oradb.Parameters.Count - 1 To 0 Step -1

    oradb.Parameters.Remove i

Next


                                
oradb.Parameters.AddTable "A_TIPDOC_GENERAD", ORAPARM_INPUT, ORATYPE_VARCHAR2, oTipDoc.ArraySize, 200
oradb.Parameters.AddTable "A_NUMDOC_GENERAD", ORAPARM_INPUT, ORATYPE_VARCHAR2, oNumDoc.ArraySize, 200
oradb.Parameters.AddTable "A_TIPDOC_GENERAD_CO", ORAPARM_INPUT, ORATYPE_VARCHAR2, oTipDocCo.ArraySize, 200
oradb.Parameters.AddTable "A_NUMDOC_GENERAD_CO", ORAPARM_INPUT, ORATYPE_VARCHAR2, oNumDocCo.ArraySize, 200

oradb.Parameters.AddTable "A_AUX_TIPDOC_GENERAD", ORAPARM_OUTPUT, ORATYPE_VARCHAR2, 100, 200
oradb.Parameters.AddTable "A_AUX_NUMDOC_GENERAD", ORAPARM_OUTPUT, ORATYPE_VARCHAR2, 100, 200

                                
                                
Dim TipDoc As OracleInProcServer.OraParamArray
Dim NumDoc As OracleInProcServer.OraParamArray
Dim TipDocCo As OracleInProcServer.OraParamArray
Dim NumDocCo As OracleInProcServer.OraParamArray
                                
                                
Set TipDoc = oradb.Parameters("A_TIPDOC_GENERAD")
Set NumDoc = oradb.Parameters("A_NUMDOC_GENERAD")
Set TipDocCo = oradb.Parameters("A_TIPDOC_GENERAD_CO")
Set NumDocCo = oradb.Parameters("A_NUMDOC_GENERAD_CO")



    For i = 0 To oTipDoc.ArraySize - 1
        TipDoc(i) = oTipDoc.get_Value(i)
    Next

    For i = 0 To oNumDoc.ArraySize - 1
        NumDoc(i) = oNumDoc.get_Value(i)
    Next
    For i = 0 To oTipDocCo.ArraySize - 1
        TipDocCo(i) = oTipDocCo.get_Value(i)
    Next
    For i = 0 To oNumDocCo.ArraySize - 1
        NumDocCo(i) = oNumDocCo.get_Value(i)
    Next


                               
                                
    oradb.ExecuteSQL " BEGIN BTLPROD.PKG_DOCUMENTO.SP_GRABA_ORDENA_DOC( " & _
                    ":A_TIPDOC_GENERAD, :A_NUMDOC_GENERAD, :A_TIPDOC_GENERAD_CO, :A_NUMDOC_GENERAD_CO, :A_AUX_TIPDOC_GENERAD, :A_AUX_NUMDOC_GENERAD ) ; END;"



    Set oAuxTipDoc = oradb.Parameters("A_AUX_TIPDOC_GENERAD")
    Set oAuxNumDoc = oradb.Parameters("A_AUX_NUMDOC_GENERAD")
                                
                                
                                
Exit Sub

Control:

    Err.Raise Err.Number, "clsVenta.OrdenaDoc", Err.Description

End Sub

'''Public Function LimpiaProyecto()
'''
''' '   On Error GoTo CtrlErr
'''    frmPedido.psub_BeginArry
'''    frm_VTA_Busqueda.grdProductos.Limpiar
'''    frm_VTA_Busqueda.grdAlternativos.Limpiar
'''    frm_VTA_Busqueda.grdComplementarios.Limpiar
''''CtrlErr:
''''        Err.Raise Err.Number, "clsVenta.LimpiaProyecto", Err.Description
'''End Function
Function LimpiaConvenio()
    Me.CodigoConvenio = ""
    Me.NombreConvenio = ""
    Me.CodigoBeneficiario = ""
    Me.NombreBeneficiario = ""
objVenta.NumMaximoUnidades = 0
    
    
    
    
    
    xConvenio.ReDim 0, -1, 0, 2
    
End Function



Public Function LimpiaServicio()
'busco posicion en el array
' guardo servicios en un array de posisiciones
'bucle y borro todos los servicios segun el array de posiciones

Dim arrServicios() As Variant
Dim i As Integer
Dim j As Integer
' hacer un bucle y marcar los servicios (5) tipo de venta
arrServicios = Array("")
While i < xProducto.Count(1)
    
    If Trim(xProducto(i, 5)) = Servicio Then  ''leer de la constante de oracle
        ReDim Preserve arrServicios(UBound(arrServicios) + 1)
        arrServicios(j) = i
        j = j + 1
    End If
    i = i + 1
Wend
Dim k As Integer
    While k < UBound(arrServicios)
       xProducto.Delete 1, arrServicios(k)
        k = k + 1
    Wend
    
    xServicio.ReDim 0, -1, 0, 13
    
End Function

Public Property Get TotalImporteAFavor() As Variant

    Dim i As Integer
    Dim dImporte As Variant
    
    
    dImporte = Array("0")
    
    For i = 0 To xProductoCobro.UpperBound(1)
            dImporte(0) = dImporte(0) + CDbl(xProductoCobro(i, 7))
    Next i
    
    TotalImporteAFavor = dImporte

End Property
Public Function LimpiaProductos()
    xProducto.ReDim 0, -1, 0, 29 'Inicializa los productos seleccionados
End Function

Public Function LimpiaDatoAdicional()
    xDatoAdicional.ReDim 0, -1, 0, 6
End Function

Public Function LimpiaRecetario()
    xRecetarioMagistral.ReDim 0, -1, 0, 9 'Inicializa el recetario magistral
End Function

Public Function LimpiaEspecieValorada(codigoProducto As String) As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim arrEncontrado() As Variant
    arrEncontrado = Array("")
    While i < xEspeciesValoradas.Count(1)
        If xEspeciesValoradas(i, 0) = codigoProducto Then
            ReDim Preserve arrEncontrado(UBound(arrEncontrado) + 1)
            arrEncontrado(j) = i
            j = j + 1
        End If
        i = i + 1
    Wend
    Dim k As Integer
    
    While k < UBound(arrEncontrado)
       xEspeciesValoradas.Delete 1, arrEncontrado(k)
        k = k + 1
    Wend
End Function


Public Property Get NumPedidoPadre() As String
    NumPedidoPadre = strNumPedidoPadre
End Property

Public Property Let NumPedidoPadre(ByVal vNewValue As String)
    strNumPedidoPadre = vNewValue
End Property

'*************************************************************************************
' Metodo nuevo para agregar un producto al arreglo xProductoCobro
'*************************************************************************************


Public Function AgregaProductoCobro(Codigo As String, _
                            Descripcion As String, _
                            cantidad As Integer, _
                            CantidadFracc As Integer, _
                            PrecioUniNetoVta As Double, _
                            CantFraccionamiento As Integer, _
                            MontoSubTotal As Double, _
                            NewCantidad As Integer, _
                            NewCantidadFracc As Integer, _
                            NewMontoSubTotal As Double) As XArrayDB
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If xProductoCobro.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    
    
    aux = xProductoCobro.Count(1)
    While i < aux
        If xProductoCobro(i, 0) = Codigo Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xProductoCobro.Count(1) 'Llena con la ultima posicion disponible
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xProductoCobro.AppendRows
    End If
j:
    If xProductoCobro.Count(1) = 0 Then ultimo = 0: xProductoCobro.AppendRows
    xProductoCobro(ultimo, 0) = 0          ' Producto Seleccionado
    xProductoCobro(ultimo, 1) = Codigo          ' Asigna el codigo de producto
    xProductoCobro(ultimo, 2) = Descripcion     ' Asigna la descripcion del producto
    xProductoCobro(ultimo, 3) = cantidad            ' Asigna si la cantidad es fraccion o Unidad
    xProductoCobro(ultimo, 4) = CantidadFracc        ' Asigna la cantidad seleccionada del producto
    xProductoCobro(ultimo, 5) = PrecioUniNetoVta        ' Asigna el tipo de Venta
    xProductoCobro(ultimo, 6) = CantFraccionamiento     ' Asigna el tipo de Venta
    xProductoCobro(ultimo, 7) = MontoSubTotal           ' Asigna si la cantidad es fraccion o Unidad
    xProductoCobro(ultimo, 8) = NewCantidad         ' Asigna si la cantidad es fraccion o Unidad
    xProductoCobro(ultimo, 9) = NewCantidadFracc       'valor libre
    xProductoCobro(ultimo, 10) = NewMontoSubTotal       'valor libre
    xProductoCobro(ultimo, 11) = "Codigo : " & Codigo & Chr(13) & "Descripción : " & Descripcion & Chr(13) & "Unid.: " & str(cantidad) & Chr(13) & "Fracc.: " & str(CantidadFracc) & Chr(13) & " Monto Sub Total S/. : " & Format(MontoSubTotal, "0.00")
    Set AgregaProductoCobro = xProductoCobro
    
End Function



Public Static Property Get DesReferenciaCli() As String
DesReferenciaCli = strDesReferenciaCli
End Property

Public Static Property Let DesReferenciaCli(ByVal lstrDesReferenciaCli As String)
strDesReferenciaCli = lstrDesReferenciaCli
End Property









Public Property Get Out_CodigoCliente() As String
    Out_CodigoCliente = strOut_CodigoCliente
End Property

Public Property Let Out_CodigoCliente(ByVal lstrOut_CodigoCliente As String)
    strOut_CodigoCliente = lstrOut_CodigoCliente

End Property

Public Property Get Out_NombreCliente() As String
    Out_NombreCliente = strOut_NombreCliente
End Property

Public Property Let Out_NombreCliente(ByVal lstrOut_NombreCliente As String)
    strOut_NombreCliente = lstrOut_NombreCliente

End Property




Public Property Get Out_NumeroId() As String
    Out_NumeroId = strOut_NumeroId
End Property

Public Property Let Out_NumeroId(ByVal lstrOut_NumeroId As String)
    strOut_NumeroId = lstrOut_NumeroId

End Property

Public Property Get Out_Tipo() As String
    Out_Tipo = strOut_Tipo
End Property

Public Property Let Out_Tipo(ByVal lstrOut_Tipo As String)
    strOut_Tipo = lstrOut_Tipo

End Property

Public Property Get Out_Direccion() As String
    Out_Direccion = strOut_Direccion
End Property

Public Property Let Out_Direccion(ByVal lstrOut_Direccion As String)
    strOut_Direccion = lstrOut_Direccion

End Property








Public Property Get ObsNotaMotorizado() As String
    ObsNotaMotorizado = strObsNotaMotorizado
End Property

Public Property Let ObsNotaMotorizado(ByVal lstrObsNotaMotorizado As String)
    strObsNotaMotorizado = lstrObsNotaMotorizado
    
End Property

Public Property Get LongTarj() As String
    LongTarj = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_LONG_TARJ")
End Property

Public Sub RemoverFormaPago(ByVal CodigoFormaPago As String, _
                         ByVal CodigoFormaPagoHijo As String, _
                         Optional NumeroTarjeta As String = "" _
                         )
                         
Dim i As Integer
    If xFormaPago.Count(1) < 0 Then Exit Sub
    
    
    
    
    For i = 0 To xFormaPago.UpperBound(1)
        If xFormaPago(i, 0) = CodigoFormaPago And xFormaPago(i, 2) = CodigoFormaPagoHijo And xFormaPago(i, 12) = NumeroTarjeta Then
            xFormaPago.DeleteRows (i)
            Exit For
        End If
    Next i
    
    
    
End Sub


'Private strNombreClienteDLV  As String
'Private strDireccionClienteDLV As String
'Private strDesDistritoDLV As String
'Private strDesUrbanizacionDLV As String
'Private strObservacionClienteDLV As String




Public Property Get DesDistritoDLV() As String
    DesDistritoDLV = strDesDistritoDLV
End Property

Public Property Let DesDistritoDLV(ByVal vstrDesDistritoDLV As String)
    strDesDistritoDLV = vstrDesDistritoDLV
End Property

Public Property Get DesUrbanizacionDLV() As String
    DesUrbanizacionDLV = strDesUrbanizacionDLV
End Property

Public Property Let DesUrbanizacionDLV(ByVal vstrDesUrbanizacionDLV As String)
    strDesUrbanizacionDLV = vstrDesUrbanizacionDLV
End Property
'MotivoNotaCredito

Public Property Get ObservacionClienteDLV() As String
    ObservacionClienteDLV = strObservacionClienteDLV
End Property

Public Property Let ObservacionClienteDLV(ByVal vstrObservacionClienteDLV As String)
    strObservacionClienteDLV = vstrObservacionClienteDLV
End Property

Public Property Get NombreClienteDLV() As String
    NombreClienteDLV = strNombreClienteDLV
End Property

Public Property Let NombreClienteDLV(ByVal vstrNombreClienteDLV As String)
    strNombreClienteDLV = vstrNombreClienteDLV
End Property

Public Property Get DireccionClienteDLV() As String
DireccionClienteDLV = strDireccionClienteDLV
End Property

Public Property Let DireccionClienteDLV(ByVal vstrDireccionClienteDLV As String)
    strDireccionClienteDLV = vstrDireccionClienteDLV
End Property

Public Property Get Out_CodDocumentoId() As String
    Out_CodDocumentoId = strOut_CodDocumentoId
End Property

Public Property Let Out_CodDocumentoId(ByVal vstrOut_CodDocumentoId As String)
    strOut_CodDocumentoId = vstrOut_CodDocumentoId
End Property

Public Property Get LineaCred() As String
    LineaCred = strLineaCred
End Property

Public Property Let LineaCred(ByVal vstrLineaCred As String)
    strLineaCred = vstrLineaCred
End Property

Public Property Get Consumo() As String
    Consumo = strConsumo
End Property

Public Property Let Consumo(ByVal vstrConsumo As String)
    strConsumo = vstrConsumo
End Property


'Public Property Let NumPedidoPadre(ByVal vstrNumPedidoPadre As String)
'    strNumPedidoPadre = vstrNumPedidoPadre
'End Property
'
'Public Property Get NumPedidoPadre() As String
'    NumPedidoPadre = strNumPedidoPadre
'End Property

Public Function CancelarVenta()
    frmPedido.psub_BeginArry
    frm_VTA_Documento.blnTipoDoc = False
    
    'subNuevo
    mdiPrincipal.ctlCliente1.Limpiar
    If mdiPrincipal.txtDireccion.Visible = True Then mdiPrincipal.txtDireccion.Text = ""
    frmPedido.lblTotal.Caption = "0.00": frmPedido.lblTotalPagar.Caption = "0.00":
    frmPedido.lblRedondeo.Caption = "0.00": frmPedido.lblPagado.Caption = "0.00":
    frmPedido.lblVuelto.Caption = "0.00"
    frmPedido.lblPctCopago.Caption = "0.00"
    frmPedido.LblCoPago.Caption = "0.00"
    frmPedido.pstrdxPrcCero = ""
    
    '-- Limpia Variables Publicas --'
    frmPedido.pstrCant = "": frmPedido.pstrPrecio = "": frmPedido.pstrProd = "": frmPedido.pstrSubTot = ""
    frmPedido.pstrCtdFracc = "": frmPedido.pstrCtdFracc = ""
    frmPedido.pstrPctUnit = "": frmPedido.pstrPrcUniKairo = ""
    frmPedido.pstrImpuesto = "": frmPedido.pstrComision = ""
    frmPedido.pstrPromocion = ""
    'frmPedido.psub_BeginArry
    frm_VTA_Busqueda.grdProductos.Limpiar
    frm_VTA_Busqueda.grdAlternativos.Limpiar
    frm_VTA_Busqueda.grdComplementarios.Limpiar
    frm_VTA_Busqueda.txtBuscar.selection
On Error GoTo g
   Dim i As Integer
   Dim intCountForm As Integer
   intCountForm = Forms.Count
   While i < intCountForm
    Dim LIMP As Boolean
    LIMP = True
    If Forms(i).name = "frmPedido" Then LIMP = False
    If Forms(i).name = "mdiPrincipal" Then LIMP = False
    If LIMP = True Then Unload (Forms(i)): i = 0: intCountForm = intCountForm - 1
    'Debug.Print "-->" & i & "<--||" & intCountForm & "||*****" & Forms(i).name & "******"
    i = i + 1
   Wend
g:
  
    mdiPrincipal.txtDLVTelefono.Text = ""
    frm_VTA_Busqueda.txtBuscar.Clear
    frm_VTA_Busqueda.SetFocus

    frm_VTA_Busqueda.grdProductos.Limpiar
    frm_VTA_Busqueda.grdAlternativos.Limpiar
    frm_VTA_Busqueda.grdComplementarios.Limpiar
    frm_VTA_Busqueda.txtBuscar.selection
    frm_VTA_Modalidad.Show vbModal
End Function

Public Property Get FlgValidaLinCre() As String
    FlgValidaLinCre = strFlgValidaLinCre
End Property

Public Property Let FlgValidaLinCre(ByVal lFlgValidaLinCre As String)
    strFlgValidaLinCre = lFlgValidaLinCre
End Property

Public Property Get FlgPlanVital() As String

    FlgPlanVital = m_sFlgPlanVital

End Property

Public Property Let FlgPlanVital(ByVal sFlgPlanVital As String)

    m_sFlgPlanVital = sFlgPlanVital

End Property

Public Property Get DevEfectivoNC() As String
    DevEfectivoNC = strDevEfectivoNC
End Property

Public Property Let DevEfectivoNC(ByVal lstrDevEfectivoNC As String)
    strDevEfectivoNC = lstrDevEfectivoNC
End Property

Public Property Get NomTitular() As String
    NomTitular = strNomTitular
End Property

Public Property Let NomTitular(ByVal lstrNomTitular As String)
    strNomTitular = lstrNomTitular
End Property

Public Property Get NumDNI() As String
    NumDNI = strNumDNI
End Property

Public Property Let NumDNI(ByVal lstrNumDNI As String)
    strNumDNI = lstrNumDNI
End Property


Public Property Get ProvPreDeterminadoRM() As String
    ProvPreDeterminadoRM = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_PROV_PRED_RM")
End Property

'-- Saca la serie de la Ticketera --'
Public Property Get SerieTKT() As String
    SerieTKT = strSerieTKT
End Property

Public Property Let SerieTKT(ByVal lstrSerieTKT As String)
    strSerieTKT = lstrSerieTKT
End Property

Public Property Get SloganBTL() As String
    SloganBTL = "BTL Calidad que Protege tu Salud"
End Property

Public Property Get Apostrofe() As String
    Apostrofe = "*"
End Property

'Public Property Let SloganBTL(ByVal lstrSloganBTL As String)
'    strSloganBTL = lstrSloganBTL
'End Property

Function AnulaTransaccionNavsat(ByVal NumeroTelefono As String, _
                                  ByVal codigoProducto As String, _
                                  ByVal MontoRecarga As Double, _
                                ByVal NumeroAprobacion As String, _
                                ByVal NumeroTraceOriginal As String) As Boolean

'objAccesoLib.AnulacionVentaDePines(
Dim rsRecargaVirtual As oraDynaset
Dim objServicio As New clsServicio
Dim objAccesoLib As New cls_AccesoLib
Set rsRecargaVirtual = objServicio.ListaRecargaX(codigoProducto)
On Error GoTo handle
'''
Dim strCodigoDeProveedor As String
Dim terminal  As String
Dim Serie As String
Dim Trace As String
Dim strIpHost  As String
Dim strPuertoHost As String
Dim strTimeOut  As String
Dim varConex  As String
Dim strCodigoComercio As String

terminal = Format(val(Mid(objUsuario.NombrePC, 8, 1)), "0000")
Serie = Format(val(objUsuario.CodigoLocal), "0000") + Format(val(Mid(objUsuario.NombrePC, 8, 1)), "0000")
Trace = Me.NumeroTrace(objUsuario.CodigoLocal)
strIpHost = Me.ParametroValor("IPHOST")
strPuertoHost = Me.ParametroValor("PUERTO")
strTimeOut = Me.ParametroValor("TIMEOUT")
strCodigoComercio = CodigoComercio(objUsuario.CodigoLocal)
MontoRecarga = MontoRecarga * 100 'enviar porque navsat lo requiere con dos ceros adicionales
'MontoRecarga = 10 * 100
 varConex = objAccesoLib.AnulacionVentaDePines(strCodigoComercio, NumeroTelefono, _
                                             rsRecargaVirtual("CODIGO"), rsRecargaVirtual("TIPOTARJETA"), _
                                             MontoRecarga, terminal, _
                                             Serie, Trace, NumeroAprobacion, NumeroTraceOriginal, _
                                             strIpHost, strPuertoHost, _
                                             CInt(strTimeOut))
Dim TramaEnvio As String
TramaEnvio = strCodigoComercio & "|" & NumeroTelefono & "|" & rsRecargaVirtual("CODIGO") & "|" & rsRecargaVirtual("TIPOTARJETA") & "|" & MontoRecarga & "|" & terminal & "|" & Serie & "|" & Trace & "|" & strIpHost & "|" & strPuertoHost & "|" & CInt(strTimeOut)
'objServicio.logRecarga varConex, objUsuario.Codigo, TramaEnvio, objUsuario.CodigoLiquidacion, MontoRecarga, NumeroTelefono, "servicio", objUsuario.CodigoEmpresa, "", "", SecuenciaGrabacion


If Not InStr(varConex, "//") > 0 Then
''''    GoTo handle
    
''''If Err.Description = "" Then
''''    Err.Number = varConex
''''    Err.Description = objServicio.ErrorNAVSAT(varConex)
''''Else
''''    'desError = Err.Description
''''End If
    
    Err.Raise varConex, "NAVSAT", objServicio.ErrorNAVSAT(varConex)
End If
    Debug.Print fncGuion(varConex, 0, "//") & "->Codigo Aprobacion"
    Debug.Print fncGuion(fncGuion(varConex, 1, "//"), 0, "//") & "->Folio"
    Debug.Print fncGuion(varConex, 0, "//") & "->Codigo Aprobacion"
    Debug.Print fncGuion(varConex, 0, "//") & "->Codigo Aprobacion"
AnulaTransaccionNavsat = True

Set objAccesoLib = Nothing
Set objServicio = Nothing
Exit Function
handle:
'Dim desError As String
AnulaTransaccionNavsat = False
Set objAccesoLib = Nothing
Set objServicio = Nothing
Err.Raise Err.Number, "clsParametro", Err.Description

End Function
'Esta funcion realiza transaccion x transaccion NAVSAT
Function ValidaTransaccionNavsat(ByVal NumeroTelefono As String, _
                                  ByVal codigoProducto As String, _
                                  ByVal MontoRecarga As Double, _
                                  ByRef NumeroAprobacion As String, _
                                  ByRef NumeroTraceOrigimal As String, _
                                  ByRef NumeroPin As String, _
                                  ByRef TramaEnvioOut As String, _
                                  ByRef TramaRespuestaOut As String, _
                                  ByVal Secuencia As String _
                                  ) As Boolean
Dim rsRecargaVirtual As oraDynaset
Dim objServicio As New clsServicio
Dim objAccesoLib As New cls_AccesoLib
Set rsRecargaVirtual = objServicio.ListaRecargaX(codigoProducto)
On Error GoTo handle
''''
Dim strCodigoDeProveedor As String
Dim terminal  As String
Dim Serie As String
Dim Trace As String
Dim strIpHost  As String
Dim strPuertoHost As String
Dim strTimeOut  As String
Dim varConex  As String
Dim strCodigoComercio As String

terminal = Format(val(Mid(objUsuario.NombrePC, 8, 1)), "0000")
Serie = Format(val(objUsuario.CodigoLocal), "0000") + Format(val(Mid(objUsuario.NombrePC, 8, 1)), "0000")
Trace = Me.NumeroTrace(objUsuario.CodigoLocal)
strIpHost = Me.ParametroValor("IPHOST")
strPuertoHost = Me.ParametroValor("PUERTO")
strTimeOut = Me.ParametroValor("TIMEOUT")
'++++++++++++++++++++++++++++++++++++++++++++++++++
'+ 09/07/2008 por JLOPEZ
'+ El código de comercio es los nueve primeros dígitos del ruc
'+ y los 3 dígitos del local
'++++++++++++++++++++++++++++++++++++++++++++++++++
'strCodigoComercio = CodigoComercio(objUsuario.CodigoLocal)
strCodigoComercio = left(objUsuario.Ruc, 9) & objUsuario.CodigoLocal
Dim TramaEnvio As String

MontoRecarga = MontoRecarga * 100 'enviar porque navsat lo requiere con dos ceros adicionales
TramaEnvio = strCodigoComercio & "|" & NumeroTelefono & "|" & rsRecargaVirtual("CODIGO") & "|" & rsRecargaVirtual("TIPOTARJETA") & "|" & MontoRecarga & "|" & terminal & "|" & Serie & "|" & Trace & "|" & strIpHost & "|" & strPuertoHost & "|" & CInt(strTimeOut)

objServicio.logRecarga "ANTES DE ENVIO", objUsuario.Codigo, TramaEnvio, objUsuario.CodigoLiquidacion, MontoRecarga / 100, NumeroTelefono, codigoProducto, objUsuario.CodigoEmpresa, "", "", "Comienza el envio", Secuencia

 varConex = objAccesoLib.VentaDePines(strCodigoComercio, NumeroTelefono, _
                                             rsRecargaVirtual("CODIGO"), rsRecargaVirtual("TIPOTARJETA"), _
                                             MontoRecarga, terminal, _
                                             Serie, Trace, _
                                             strIpHost, strPuertoHost, _
                                             CInt(strTimeOut))
 Debug.Print "Comercio=>" & strCodigoComercio & Chr(13) & "Telefono=>" & NumeroTelefono & Chr(13) & "Codigo=>" & rsRecargaVirtual("TIPOTARJETA") & Chr(13) & "Monto=>" & MontoRecarga & Chr(13) & "Terminal=>" & terminal & _
            Chr(13) & "Serie=>" & Serie & Chr(13) & "ip =>" & strIpHost & Chr(13) & "Puerto =>" & strPuertoHost & Chr(13) & "Timeout =>" & CInt(strTimeOut)
            

TramaEnvio = strCodigoComercio & "|" & NumeroTelefono & "|" & rsRecargaVirtual("CODIGO") & "|" & rsRecargaVirtual("TIPOTARJETA") & "|" & MontoRecarga & "|" & terminal & "|" & Serie & "|" & Trace & "|" & strIpHost & "|" & strPuertoHost & "|" & CInt(strTimeOut)
'varConex = "5691//064902//000003824901//337654073347"
' guardar el valor pin en la base de datos
objServicio.logRecarga varConex, objUsuario.Codigo, TramaEnvio, objUsuario.CodigoLiquidacion, MontoRecarga / 100, NumeroTelefono, codigoProducto, objUsuario.CodigoEmpresa, "", "", "Fin del envio", Secuencia
TramaEnvioOut = TramaEnvio
TramaRespuestaOut = varConex
Dim desError As String
If Not InStr(varConex, "//") > 0 Then
    desError = LCase(varConex) & "-" & "" & objServicio.ErrorNAVSAT(varConex)
    Err.Raise vbObjectError + 1, "Navsat", desError
End If
    NumeroAprobacion = fncGuion(Mid(varConex, 7), 0, "//")
    NumeroTraceOrigimal = Trace
    NumeroPin = fncGuion(fncGuion(fncGuion(fncGuion(varConex, 1, "//"), 1, "//"), 1, "//"), 1, "/")
    If NumeroPin = NumeroAprobacion Then
        NumeroPin = NumeroTelefono
    End If
 '   ValidaTransaccionNavsat = False
    ValidaTransaccionNavsat = True
Exit Function

Set objAccesoLib = Nothing
Set objServicio = Nothing

handle:
    ValidaTransaccionNavsat = False
    Set objAccesoLib = Nothing
    Set objServicio = Nothing
    
    'Err.Raise vbObjectError + 1, "clVenta.clsParametro", Err.Description
End Function

Public Function NumeroTrace(ByVal CodBtl As String) As String
Dim OraSqlStmt As OraSqlStmt

On Error GoTo ERROR
    
            NumeroTrace = gclsOracle.FN_Valor("BTLCADENA.PKG_TARJETAS_VIRTUALES.FN_DEV_SEQ", CodBtl)
            Exit Function
ERROR:
    Err.Raise Err.Number, "clsTarjetaVirtual.numeroTrace", Err.Description

End Function
Public Function CodigoComercio(ByVal CodBtl As String) As String
Dim OraSqlStmt As OraSqlStmt

On Error GoTo ERROR
    
            CodigoComercio = gclsOracle.FN_Valor("BTLCADENA.FN_DEV_VALOR_AUXILIAR", CodBtl, "CC")
            Exit Function
ERROR:
    Err.Raise Err.Number, "clsTarjetaVirtual.numeroTrace", Err.Description

End Function


Public Function ParametroValor(ByVal Param As Variant) As Variant
Dim OraSqlStmt As OraSqlStmt

    On Error GoTo ERROR
            ParametroValor = gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", Param)
    
    Exit Function
ERROR:
    Err.Raise Err.Number, "clsParametro", Err.Description
End Function


Public Property Get CodMotorizadoConv() As String
    CodMotorizadoConv = strCodMotorizadoConv
End Property

Public Property Let CodMotorizadoConv(ByVal lstrCodMotorizadoConv As String)
    strCodMotorizadoConv = lstrCodMotorizadoConv
End Property

Public Property Get CodFPagoDocDcto() As String
    CodFPagoDocDcto = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_FORMA_PAGO_DOC_PAGO")
End Property

Public Property Get TipoDocTKB() As String
    TipoDocTKB = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIP_DOC_TKT_BOL")
End Property

Public Function EsNavsat(CodigoServicio As String) As String
    EsNavsat = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO.FN_ES_NAVSAT", CodigoServicio)
End Function

Public Property Get TipoDocTKF() As String
    TipoDocTKF = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIP_DOC_TKT_FAC")
End Property


Public Function AgregaProductoLote(Codigo As String, _
                               Descripcion As String, _
                               cantidad As Integer, _
                               FlagFraccion As String, _
                               Precio As Double, _
                               TipoVenta As TipoVenta, _
                               Regalo As eProdRegalo, _
                               Optional Dato1 As String = "", _
                               Optional Dato2 As String = "", _
                               Optional Dato3 As String = "", _
                               Optional Dato4 As String = "", _
                               Optional Dato5 As String = "", _
                               Optional IndicadorReceta As String = "", _
                               Optional pctComision As Double = 0, _
                               Optional NroLote As String = "", _
                               Optional FchVmto As String = "", _
                               Optional flgFarmaco As String = "0", _
                               Optional PrecPublic As Double, Optional CadenaPromocion As String = "", Optional CadenaPromocionDcto As String = "", Optional FracEnabled As String = "") As XArrayDB

'                               Optional NroLote As String = "", _
'                               ) As XArrayDB
'
'                            Optional NroLote As String = "", _
'                               Optional FchVmto As String = "",
                               
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If xProductoLote.Count(1) < 0 Then Exit Function
    
    Dim i As Integer
    Dim encontro As Boolean
    Dim PctComi As Double
    Dim objProducto As New clsProducto
    Dim oldCant As Integer
    Dim oldFrac As String
    
    Dim canti As Double
    Dim prec As Double
    Dim cantidadFraccionable As Integer
    Dim flgFracEnabled As Integer
    
    '''''''''''''PctComi = objProducto.pctComision(Codigo, objUsuario.CodigoLocal, Format(ptmTipoPrecio, "000"))
    
    PctComi = pctComision
    
    aux = xProductoLote.Count(1)
    While i < aux
        'If xProducto(i, 0) = Codigo And xProducto(i, 5) = TipoVenta Then
        If xProductoLote(i, 0) = Codigo And xProductoLote(i, 7) = Regalo And xProductoLote(i, 22) = NroLote Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xProductoLote.Count(1) 'Llena con la ultima posicion disponible
        End If
        i = i + 1
    Wend

    If encontro = False Then
        xProductoLote.AppendRows
    End If
j:
    If xProductoLote.Count(1) = 0 Then ultimo = 0: xProductoLote.AppendRows
    
    
    If xProductoLote(i, 0) = Codigo And xProductoLote(i, 7) = Regalo And xProductoLote(i, 22) = NroLote Then
        If xProductoLote(ultimo, 14) = 0 Then
            oldCant = val(cantidad)
        Else
            oldCant = xProductoLote(ultimo, 3)
        End If
        If xProductoLote(ultimo, 15) = "" Then
            oldFrac = FlagFraccion
        Else
            oldFrac = IIf(xProductoLote(ultimo, 2) = "F", "1", "0")
        End If
    End If
    

    
    
    xProductoLote(ultimo, 0) = Codigo ' Asigna el codigo de producto
    xProductoLote(ultimo, 1) = Descripcion ' Asigna la descripcion del producto
    xProductoLote(ultimo, 2) = IIf(FlagFraccion = "1", "F", "U") ' Asigna si la cantidad es fraccion o Unidad
    xProductoLote(ultimo, 3) = val(cantidad) ' Asigna la cantidad seleccionada del producto
    xProductoLote(ultimo, 4) = Format(Precio, "###,##0.00") ' Asigna el tipo de Venta
    xProductoLote(ultimo, 5) = TipoVenta ' Asigna el tipo de Venta
    xProductoLote(ultimo, 6) = FlagFraccion ' Asigna si la cantidad es fraccion o Unidad
    xProductoLote(ultimo, 7) = Regalo 'valor libre
    xProductoLote(ultimo, 8) = ""
    xProductoLote(ultimo, 9) = 0
    xProductoLote(ultimo, 10) = ""
    xProductoLote(ultimo, 11) = ""
    xProductoLote(ultimo, 12) = ""
    xProductoLote(ultimo, 13) = PctComi
    If Regalo = Producto_Normal Then
        xProductoLote(ultimo, 14) = 0
        xProductoLote(ultimo, 15) = 0
    Else
        xProductoLote(ultimo, 14) = oldCant
        xProductoLote(ultimo, 15) = oldFrac
    End If
    xProductoLote(ultimo, 16) = Dato1 'Formulario de SOAT o Especie Valorada
    xProductoLote(ultimo, 17) = Dato2 'Número de placa o de referencia de la especie Valorada
    xProductoLote(ultimo, 18) = Dato3
    xProductoLote(ultimo, 19) = Dato4
    xProductoLote(ultimo, 20) = Dato5
    xProductoLote(ultimo, 21) = IndicadorReceta
    xProductoLote(ultimo, 22) = NroLote
    xProductoLote(ultimo, 23) = FchVmto
    
    xProductoLote(ultimo, 24) = flgFarmaco
    xProductoLote(ultimo, 25) = Format(PrecPublic, "###,##0.00") 'Precio Publico
    xProductoLote(ultimo, 26) = CadenaPromocion 'Cadena de Promociones
    xProductoLote(ultimo, 27) = CadenaPromocionDcto 'Cadena de descuento por promocion
    
    If FracEnabled = "1" Then
       flgFracEnabled = 1
    Else
       flgFracEnabled = 0
    End If
    
    i = 0
    cantidadFraccionable = objProducto.DevFraccionProducto(Codigo)
    While i < xProductoLote.Count(1)
        If xProductoLote(i, 0) = Codigo Then
            If xProductoLote(i, 2) = "F" Then
                canti = canti + xProductoLote(i, 3)
                prec = prec + xProductoLote(i, 4)
            Else
                canti = canti + (xProductoLote(i, 3) * cantidadFraccionable)
                prec = prec + xProductoLote(i, 4)
            End If
        End If
         
            If objProducto.EvaluaCtdVta(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, Codigo, canti, flgFracEnabled) = "0" Then
               vstrStock1 = objProducto.Stock(Codigo, objUsuario.CodigoLocal, "1")
               MsgBox "La cantidad solicitada es mayor a la existente" & Chr(13) & "Stock Actual: " & vstrStock1, vbExclamation + vbOKOnly, App.ProductName
               xProductoLote.DeleteRows (ultimo)
               If encontro = True Then
                  evaluaStockGuiaLote = 0
               Else
                  evaluaStockGuiaLote = 1
               End If
               Set AgregaProductoLote = xProductoLote
               Exit Function
            End If
        i = i + 1
    Wend
    
    evaluaStockGuiaLote = 0
    
    Set AgregaProductoLote = xProductoLote
    
End Function


Public Function AgregaProductoLoteDetalle(arr As XArrayDB, Codigo As String, FracEnabled As String) As XArrayDB

'
'
'    Dim ultimo As Integer ' declara variable contador
'    Dim ultimoDet As Integer
'    Dim ultimoCant As Integer
'    Dim aux As Integer
'    Dim aux2 As Integer
'    Dim aux3 As Integer
'    If arr.Count(1) < 0 Then Exit Function
'
'    Dim i As Integer
'    Dim j As Integer
'    Dim k As Integer
'    Dim encontro As Boolean
'    Dim PctComi As Double
'    Dim objProducto As New clsProducto
'    Dim oldCant As Integer
'    Dim oldFrac As String
'
'    Dim canti As Integer
'    Dim FraccionaEnabled As String
'    Dim cantidadFraccionable As Integer
'
''    PctComi = pctComision
'
'    aux = xProductoLoteDetalle.Count(1)
'    While i < aux
'        If xProductoLoteDetalle(i, 0) = Codigo Then 'And xProductoLote(i, 7) = Regalo And xProductoLote(i, 22) = NroLote Then
'            ultimo = i
'            encontro = True
'GoTo m
'        Else
'            encontro = False
'            ultimo = xProductoLoteDetalle.Count(1) 'Llena con la ultima posicion disponible
'        End If
'        i = i + 1
'    Wend
'
'    If encontro = False Then
'        xProductoLoteDetalle.AppendRows
'    End If
'm:
'    If xProductoLoteDetalle.Count(1) = 0 Then ultimo = 0: xProductoLoteDetalle.AppendRows
'
'
'    If xProductoLoteDetalle(i, 0) = Codigo Then
'
'    End If
'
'    aux2 = xProductoLote.Count(1)
'    While j < aux2
'        If xProductoLote(j, 0) = Codigo Then
'            ultimoDet = j
'            encontro = True
'GoTo n
'        Else
'            encontro = False
'            ultimoDet = xProductoLote.Count(1) 'Llena con la ultima posicion disponible
'        End If
'        j = j + 1
'    Wend
'n:
'
'cantidadFraccionable = objProducto.DevFraccionProducto(Codigo)
'    While k < aux2
'        If xProductoLote(k, 0) = Codigo Then
'            If xProductoLote(k, 2) = "F" Then
'                canti = canti + xProductoLote(k, 3)
'            Else
'                canti = canti + (xProductoLote(k, 3) * cantidadFraccionable)
'            End If
'        End If
'        k = k + 1
'    Wend
'
''
''    If canti Mod cantidadFraccionable = 0 Then
''        canti = canti / cantidadFraccionable
''        FraccionaEnabled = "U"
''    Else
''        canti = canti
''        If FracEnabled = "1" Then
''            FraccionaEnabled = "F"
''        Else
''            FraccionaEnabled = "U"
''        End If
''    End If
'
'        If FracEnabled = "1" Then
'            FraccionaEnabled = "F"
'        Else
'            FraccionaEnabled = "U"
'        End If
'
'
'    xProductoLoteDetalle(ultimo, 0) = Codigo ' Asigna el codigo de producto
'    xProductoLoteDetalle(ultimo, 1) = xProductoLote(ultimoDet, 1) ' Asigna la descripcion del producto
'    xProductoLoteDetalle(ultimo, 2) = FraccionaEnabled ' Asigna si la cantidad es fraccion o Unidad
'    xProductoLoteDetalle(ultimo, 3) = canti  ' Asigna la cantidad seleccionada del producto
'
'    Set AgregaProductoLoteDetalle = xProductoLoteDetalle
'
'
End Function


Public Function GrabarVentaFallida(ByVal Cia As String, _
                       ByVal CodigoLocal As String, _
                       ByVal CantidadSolicitada As Integer, _
                       ByVal CantidaStock As String, _
                       ByVal flgFraccion As String, _
                       ByVal usuario As String, _
                       ByVal CodigoModalidad As String, _
                       ByVal codigoProducto As String _
                       ) As String
    On Error GoTo Pase
    Dim arrValores As Variant
    Dim arrDireccion As Variant
    arrValores = Array(Cia, CodigoLocal, CantidadSolicitada, flgFraccion, usuario, CodigoModalidad, codigoProducto)
    arrDireccion = Array(entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada)
    GrabarVentaFallida = gclsOracle.SP("BTLPROD.PKG_PRE_VENTA.SP_GRABA_FALLIDA", arrValores, arrDireccion)
    If GrabarVentaFallida = "" Then Exit Function
        Err.Raise 20000, "clsVenta", GrabarVentaFallida
    Exit Function
Pase:
    Err.Raise Err.Number, "clsVenta", Err.Description
End Function

Public Property Get Beneficiario() As String
    Beneficiario = strBenef
End Property

Public Property Let Beneficiario(ByVal vstrBenef As String)
    strBenef = vstrBenef
End Property


Public Function HayFarmacos() As Boolean
Dim bolRetorna As Boolean
Dim i As Integer

    bolRetorna = False
    For i = 0 To Producto.UpperBound(1)
            If Producto(i, 24) = "1" Then
                bolRetorna = True
                Exit For
            End If
    Next i

    HayFarmacos = bolRetorna
    

End Function

Public Property Get Ubigeo() As String
    Ubigeo = strUbigeo
End Property

Public Property Let Ubigeo(ByVal lstrUbigeo As String)
    strUbigeo = lstrUbigeo
End Property



Function debeRegularizar() As Boolean
On Error GoTo handle
    'Ejecuta la consulta sql para llenar el recordset de detalle venta
    Dim rsDetVta  As New ADODB.Recordset
    With rsDetVta
'        .ActiveConnection = m_strConexionLocal 'conn
        .ActiveConnection = gstrConexion
'        .Source = "SELECT * FROM detalleventa.txt"
        .Source = App.Path & IIf(right(App.Path, 1) = "\", "", "\") & "detalleventa.xml"
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open
        .ActiveConnection = Nothing
    End With
    While Not rsDetVta.EOF
        debeRegularizar = True
        rsDetVta.MoveNext
        Set rsDetVta = Nothing
        Exit Function
    Wend
    Set rsDetVta = Nothing
    debeRegularizar = False
    Exit Function
handle:
    Set rsDetVta = Nothing
    debeRegularizar = False
End Function


Public Function ImpCopagoxRango(Codigo As String, Importe As Double) As Double
On Error GoTo ctrlErrCnv
    ImpCopagoxRango = gclsOracle.FN_Valor("BTLPROD.PKG_CONVENIO.FN_CALCULA_COPAGO", Codigo, Importe)
Exit Function
ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.ImpCopagoxRango", Err.Description
End Function


'** Propiedades / 30/12/2008 para el convenio **'

Public Property Get ImprimeImp() As String
   ImprimeImp = strFlgImprimeImp
End Property

Public Property Let ImprimeImp(ByVal vNewValue As String)
   strFlgImprimeImp = vNewValue
End Property

Public Property Get PrecioMenor() As String
    PrecioMenor = strFlgPrecioMenor
End Property

Public Property Let PrecioMenor(ByVal vNewValue As String)
    strFlgPrecioMenor = vNewValue
End Property

Public Property Get PrecioMenorDeducible() As Variant
    PrecioMenorDeducible = strFlgPrecioDeducible
End Property

Public Property Let PrecioMenorDeducible(ByVal vNewValue As Variant)
    strFlgPrecioDeducible = vNewValue
End Property

Public Property Get PrecPublic() As Double
    PrecPublic = dblPrecioPublico
End Property

Public Property Let PrecPublic(ByVal vNewValue As Double)
    dblPrecioPublico = vNewValue
End Property

Public Property Get NewPctBeneficario() As Double
    NewPctBeneficario = dblNewPctBenef
End Property

Public Property Let NewPctBeneficario(ByVal vNewValue As Double)
    dblNewPctBenef = vNewValue
End Property

Public Function ErrorFracciones(ModalidadVenta As String, ByRef Mensaje As String) As Boolean
ErrorFracciones = False
Dim i As Integer
Dim j As Integer
Dim objProducto As New clsProducto
j = xProducto.Count(1)
    While i < j
         If xProducto(i, 6) = 1 And Not xProducto(i, 6) = objProducto.ListaDevFracciona(xProducto(i, 0), objUsuario.CodigoLocal, ModalidadVenta) Then
                Mensaje = Mensaje & i + 1 & xProducto(i, 0) & " - " & xProducto(i, 1) & Chr(13)
         End If
         i = i + 1
    Wend
    If Not Mensaje = "" Then
          ErrorFracciones = True
          Exit Function
    End If
    ErrorFracciones = False
End Function
'*************************************************************************************************************************
' Autor     : Juan Arturo Escate Espichán
' Fecha     : 21/07/2009 8:28 AM
' Proposito : Valida la venta con los parametros de mayorista
'*************************************************************************************************************************
Public Function ValidaVentaMayorista(ByVal CodigoCliente As String, _
                                        ByVal CodigoLocal As String, _
                                        ByVal MontoVenta As String _
) As String

On Error GoTo CtrlErr

    ValidaVentaMayorista = "" & gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO.FN_VALIDA_VTAMAY", CodigoCliente, CodigoLocal, MontoVenta)

    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.ValidaVentaMayorista", Err.Description
End Function
'*************************************************************************************************************************

'*************************************************************************************************************************
' Autor     : Juan Arturo Escate Espichán
' Fecha     : 21/07/2009 8:26AM
' Proposito : Valida si la venta necesita que se ingrese lote y/o fecha de vencimiento
'*************************************************************************************************************************
Public Function ValidaLote(ByVal CodigoModalidad As String, _
                           ByVal CodigoConvenio As String, _
                           ByVal MontoVenta As String _
) As String

On Error GoTo CtrlErr

    ValidaLote = "" & gclsOracle.FN_Valor("BTLPROD.PKG_CONVENIO.FN_LOTE", CodigoModalidad, CodigoConvenio, MontoVenta)

    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.ValidaLote", Err.Description
End Function
'*************************************************************************************************************************


Public Function ValidaVentaLotes() As Boolean
    Dim strDebeValidar As String
    Dim rs As oraDynaset
    Dim PrimeraVez As Boolean
    PrimeraVez = True
    strDebeValidar = Me.ValidaLote(Me.CodModalidadVenta, Me.CodigoConvenio, Totales(0))
    If strDebeValidar = "0" Then
        ValidaVentaLotes = True
    Else
        Dim flagVencimiento As String
        Dim flagLote As String

        Set rs = Me.ListaVentaxLotes(Me.CodModalidadVenta, Me.CodigoConvenio, Totales(0))
        flagVencimiento = "" & rs("FLG_FCH_VENCIMIENTO")
        flagLote = "" & rs("FLG_NUM_LOTE")
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim i As Integer
        Dim debeLote As String
        Dim debefchVencimiento As String
       
       'prueba
        'If strCodigoTipoVenta = 10 Then
         'mlevano 10/10/2012
         'While i < Me.Producto.Count(1)
           
        ' If Me.CodModalidadVenta = "010" Then
'            While i < Me.ProductoLote.Count(1)
'                 debeLote = "0"
'                 debefchVencimiento = "0"
'                 If flagVencimiento = "1" And Replace(Replace(ProductoLote(i, 23), "_", ""), "/", "") = "" Then
'                ' If flagVencimiento = "1" And Replace(Replace(Producto(i, 23), "_", ""), "/", "") = "" Then
'                     debefchVencimiento = "1"
'                 End If
'                 If flagLote = "1" And ProductoLote(i, 22) = "" Then
'                 'If flagLote = "1" And Producto(i, 22) = "" Then
'                     debeLote = "1"
'                 End If
'                 If debefchVencimiento = "1" Or debeLote = "1" Then
'                     If PrimeraVez = True Then
'                         If MsgBox("Existen algunos productos sin datos lote o fecha de vencimiento" & Chr(13) & "¿Desea completarlos?", vbYesNo + vbQuestion, App.ProductName) = vbNo Then
'                             ValidaVentaLotes = False
'                             Exit Function
'                         End If
'                     End If
'                     PrimeraVez = False
'                     frmLoteFechVen.CargaDatos flagLote, flagVencimiento, CStr(i)
'                     'frm_VTA_CantidadProducto.subDatos flagLote, flagVencimiento, CStr(i)
'                 End If
'                 i = i + 1
'            Wend
        ' Else
             While i < Me.Producto.Count(1)
                 debeLote = "0"
                 debefchVencimiento = "0"
                 'If flagVencimiento = "1" And Replace(Replace(ProductoLote(i, 23), "_", ""), "/", "") = "" Then
                 If flagVencimiento = "1" And Replace(Replace(Producto(i, 23), "_", ""), "/", "") = "" Then
                     debefchVencimiento = "1"
                 End If
                 'If flagLote = "1" And ProductoLote(i, 22) = "" Then
                 If flagLote = "1" And Producto(i, 22) = "" Then
                     debeLote = "1"
                 End If
                 If debefchVencimiento = "1" Or debeLote = "1" Then
                     If PrimeraVez = True Then
                         If MsgBox("Existen algunos productos sin datos lote o fecha de vencimiento" & Chr(13) & "¿Desea completarlos?", vbYesNo + vbQuestion, App.ProductName) = vbNo Then
                             ValidaVentaLotes = False
                             Exit Function
                         End If
                     End If
                     PrimeraVez = False
                     frmLoteFechVen.CargaDatos flagLote, flagVencimiento, CStr(i)
                     'frm_VTA_CantidadProducto.subDatos flagLote, flagVencimiento, CStr(i)
                 End If
                 i = i + 1
            Wend
         'End If
        'End If
    End If
    ValidaVentaLotes = True
End Function
Public Function ListaVentaxLotes(ByVal CodigoModalidad As String, _
                           ByVal CodigoConvenio As String, _
                           ByVal MontoVenta As String _
) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaVentaxLotes = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_REG_LOTE", 0, CodigoModalidad, CodigoConvenio, MontoVenta)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsVenta", Err.Description
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Autor: Juan Arturo Escate Espichan
' Fecha: 26/10/2009
' Proposito:
Public Function ValidaTransaccionNew(ByVal Importe As String, _
                                     ByVal NumeroTelefonico As String, _
                                     ByVal NombreMaquina As String, _
                                     ByRef NumeroError As String, _
                                     ByRef MensajeError As String, _
                                     ByRef CodigoAutorizacion As String, _
                                     ByVal codigoProducto As String _
                                      )
On Error GoTo handle
        Dim objServicio As New clsServicio
        Dim venta As CanalVenta
        'aqui configuramos el objeto de venta
        If venta Is Nothing Then
            Set venta = New CanalVenta
        End If
        venta.host = Me.ParametroValor("IPHOSTBS") '"200.76.44.114"
        venta.PORT = Me.ParametroValor("PUERTOBS") '8599
        'creamos peticion
        Dim req As New Peticion
        req.ClaveCanal = Me.ParametroValor("CLAVECANAL") '"BTLPE"
        req.terminal = NombreMaquina '"TerminalPrueba01"
        Dim rsRecargaVirtual  As oraDynaset
        Set rsRecargaVirtual = objServicio.ListaRecargaX(codigoProducto)
        If Not rsRecargaVirtual.EOF Then
            req.Producto = IIf("" & Trim(rsRecargaVirtual("CODIGO")) = "", "A", "" & Trim(rsRecargaVirtual("CODIGO")))
        Else
            req.Producto = "A"
        End If
        
        req.Destino = NumeroTelefonico '"01991234567"
        req.Password = NombreMaquina 'Me.ParametroValor("PASSWORDBS") '"btlvb6test"
        'req.Password = "btlvb6test"
        req.SetMonto Importe '10
        'req.Monto = CDec(0.2)
        req.EsPasswordTerminal = Me.ParametroValor("FLGPASSBS") '0
        req.Fecha = objUsuario.sysdate
        'enviamos
        Dim resp As Respuesta
        Set resp = venta.venta(req)
        'desplegamos respuesta
        NumeroError = CStr(resp.CodigoRespuesta)
        MensajeError = resp.Mensaje
        CodigoAutorizacion = resp.Atributo("CONF_EXT")
        
        Set venta = Nothing
        Set resp = Nothing
        
        Exit Function
handle:
    NumeroError = Err.Number
    MensajeError = Err.Description
    CodigoAutorizacion = ""
    Set venta = Nothing
    Set resp = Nothing
End Function


Public Property Get MotivoNotaCredito() As String
    MotivoNotaCredito = strMotivoNotaCredito
End Property
Public Property Let MotivoNotaCredito(ByVal lMotivoNotaCredito As String)
    strMotivoNotaCredito = lMotivoNotaCredito
End Property

Public Property Get Observacion() As String
    Observacion = strObservacion
End Property
Public Property Let Observacion(ByVal lObservacion As String)
    strObservacion = lObservacion
End Property


Public Function EliminaProductoNew(Codigo As String, TipoVenta As TipoVenta) As XArrayDB
    Dim ultimo As Integer ' declara variable contador
    If xProducto.Count(1) <= 0 Then Exit Function
    Dim i As Integer
    Dim aux As Integer
    aux = xProducto.Count(1)
    While i < xProducto.Count(1)
        If xProducto(i, 0) = Codigo And xProducto(i, 5) = TipoVenta Then
            ultimo = i
        End If
        i = i + 1
    Wend
    ''esto es cuando es un regalo elimine la promociones de los productos
    If xProducto(ultimo, 7) <> "0" And xProducto(ultimo, 5) = TipoVenta Then
        ''primero busco si existe otro regalo con la misma promocion
        Dim Y As Integer
        Dim existe As Boolean
        existe = False
            While Y < aux
                If val(xProducto(ultimo, 3)) > 0 And xProducto(Y, 5) = TipoVenta And xProducto(Y, 0) <> Codigo And xProducto(Y, 7) <> "0" And xProducto(Y, 26) = xProducto(ultimo, 26) Then  '' si no es el mismo producto y si es un regalo  y es la misma venta y es el mismo codigo de promocion
                    existe = True
                End If
                Y = Y + 1
            Wend
            
            If existe = False Then
                Y = 0
                While Y < aux
                    xProducto(Y, 26) = Replace(xProducto(Y, 26), Replace(xProducto(ultimo, 26), ";", ""), "")
                    Y = Y + 1
                Wend
            End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''
    xProducto.DeleteRows (ultimo)
    Set EliminaProductoNew = xProducto
End Function

Public Function ListaIndicadores(ByVal CodigoLocal As String, ByVal CodigoUsuario As String, ByVal FechaEmision As String, ByVal MetaValePromedio As String, ByVal MetaNumCliente As String) As OracleInProcServer.oraDynaset
    On Error GoTo handle
    Set ListaIndicadores = gclsOracle.FN_Cursor("BTLPROD.PKG_REPO_GESTION.FN_AVANCE_VALE_PROMEDIO", 0, CodigoLocal, CodigoUsuario, FechaEmision, MetaValePromedio, MetaNumCliente)
    Exit Function
handle:
    Err.Raise Err.Number, "clsVenta.ListaIndicadores", Err.Description
End Function

Public Property Get PlanVitalArray() As XArrayDB
    Set PlanVitalArray = xArrayPlanVital 'Devuelve la variable de la clase en forma de XArray
End Property

Public Function Graba_Carga_CNV(ByVal oradb As OraDatabase, _
                                ByVal vstrNomArch As String, _
                                ByVal vstrCodUsuario As String, ByRef vstrMensaje As String, _
                                ByVal vstrFecAte As String, _
                                ByVal vstrNumAte As String, _
                                ByVal vstrNombBenef As String, _
                                ByVal vstrApePatBenef As String, _
                                ByVal vstrApeMatBenef As String, _
                                ByVal vstrCodCentral As String, _
                                ByVal vstrNumPedido As String, _
                                ByVal vstrCodPostal As String, _
                                ByVal vstrDirBenef As String, _
                                ByVal vstrDesReferencia As String, _
                                ByVal vstrFonoBenef As String, _
                                ByVal vstrSexoBenef As String, _
                                ByVal vdblCtdPedido As Double, _
                                ByVal vstrNumPoliza As String, _
                                ByVal vstrNumCertif As String, _
                                ByVal vstrCodParentesco As String, _
                                ByVal vstrCodPlanDeducible As String, _
                                ByVal vstrTipoFactura As String, _
                                ByVal vstrFecNacimiento As String, _
                                ByVal vstrCodInterno As String, _
                                ByVal vstrDesObs As String, _
                                ByVal vdblPctCoPago As Double, ByVal vstrCodCovenio As String, ByVal CMP As String, ByVal CIE10 As String, ByVal LocalOrigen As String) As String
                 
Dim i As Integer

On Error GoTo Control

byteNumElementosPlanVital = IIf(PlanVitalArray.UpperBound(1) < 0, 1, PlanVitalArray.UpperBound(1)) + 1

    For i = oradb.Parameters.Count - 1 To 0 Step -1
       oradb.Parameters.Remove i
    Next

oradb.Parameters.Add "A_NOM_ARCHIVO", vstrNomArch, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_USUARIO", vstrCodUsuario, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_DES_MENSAJE", "", ORAPARM_OUTPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_FEC_ATENCION", vstrFecAte, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_NUM_ATENCION", vstrNumAte, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_NOM_BENEFICIARIO", vstrNombBenef, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_APE_PAT_BENEFICIARIO", vstrApePatBenef, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_APE_MAT_BENEFICIARIO", vstrApeMatBenef, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_CENTRAL", vstrCodCentral, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_NUM_PEDIDO", vstrNumPedido, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_POSTAL", vstrCodPostal, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_DIR_BENEFICIARIO", vstrDirBenef, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_DES_REFERENCIA", vstrDesReferencia, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_FONO_BENEFICIARIO", vstrFonoBenef, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_SEXO_BENEFICIARIO", vstrSexoBenef, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_CTD_PEDIDO", vdblCtdPedido, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_NUM_POLIZA", vstrNumPedido, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_NUM_CERTIFICADO", vstrNumCertif, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_PARENTESCO", vstrCodParentesco, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_PLAN_DEDUCIBLE", vstrCodPlanDeducible, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_TIPO_FACTURACION", vstrTipoFactura, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_FEC_NACIMIENTO", vstrFecNacimiento, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_INTERNO", vstrCodInterno, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_DES_OBSERVACION", vstrDesObs, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_PCT_COPAGO", vdblPctCoPago, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_CONVENIO", vstrCodCovenio, ORAPARM_INPUT, ORATYPE_VARCHAR2

oradb.Parameters.Add "A_NUM_ARCHIVO", "", ORAPARM_OUTPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_CMP", CMP, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_CIE10", CIE10, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_LOCAL_ORIGEN", LocalOrigen, ORAPARM_INPUT, ORATYPE_VARCHAR2


oradb.Parameters.AddTable "A_CAD_FECHA", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosPlanVital, 200
oradb.Parameters.AddTable "A_CAD_COD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosPlanVital, 200
oradb.Parameters.AddTable "A_CAD_CTD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosPlanVital, 200

Dim FecAtencion As OracleInProcServer.OraParamArray
Dim codProducto As OracleInProcServer.OraParamArray
Dim CtdProducto As OracleInProcServer.OraParamArray

Set FecAtencion = oradb.Parameters("A_CAD_FECHA")
Set codProducto = oradb.Parameters("A_CAD_COD_PRODUCTO")
Set CtdProducto = oradb.Parameters("A_CAD_CTD_PRODUCTO")
    
    For i = 0 To PlanVitalArray.UpperBound(1)
        
        FecAtencion(i) = PlanVitalArray(i, 1)
        codProducto(i) = PlanVitalArray(i, 2)
        CtdProducto(i) = PlanVitalArray(i, 3)
        
    Next
    
    oradb.ExecuteSQL " BEGIN BTLPROD.PKG_CARGA_ARCHIVO_CNV.SP_GRABA(:A_NOM_ARCHIVO, :A_COD_USUARIO, :A_DES_MENSAJE, :A_FEC_ATENCION, :A_NUM_ATENCION, :A_NOM_BENEFICIARIO, :A_APE_PAT_BENEFICIARIO, :A_APE_MAT_BENEFICIARIO, :A_COD_CENTRAL, :A_NUM_PEDIDO, :A_COD_POSTAL, :A_DIR_BENEFICIARIO, :A_DES_REFERENCIA, :A_FONO_BENEFICIARIO, :A_SEXO_BENEFICIARIO, :A_CTD_PEDIDO, :A_NUM_POLIZA, :A_NUM_CERTIFICADO, :A_COD_PARENTESCO, :A_COD_PLAN_DEDUCIBLE, :A_TIPO_FACTURACION, :A_FEC_NACIMIENTO, :A_COD_INTERNO, :A_DES_OBSERVACION, :A_PCT_COPAGO, :A_CAD_FECHA, :A_CAD_COD_PRODUCTO, :A_CAD_CTD_PRODUCTO, :A_COD_CONVENIO, :A_NUM_ARCHIVO, :A_CMP,:A_CIE10,:A_COD_LOCAL_ORIGEN) ; END;"
    ArchivoVital = "" & Trim(oradb.Parameters("A_NUM_ARCHIVO").Value)
    vstrMensaje = "" & Trim(oradb.Parameters("A_DES_MENSAJE").Value)
Exit Function

Control:

    Err.Raise Err.Number, "clsRemito.Grabar", Err.Description

End Function

Public Property Get NomArchVital() As String

    NomArchVital = m_sNomArchVital

End Property

Public Property Let NomArchVital(ByVal sNomArchVital As String)

    m_sNomArchVital = sNomArchVital

End Property

Public Property Get FecAteVital() As String

    FecAteVital = m_sFecAteVital

End Property

Public Property Let FecAteVital(ByVal sFecAteVital As String)

    m_sFecAteVital = sFecAteVital

End Property

Public Property Get NumAteVital() As String

    NumAteVital = m_sNumAteVital

End Property

Public Property Let NumAteVital(ByVal sNumAteVital As String)

    m_sNumAteVital = sNumAteVital

End Property

Public Property Get CodCentralVital() As String

    CodCentralVital = m_sCodCentralVital

End Property

Public Property Let CodCentralVital(ByVal sCodCentralVital As String)

    m_sCodCentralVital = sCodCentralVital

End Property

Public Property Get NumPedidoVital() As String

    NumPedidoVital = m_sNumPedidoVital

End Property

Public Property Let NumPedidoVital(ByVal sNumPedidoVital As String)

    m_sNumPedidoVital = sNumPedidoVital

End Property

Public Property Get CodPostalVital() As String

    CodPostalVital = m_sCodPostalVital

End Property

Public Property Let CodPostalVital(ByVal sCodPostalVital As String)

    m_sCodPostalVital = sCodPostalVital

End Property

Public Property Get DesRefVital() As String

    DesRefVital = m_sDesRefVital

End Property

Public Property Let DesRefVital(ByVal sDesRefVital As String)

    m_sDesRefVital = sDesRefVital

End Property

Public Property Get FonoBenefVital() As String

    FonoBenefVital = m_sFonoBenefVital

End Property

Public Property Let FonoBenefVital(ByVal sFonoBenefVital As String)

    m_sFonoBenefVital = sFonoBenefVital

End Property

Public Property Get SexoBenefVital() As String

    SexoBenefVital = m_sSexoBenefVital

End Property

Public Property Let SexoBenefVital(ByVal sSexoBenefVital As String)

    m_sSexoBenefVital = sSexoBenefVital

End Property

Public Property Get CtdPedidoVital() As Double

    CtdPedidoVital = m_dCtdPedidoVital

End Property

Public Property Let CtdPedidoVital(ByVal dCtdPedidoVital As Double)

    m_dCtdPedidoVital = dCtdPedidoVital

End Property

Public Property Get NumPolizaVital() As String

    NumPolizaVital = m_sNumPolizaVital

End Property

Public Property Let NumPolizaVital(ByVal sNumPolizaVital As String)

    m_sNumPolizaVital = sNumPolizaVital

End Property

Public Property Get NumCertifVital() As String

    NumCertifVital = m_sNumCertifVital

End Property

Public Property Let NumCertifVital(ByVal sNumCertifVital As String)

    m_sNumCertifVital = sNumCertifVital

End Property

Public Property Get CodParentesco() As String

    CodParentesco = m_sCodParentesco

End Property

Public Property Let CodParentesco(ByVal sCodParentesco As String)

    m_sCodParentesco = sCodParentesco

End Property

Public Property Get CodDeducibleVital() As String

    CodDeducibleVital = m_sCodDeducibleVital

End Property

Public Property Let CodDeducibleVital(ByVal sCodDeducibleVital As String)

    m_sCodDeducibleVital = sCodDeducibleVital

End Property

Public Property Get TipoFacturaVital() As String

    TipoFacturaVital = m_sTipoFacturaVital

End Property

Public Property Let TipoFacturaVital(ByVal sTipoFacturaVital As String)

    m_sTipoFacturaVital = sTipoFacturaVital

End Property

Public Property Get FecNacVital() As String

    FecNacVital = m_sFecNacVital

End Property

Public Property Let FecNacVital(ByVal sFecNacVital As String)

    m_sFecNacVital = sFecNacVital

End Property

Public Property Get CodInternoVital() As String

    CodInternoVital = m_sCodInternoVital

End Property

Public Property Let CodInternoVital(ByVal sCodInternoVital As String)

    m_sCodInternoVital = sCodInternoVital

End Property

Public Property Get DesObsVital() As String

    DesObsVital = m_sDesObsVital

End Property

Public Property Let DesObsVital(ByVal sDesObsVital As String)

    m_sDesObsVital = sDesObsVital

End Property

Public Property Get PctCoPagoVital() As Double

    PctCoPagoVital = m_dPctCoPagoVital

End Property

Public Property Let PctCoPagoVital(ByVal dPctCoPagoVital As Double)

    m_dPctCoPagoVital = dPctCoPagoVital

End Property

Public Property Get NombBenef() As String

    NombBenef = m_sNombBenef

End Property

Public Property Let NombBenef(ByVal sNombBenef As String)

    m_sNombBenef = sNombBenef

End Property

Public Property Get ApePatBenef() As String

    ApePatBenef = m_sApePatBenef

End Property

Public Property Let ApePatBenef(ByVal sApePatBenef As String)

    m_sApePatBenef = sApePatBenef

End Property

Public Property Get ApeMatBenef() As String

    ApeMatBenef = m_sApeMatBenef

End Property

Public Property Let ApeMatBenef(ByVal sApeMatBenef As String)

    m_sApeMatBenef = sApeMatBenef

End Property

Public Property Get DirBenef() As String

    DirBenef = m_sDirBenef

End Property

Public Property Let DirBenef(ByVal sDirBenef As String)

    m_sDirBenef = sDirBenef

End Property

Public Property Get ArchivoVital() As String

    ArchivoVital = m_sArchivoVital

End Property

Public Property Let ArchivoVital(ByVal sArchivoVital As String)

    m_sArchivoVital = sArchivoVital

End Property

Public Function DevCabVital() As oraDynaset
    Set DevCabVital = gclsOracle.FN_Cursor("BTLPROD.PKG_CARGA_ARCHIVO_CNV.FN_LISTA_CAB", 0, objVenta.ArchivoVital)
End Function

Public Function DevDetVital(ByVal vstrCodLocalOrig As String) As oraDynaset
    Set DevDetVital = gclsOracle.FN_Cursor("BTLPROD.PKG_CARGA_ARCHIVO_CNV.FN_LISTA_DET", 0, objVenta.ArchivoVital, vstrCodLocalOrig)
End Function

Public Property Get PctBeneficiarioReal() As Double
    PctBeneficiarioReal = dblPctBeneficiarioReal
End Property

Public Property Let PctBeneficiarioReal(ByVal ldblPctBeneficiario As Double)
    dblPctBeneficiarioReal = ldblPctBeneficiario
End Property


Public Property Get CMPVital() As String
    CMPVital = lCMPVital
End Property

Public Property Let CMPVital(ByVal dPctCoPagoVital As String)
    lCMPVital = dPctCoPagoVital
End Property

Public Property Get CIE10Vital() As String
    CIE10Vital = lCIE10Vital
End Property

Public Property Let CIE10Vital(ByVal dPctCoPagoVital As String)
    lCIE10Vital = dPctCoPagoVital
End Property
Public Property Get LocalOrigenVital() As String
    LocalOrigenVital = lLocalOrigenVital
End Property

Public Property Let LocalOrigenVital(ByVal dPctCoPagoVital As String)
    lLocalOrigenVital = dPctCoPagoVital
End Property

Public Property Get CodigoConvenioVital() As String
    CodigoConvenioVital = lCodigoConvenioVital
End Property

Public Property Let CodigoConvenioVital(ByVal dPctCoPagoVital As String)
    lCodigoConvenioVital = dPctCoPagoVital
End Property

Public Property Get NumMaximoUnidades() As Integer
    NumMaximoUnidades = lNumMaximoUnidades
End Property

Public Property Let NumMaximoUnidades(ByVal dNumMaximoUnidades As Integer)
    lNumMaximoUnidades = dNumMaximoUnidades
End Property

'Autor: mlaguna
'Proposito: Reporte de ventas por producto - resumen
'Fecha: 09-04-2010

Public Function ReporteVentasProducto(ByVal strCia As String, ByVal strLocal As String, ByVal strProducto As String, strModalidad As String, ByVal FechaInicial As String, ByVal FechaFinal As String) As oraDynaset

On Error GoTo Control

   gclsOracle.Num_Intentos = 1
   Set ReporteVentasProducto = gclsOracle.FN_Cursor("btlprod.PKG_REPO_VENTA.FN_LISTA_VENTA_POR_PRODUCTO", 0, strCia, strLocal, strProducto, strModalidad, FechaInicial, FechaFinal)

Exit Function

Control:
    Err.Raise Err.Number, "clsVenta.ReporteVentasProducto", Err.Description
End Function

'Autor: mlaguna
'Proposito: Reporte de ventas por producto - detalle
'Fecha: 09-04-2010

Public Function ReporteVentasProductoDet(ByVal strCia As String, ByVal strLocal As String, ByVal strProducto As String, strModalidad As String, ByVal FechaInicial As String, ByVal FechaFinal As String, ByVal strVendedor As String) As oraDynaset

On Error GoTo Control

   gclsOracle.Num_Intentos = 1
   Set ReporteVentasProductoDet = gclsOracle.FN_Cursor("btlprod.PKG_REPO_VENTA.FN_LISTA_VENTA_POR_PRODUCTO_D", 0, strCia, strLocal, strProducto, strModalidad, FechaInicial, FechaFinal, strVendedor)

Exit Function

Control:
    Err.Raise Err.Number, "clsVenta.ReporteVentasProductoDet", Err.Description
End Function

Public Function fnSolicitaCliente(codigopromocion As String) As Integer
On Error GoTo Control
    Dim h As Variant
    Dim i As Integer
    Dim strSolCli As String
    h = Split(codigopromocion, ";")
    
    strSolCli = gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_CIA", "SOLICICLIE", "10")
    
    If strSolCli = "0" Then
        GoTo Control
    End If
    
    While i < UBound(h)
        If CStr(h(i)) = "" Then fnSolicitaCliente = "0"
        fnSolicitaCliente = gclsOracle.FN_Valor("btlprod.fn_dev_dato_cliente", CStr(h(i)))
        
        i = i + 1
    Wend
    Exit Function
Control:
    fnSolicitaCliente = "0"
    
End Function

Public Function fnListaFidelizado(ByVal CodigoDocuIdentidad As String, ByVal NumDocuIdentidad As String) As oraDynaset

On Error GoTo Control

   gclsOracle.Num_Intentos = 1
   Set fnListaFidelizado = gclsOracle.FN_Cursor("BTLPROD.PKG_FIDELIZADO.FN_LISTA_CLIENTE", 0, CodigoDocuIdentidad, NumDocuIdentidad)

Exit Function

Control:
    Err.Raise Err.Number, "clsVenta.fnListaFidelizado", Err.Description
End Function

Public Function fnGrabaFidelizado(ByVal CodigoCliente As String, _
                                  ByVal CodigoDocuIdentidad As String, _
                                  ByVal Num_Documento_ID As String, _
                                  ByVal NombreCliente As String, _
                                  ByVal ApellidoPaterno As String, _
                                  ByVal ApellidoMaterno As String, _
                                  ByVal CodigoEstadoCivil As String, _
                                  ByVal Email As String, _
                                  ByVal Cargo As String, _
                                  ByVal Sexo As String, _
                                  ByVal FCH_NACIMIENTO As String, _
                                  ByVal NumeroHijos As String, _
                                  ByVal Ubigeo As String, _
                                  ByVal COD_LOCAL As String, _
                                  ByVal COD_USUARIO As String, _
                                  ByVal Cod_Tarjeta As String) As String 'oraDynaset

On Error GoTo Control

   gclsOracle.Num_Intentos = 1
   fnGrabaFidelizado = gclsOracle.FN_Valor("BTLPROD.PKG_FIDELIZADO.SP_GRABA_CLIENTE", CodigoCliente, _
                                  CodigoDocuIdentidad, _
                                  Num_Documento_ID, _
                                  NombreCliente, _
                                  ApellidoPaterno, _
                                  ApellidoMaterno, _
                                  CodigoEstadoCivil, _
                                  Email, _
                                  Cargo, _
                                  Sexo, _
                                  FCH_NACIMIENTO, _
                                  NumeroHijos, _
                                  Ubigeo, _
                                  objUsuario.CodigoLocal, _
                                  objUsuario.Codigo, _
                                  Cod_Tarjeta)


'''objVenta.CodigoCliente

Exit Function

Control:
    Err.Raise Err.Number, "clsVenta.fnGrabaFidelizado", Err.Description
End Function

Public Function MuestraFidelizado(Cod_parametro As String) As Boolean
Dim ih As String
On Error GoTo CtrlErr
    ih = gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_CIA", Cod_parametro, "10")  '"ACTDIDELIZ"
        MuestraFidelizado = IIf(IIf(ih = "", 0, ih) = 0, False, True)
    Exit Function
    
CtrlErr:
    Err.Raise Err.Number, "clsUsuario.MuestraFidelizado", Err.Description
End Function

Public Function Saldo_Cliente(Cod_Cliente As String) As String
Dim ih As String
Dim Maximo As Double
On Error GoTo CtrlErr
    ih = gclsOracle.FN_Valor("BTLPROD.FN_SALDO_CLIENTE", Cod_Cliente)
    Maximo = gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_CIA", "MONMAXDCTO", "10")
    Saldo_Cliente = Maximo - ih
    Exit Function
    
CtrlErr:
        
    Err.Raise Err.Number, "clsUsuario.Saldo_Cliente", Err.Description
End Function


Public Function fnRedondeo(Monto As Double) As Double

If 0 < val(Monto - Round(Monto, 1)) Then
    fnRedondeo = -Round((Monto - Round(Monto, 1)), 2)
ElseIf 0 = (Monto - Round(Monto, 1)) Then
    fnRedondeo = 0
Else
    fnRedondeo = Round((Monto - Round(Monto, 1)), 2)
     'fnRedondeo = Round(0.05 + (Monto - Round(Monto, 1)), 2)
End If
End Function

Public Function fnDevuelveSecuencia(Cia As String, TipoDocumento As String, numDocumento As String) As String
On Error GoTo CtrlErr
    fnDevuelveSecuencia = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO_DSCTO.FN_DEV_SEV_VENTA", Cia, TipoDocumento, numDocumento)
    Exit Function
CtrlErr:
        
    Err.Raise Err.Number, "clsVenta.fnDevuelveSecuencia", Err.Description
End Function

Public Function NombreCuponera(ByVal CodMaquina As String) As String
On Error GoTo CtrlErr
    NombreCuponera = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO_DSCTO.FN_DEV_IMP_CUPON", CodMaquina)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsVenta.NombreCuponera", Err.Description
End Function

Public Function MensajeRemesa(ByVal CodigoUsuario As String, ByVal NombreMaquina As String) As String
On Error GoTo CtrlErr
    MensajeRemesa = "" & gclsOracle.FN_Valor("BTLPROD.FN_DEV_VAL_DEPOSITO", CodigoUsuario, NombreMaquina)
    Exit Function
CtrlErr:
        
    Err.Raise Err.Number, "clsVenta.MensajeRemesa", Err.Description
End Function

Public Function GrabaVentaPerdida(ByVal vstrCodLocal As String, _
                               ByVal vstrCodProducto As String, _
                               ByVal vstrUsuario As String)
   On Error GoTo CtrlErr
   
   Dim gvarValores As Variant
   Dim gvarIO  As Variant
   
   
   
   gvarValores = Array(vstrCodLocal, vstrCodProducto, vstrUsuario)
                          
   gvarIO = Array(entrada, entrada, entrada)
                  
   GrabaVentaPerdida = gclsOracle.SP("BTLPROD.PKG_DOCUMENTO_DSCTO.SP_GRABA_VTA_PERDIDA", _
                                    gvarValores, _
                                    gvarIO)
                  
   Exit Function
CtrlErr:
   Err.Raise Err.Number, "clsVenta.GrabaVentaPerdida", Err.Description
        
        
End Function


Function EsRegaloTodo() As Boolean
Dim Respuesta  As Boolean
Respuesta = True
Dim ValTot As Double
Dim o As Integer
While o < xProducto.Count(1)
ValTot = ValTot + val(xProducto(o, 5))
    If xProducto(o, 7) <> 1 Then
        Respuesta = False
     '   EsRegaloTodo = False
     '   Exit Function
     If ValTot = 0 Then
        Respuesta = True
        xProducto(o, 7) = "1"
     End If
    End If

o = o + 1
Wend
EsRegaloTodo = Respuesta
End Function


Public Function NombreMedico(ByVal CMP As String) As String
On Error GoTo CtrlErr
    NombreMedico = "" & gclsOracle.FN_Valor("BTLPROD.PKG_PRE_VENTA.FN_DEV_CMP", CMP)
    Exit Function
CtrlErr:
        
    Err.Raise Err.Number, "clsVenta.NombreMedico", Err.Description
End Function

Public Function NombreTermica(ByVal CodMaquina As String) As String
On Error GoTo CtrlErr
    NombreTermica = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO_DSCTO.FN_DEV_IMP_TERMICO", CodMaquina)
    Exit Function
CtrlErr:
        
    Err.Raise Err.Number, "clsVenta.NombreTermica", Err.Description
End Function


Public Function EsMFA(ByVal CodLocal As String) As Boolean
On Error GoTo CtrlErr
    EsMFA = IIf("" & gclsOracle.FN_Valor("btlprod.fn_dev_local_mf", CodLocal) = 1, True, False)
    Exit Function
CtrlErr:
        
    Err.Raise Err.Number, "clsVenta.EsMFA", Err.Description
End Function

Public Function EsImprCupon(ByVal CodLocal As String) As Boolean
On Error GoTo CtrlErr
    EsImprCupon = IIf("" & gclsOracle.FN_Valor("btlprod.FN_IMP_CUPON_LOCAL", CodLocal) = 1, True, False)
    Exit Function
CtrlErr:
        
    Err.Raise Err.Number, "clsVenta.EsImprCupon", Err.Description
End Function


Public Function getExisteDNI_RENIEC(ByVal aDni As String) As String
On Error GoTo CtrlErr
    getExisteDNI_RENIEC = "" & gclsOracle.FN_Valor("BTLPROD.PKG_FIDELIZADO.FN_EXISTE_DNI_RENIEC", aDni)
    Exit Function
CtrlErr:
        
    Err.Raise Err.Number, "Busqueda DNI RENIEC", Err.Description
End Function

Public Function ValidaTarjetaCMR(ByVal vstrNumTarj As String) As Integer
    On Error GoTo CtrlErr
    ValidaTarjetaCMR = gclsOracle.FN_Valor("BTLPROD.FN_LISTA_BIN_CMR", vstrNumTarj)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago.ValidaTarjetaCMR", Err.Description
End Function

' ********** Programa Monedero **********
' ********** DJARA 01/06/2015  **********
Public Function fnGetEstadoTarjetaMonedero(ByVal NumTarjeta As String) As String
    On Error GoTo Control

    gclsOracle.Num_Intentos = 1
    fnGetEstadoTarjetaMonedero = gclsOracle.FN_Valor("BTLPROD.PKG_MONEDERO_MF.FN_GET_ESTADO_TARJETA", _
                                                     NumTarjeta)
    
    Exit Function
Control:
    Err.Raise Err.Number, "clsVenta.fnGetEstadoTarjetaMonedero", Err.Description
End Function

Public Function fnListaAfiliadoMonedero(ByVal CodigoDocuIdentidad As String, _
                                        ByVal NumDocuIdentidad As String, _
                                        Optional ByVal NumTarjeta As String = "") As oraDynaset
    On Error GoTo Control

    gclsOracle.Num_Intentos = 1
    Set fnListaAfiliadoMonedero = gclsOracle.FN_Cursor("BTLPROD.PKG_MONEDERO_MF.FN_LISTA_CLIENTE", _
                                                       0, _
                                                       CodigoDocuIdentidad, _
                                                       NumDocuIdentidad, _
                                                       NumTarjeta)
    
    Exit Function
Control:
    Err.Raise Err.Number, "clsVenta.fnListaAfiliadoMonedero", Err.Description
End Function

Public Function fnListaAfiliadoMonederoOff(ByVal CodigoDocuIdentidad As String, _
                                           ByVal NumDocuIdentidad As String) As clsAfiliado
    Dim rs As oraDynaset
    Dim oa As clsAfiliado
    Dim i As Integer, vtarjetas() As String
    On Error GoTo Control

    Set rs = Me.fnListaAfiliadoMonedero(CodigoDocuIdentidad, NumDocuIdentidad)
    
    If rs.RecordCount > 0 Then
        Set oa = New clsAfiliado
        oa.Nombre = "" & rs("DES_NOM_CLIENTE").Value
        oa.ApParterno = "" & rs("DES_APE_CLIENTE").Value
        oa.ApMarterno = "" & rs("DES_APE2_CLIENTE").Value
        oa.TipoDni = "" & rs("COD_DOCUMENTO_IDENTIDAD").Value
        oa.DNI = "" & rs("NUM_DOCUMENTO_ID").Value
        oa.FechaNacimiento = "" & rs("FCH_NACIMIENTO").Value
        oa.Genero = IIf("" & rs("FLG_SEXO").Value = "1", "M", "F")
        oa.Email = "" & rs("DES_EMAIL").Value
        oa.Telefono = "" & rs("NUM_TEL_FIJO").Value
        oa.Celular = "" & rs("NUM_TEL_MOVIL").Value
        oa.TipoLugar = "" & rs("COD_SUFIJO_DIRECCION").Value
        oa.Direccion = "" & rs("DES_DIRECCION_SOCIAL").Value
        oa.Departamento = "" & rs("DEPARTAMENTO").Value
        oa.Provincia = "" & rs("PROVINCIA").Value
        oa.Distrito = "" & rs("DISTRITO").Value
        oa.TipoDireccion = "" & rs("COD_TIPO_DIRECCION").Value
        oa.Referencias = "" & rs("DES_REFERENCIA").Value
        If ("" & rs("TARJETAS").Value <> "") Then
            vtarjetas = Split("" & rs("TARJETAS").Value, ",")
            oa.Tarjetas = New XArrayDB
            oa.Tarjetas.ReDim LBound(vtarjetas), UBound(vtarjetas), 0, 0
            For i = LBound(vtarjetas) To UBound(vtarjetas)
                oa.Tarjetas.Value(i, 0) = vtarjetas(i)
            Next i
        End If
        Set fnListaAfiliadoMonederoOff = oa
    Else
        Set fnListaAfiliadoMonederoOff = Nothing
    End If
    Set rs = Nothing
    Exit Function
Control:
    Err.Raise Err.Number, "clsVenta.fnListaAfiliadoMonederoOff", Err.Description
End Function

Public Function fnListaCamposAfiliacionMonedero() As oraDynaset

On Error GoTo Control

   gclsOracle.Num_Intentos = 1
   Set fnListaCamposAfiliacionMonedero = gclsOracle.FN_Cursor("BTLPROD.PKG_MONEDERO_MF.FN_LISTA_CAMPOS", 0)

Exit Function

Control:
    Err.Raise Err.Number, "clsVenta.fnListaCamposAfiliacionMonedero", Err.Description
End Function

Public Function fnGrabaAfiliadoMonedero(ByVal CodigoCliente As String, _
                                        ByVal CodigoDocuIdentidad As String, _
                                        ByVal Num_Documento_ID As String, _
                                        ByVal NombreCliente As String, _
                                        ByVal ApellidoPaterno As String, _
                                        ByVal ApellidoMaterno As String, _
                                        ByVal Email As String, _
                                        ByVal Sexo As String, _
                                        ByVal FechaNacimiento As String, _
                                        ByVal Ubigeo As String, _
                                        ByVal NumeroTarjeta As String, _
                                        Optional ByVal Telefono As String = "", _
                                        Optional ByVal Celular As String = "", _
                                        Optional ByVal DirecSufijo As String = "", _
                                        Optional ByVal Direccion As String = "", _
                                        Optional ByVal DirecTipo As String = "", _
                                        Optional ByVal DirecRefer As String = "", _
                                        Optional ByVal IndEnvOrbis As String = "", _
                                        Optional ByVal IndTrjOrbis As String = "N") As String 'oraDynaset
    On Error GoTo Control

    gclsOracle.Num_Intentos = 1
    fnGrabaAfiliadoMonedero = gclsOracle.FN_Valor("BTLPROD.PKG_MONEDERO_MF.SP_GRABA_CLIENTE", _
                                                  CodigoCliente, _
                                                  CodigoDocuIdentidad, _
                                                  Num_Documento_ID, _
                                                  NombreCliente, _
                                                  ApellidoPaterno, _
                                                  ApellidoMaterno, _
                                                  "", _
                                                  Email, _
                                                  "", _
                                                  Sexo, _
                                                  FechaNacimiento, _
                                                  0, _
                                                  Ubigeo, _
                                                  objUsuario.CodigoLocal, _
                                                  objUsuario.Codigo, _
                                                  NumeroTarjeta, _
                                                  Telefono, _
                                                  Celular, _
                                                  DirecSufijo, _
                                                  Direccion, _
                                                  DirecTipo, _
                                                  DirecRefer, _
                                                  IndEnvOrbis, _
                                                  IndTrjOrbis)
    Exit Function
Control:
    Err.Raise Err.Number, "clsVenta.fnGrabaAfiliadoMonedero", Err.Description
End Function

Public Function RecuperarPuntosMonedero(ByVal Cia As String, _
                                        ByVal TipoDocumento As String, _
                                        ByVal NumeroDocumento As String, _
                                        ByVal CodigoLocal As String, _
                                        ByVal usuario As String, _
                                        ByVal NumeroTarjeta As String, _
                                        ByVal SaldoActual As String, _
                                        ByVal EnviadoAOrbis As String _
                                        ) As String
    On Error GoTo Pase
    Dim arrValores As Variant
    Dim arrDireccion As Variant
    Dim dPuntos As Double
    Dim res As String
    
    arrValores = Array(dPuntos, Cia, TipoDocumento, NumeroDocumento, CodigoLocal, usuario, NumeroTarjeta, SaldoActual, EnviadoAOrbis)
    arrDireccion = Array(Salida, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada)
    
    res = gclsOracle.SP("BTLPROD.PKG_DOCUMENTO.SP_RECUPERA_PUNTOS_MONEDERO", arrValores, arrDireccion)
    
    If res = "" Then
        RecuperarPuntosMonedero = arrValores(0)
        Exit Function
    End If
    
    Err.Raise 20000, "clsVenta", res
    Exit Function

Pase:
    Err.Raise Err.Number, "clsVenta", Err.Description
End Function

Public Property Get MontoRedime() As String
    MontoRedime = lMontoRedime
End Property

Public Property Let MontoRedime(ByVal dMontoRedime As String)
    lMontoRedime = dMontoRedime
End Property

Public Property Let EsVentaMonedero(ByVal vData As Boolean)
    bEsVentaMonedero = vData
End Property

Public Property Get EsVentaMonedero() As Boolean
    EsVentaMonedero = IIf(bEsVentaMonedero, True, False)
End Property

Public Property Let NumeroTarjetaMonedero(ByVal vData As String)
    sNumeroTarjetaMonedero = vData
End Property

Public Property Get NumeroTarjetaMonedero() As String
    NumeroTarjetaMonedero = "" & sNumeroTarjetaMonedero
End Property

Public Property Let PuntosTarjetaMonedero(ByVal vData As Double)
    mvarPuntosTarjetaMonedero = vData
End Property

Public Property Get PuntosTarjetaMonedero() As Double
    PuntosTarjetaMonedero = mvarPuntosTarjetaMonedero
End Property

Public Property Let EscaneoTarjeta(ByVal vData As String)
    sEscaneoTarjeta = vData
End Property

Public Property Get EscaneoTarjeta() As String
    EscaneoTarjeta = "" & sEscaneoTarjeta
End Property


Public Property Let PuntoInicial(ByVal vData As String)
    sPuntoInicial = vData
End Property

Public Property Get PuntoInicial() As String
    PuntoInicial = "" & sPuntoInicial
End Property


Public Property Let NumTarjetaPuntos(ByVal vData As String)
    sNumTarjetaPuntos = vData
End Property

Public Property Get NumTarjetaPuntos() As String
    NumTarjetaPuntos = "" & sNumTarjetaPuntos
End Property

' ********** Fin Programa Monedero **********

''''''
''''''Public Function agregaRegalo(Codigo As String, _
''''''                               Descripcion As String, _
''''''                               cantidad As Integer, _
''''''                               FlagFraccion As String, _
''''''                               Precio As Double, _
''''''                               TipoVenta As TipoVenta, _
''''''                               Regalo As eProdRegalo, _
''''''                               Optional Dato1 As String = "", _
''''''                               Optional Dato2 As String = "", _
''''''                               Optional Dato3 As String = "", _
''''''                               Optional Dato4 As String = "", _
''''''                               Optional Dato5 As String = "", _
''''''                               Optional IndicadorReceta As String = "", _
''''''                               Optional pctComision As Double = 0, _
''''''                               Optional NroLote As String = "", _
''''''                               Optional FchVmto As String = "", _
''''''                               Optional flgFarmaco As String = "0", _
''''''                               Optional PrecPublic As Double, Optional CadenaPromocion As String = "", Optional CadenaPromocionDcto As String = "", Optional FracEnabled As String = "", Optional CantPuntos As String, Optional CantPuntosRedimir As Double) As XArrayDB
''''''
'''''''xProductoRegalo.AppendRows
''''''Dim ultimo As Integer
''''''Dim i As Integer
''''''Dim encontro As Boolean
''''''
''''''ultimo = xProductoRegalo.Count(1) 'Llena con la ultima posicion disponible
''''''If CUANTOTENGOREGALO(CadenaPromocion) - cantidad = 0 Then
''''''
''''''  AgregaProducto xProductoRegaloBK(i, 0), _
''''''   xProductoRegaloBK(i, 1), _
''''''    xProductoRegaloBK(i, 2), _
''''''    xProductoRegaloBK(i, 3), _
''''''    xProductoRegaloBK(i, 4), _
''''''    xProductoRegaloBK(i, 5), _
''''''    xProductoRegaloBK(i, 6), _
''''''    xProductoRegaloBK(i, 7), _
''''''    xProductoRegaloBK(i, 8), _
''''''    xProductoRegaloBK(i, 9), _
''''''    xProductoRegaloBK(i, 10), _
''''''    xProductoRegaloBK(i, 11), _
''''''    xProductoRegaloBK(i, 12), _
''''''    xProductoRegaloBK(i, 13), _
''''''    xProductoRegaloBK(i, 14), _
''''''    xProductoRegaloBK(i, 15), _
''''''    xProductoRegaloBK(i, 16), _
''''''    xProductoRegaloBK(i, 17), _
''''''    xProductoRegaloBK(i, 18), _
''''''    xProductoRegaloBK(i, 19), _
''''''    xProductoRegaloBK(i, 20), _
''''''    xProductoRegaloBK(i, 21), _
''''''    xProductoRegaloBK(i, 22)
''''''
''''''    Cumplio = True
''''''    Else
''''''    Cumplio = False
''''''End If
''''''While i < ultimo
''''''  If xProductoRegalo(i, 0) = Codigo And xProductoRegalo(i, 6) = Regalo Then
'''''''  Dim ultimo2 As Integer
'''''' ' Dim i2 As Integer
'''''' ' ultimo2 = xProducto.Count(1) 'Llena con la ultima posicion disponible
'''''''  While i2 < ultimo2
'''''''    If xProducto(i2, 0) = Codigo And xProducto(i2, 7) = Regalo And xProducto(i2, 26) = CadenaPromocion Then
''''''
''''''    If xProductoRegalo(i, 2) = cantidad Then
''''''
''''''    AgregaProducto xProductoRegalo(i, 0), _
''''''   xProductoRegalo(i, 1), _
''''''    xProductoRegalo(i, 2), _
''''''    xProductoRegalo(i, 3), _
''''''    xProductoRegalo(i, 4), _
''''''    xProductoRegalo(i, 5), _
''''''    xProductoRegalo(i, 6), _
''''''    xProductoRegalo(i, 7), _
''''''    xProductoRegalo(i, 8), _
''''''    xProductoRegalo(i, 9), _
''''''    xProductoRegalo(i, 10), _
''''''    xProductoRegalo(i, 11), _
''''''    xProductoRegalo(i, 12), _
''''''    xProductoRegalo(i, 13), _
''''''    xProductoRegalo(i, 14), _
''''''    xProductoRegalo(i, 15), _
''''''    xProductoRegalo(i, 16), _
''''''    xProductoRegalo(i, 17), _
''''''    xProductoRegalo(i, 18), _
''''''    xProductoRegalo(i, 19), _
''''''    xProductoRegalo(i, 20), _
''''''    xProductoRegalo(i, 21), _
''''''    xProductoRegalo(i, 22)
''''''
''''''
''''''    Dim JU As Integer
''''''    JU = xProductoRegaloBK.Count(1)
''''''    xProductoRegaloBK.AppendRows
''''''    xProductoRegaloBK(JU, 0) = xProductoRegalo(i, 0)
''''''    xProductoRegaloBK(JU, 1) = xProductoRegalo(i, 1)
''''''    xProductoRegaloBK(JU, 2) = xProductoRegalo(i, 2)
''''''    xProductoRegaloBK(JU, 3) = xProductoRegalo(i, 3)
''''''    xProductoRegaloBK(JU, 4) = xProductoRegalo(i, 4)
''''''    xProductoRegaloBK(JU, 5) = xProductoRegalo(i, 5)
''''''    xProductoRegaloBK(JU, 6) = xProductoRegalo(i, 6)
''''''    xProductoRegaloBK(JU, 7) = xProductoRegalo(i, 7)
''''''    xProductoRegaloBK(JU, 8) = xProductoRegalo(i, 8)
''''''    xProductoRegaloBK(JU, 9) = xProductoRegalo(i, 9)
''''''    xProductoRegaloBK(JU, 10) = xProductoRegalo(i, 10)
''''''    xProductoRegaloBK(JU, 11) = xProductoRegalo(i, 11)
''''''    xProductoRegaloBK(JU, 12) = xProductoRegalo(i, 12)
''''''    xProductoRegaloBK(JU, 13) = xProductoRegalo(i, 13)
''''''    xProductoRegaloBK(JU, 14) = xProductoRegalo(i, 14)
''''''    xProductoRegaloBK(JU, 15) = xProductoRegalo(i, 15)
''''''    xProductoRegaloBK(JU, 16) = xProductoRegalo(i, 16)
''''''    xProductoRegaloBK(JU, 17) = xProductoRegalo(i, 17)
''''''    xProductoRegaloBK(JU, 18) = xProductoRegalo(i, 18)
''''''    xProductoRegaloBK(JU, 19) = xProductoRegalo(i, 19)
''''''    xProductoRegaloBK(JU, 20) = xProductoRegalo(i, 20)
''''''    xProductoRegaloBK(JU, 21) = xProductoRegalo(i, 21)
''''''    xProductoRegaloBK(JU, 22) = xProductoRegalo(i, 22)
''''''    xProductoRegaloBK(JU, 23) = xProductoRegalo(i, 23)
''''''    xProductoRegaloBK(JU, 24) = xProductoRegalo(i, 24)
''''''
''''''        xProductoRegalo.DeleteRows (i)
''''''        ultimo = xProductoRegalo.Count(1)
''''''        Cumplio = True
''''''        Else
''''''        If Cumplio = True Then
''''''            xProductoRegalo.DeleteRows (i)
''''''            ultimo = xProductoRegalo.Count(1)
''''''        Else
''''''
''''''        End If
''''''        End If
''''''    'End If
'''''' '   i2 = i2 + 1
'''''''  Wend
''''''    encontro = True
''''''    Else
''''''    'encontro = False
''''''  End If
''''''i = i + 1
''''''Wend
''''''If encontro = False And Cumplio = False Then
''''''    xProductoRegalo.AppendRows
''''''    ultimo = xProductoRegalo.Count(1) - 1 'Llena con la ultima posicion disponible
''''''    xProductoRegalo(ultimo, 0) = Codigo
''''''    xProductoRegalo(ultimo, 1) = Descripcion
''''''    xProductoRegalo(ultimo, 2) = 0
''''''    xProductoRegalo(ultimo, 3) = FlagFraccion
''''''    xProductoRegalo(ultimo, 4) = Precio
''''''    xProductoRegalo(ultimo, 5) = TipoVenta
''''''    xProductoRegalo(ultimo, 6) = Regalo
''''''    xProductoRegalo(ultimo, 7) = Dato1
''''''    xProductoRegalo(ultimo, 8) = Dato2
''''''    xProductoRegalo(ultimo, 9) = Dato3
''''''    xProductoRegalo(ultimo, 10) = Dato4
''''''    xProductoRegalo(ultimo, 11) = Dato5
''''''    xProductoRegalo(ultimo, 12) = IndicadorReceta
''''''    xProductoRegalo(ultimo, 13) = pctComision
''''''    xProductoRegalo(ultimo, 14) = NroLote
''''''    xProductoRegalo(ultimo, 15) = FchVmto
''''''    xProductoRegalo(ultimo, 16) = flgFarmaco
''''''    xProductoRegalo(ultimo, 17) = PrecPublic
''''''    xProductoRegalo(ultimo, 18) = CadenaPromocion
''''''    xProductoRegalo(ultimo, 19) = CadenaPromocionDcto
''''''    xProductoRegalo(ultimo, 20) = FracEnabled
''''''    xProductoRegalo(ultimo, 21) = CantPuntos
''''''    xProductoRegalo(ultimo, 22) = CantPuntosRedimir
''''''    xProductoRegalo(ultimo, 23) = cantidad
''''''    If ultimo = 0 Then
''''''        xProductoRegalo(ultimo, 24) = cantidad
''''''        Else
''''''        If xProductoRegalo(ultimo - 1, 23) < cantidad Then
''''''        xProductoRegalo(ultimo, 24) = cantidad
''''''        Else
''''''        xProductoRegalo(ultimo, 24) = xProductoRegalo(ultimo - 1, 24)
''''''        End If
''''''    End If
''''''
''''''End If
''''''
''''''End Function

Function CUANTOTENGOREGALO(strPromocion As String)
Dim i As Integer
Dim Falta As Integer

While i < xProductoRegaloBK.Count(1)
    If xProductoRegaloBK(i, 18) = strPromocion And xProductoRegaloBK(i, 2) <> 0 Then
        Falta = Falta + val(xProductoRegaloBK(i, 2))
    End If
    i = i + 1
Wend
CUANTOTENGOREGALO = Falta

End Function
Function CUANTOREGALO(strPromocion As String)
Dim i As Integer
Dim Falta As Integer

While i < xProductoRegalo.Count(1)
    If xProductoRegalo(i, 18) = strPromocion And xProductoRegalo(i, 23) <> 0 Then
        Falta = Falta + val(xProductoRegalo(i, 23))
    End If
    i = i + 1
Wend
CUANTOREGALO = Falta

End Function


Public Function agregaRegalo(Codigo As String, _
                               Descripcion As String, _
                               cantidad As Integer, _
                               FlagFraccion As String, _
                               Precio As Double, _
                               TipoVenta As TipoVenta, _
                               Regalo As eProdRegalo, _
                               Optional Dato1 As String = "", _
                               Optional Dato2 As String = "", _
                               Optional Dato3 As String = "", _
                               Optional Dato4 As String = "", _
                               Optional Dato5 As String = "", _
                               Optional IndicadorReceta As String = "", _
                               Optional pctComision As Double = 0, _
                               Optional NroLote As String = "", _
                               Optional FchVmto As String = "", _
                               Optional flgFarmaco As String = "0", _
                               Optional PrecPublic As Double, Optional CadenaPromocion As String = "", Optional CadenaPromocionDcto As String = "", Optional FracEnabled As String = "", Optional CantPuntos As String, Optional CantPuntosRedimir As Double) As XArrayDB

'xProductoRegalo.AppendRows
Dim ultimo As Integer
Dim i As Integer
Dim encontro As Boolean
Dim esteen As Integer
ultimo = xProductoRegalo.Count(1) 'Llena con la ultima posicion disponible
encontro = False
While i < ultimo
If encontro = False Then
  If xProductoRegalo(i, 0) = Codigo And xProductoRegalo(i, 6) = Regalo Then
    esteen = i
    encontro = True
    
    
    Else
    'Exit Function
    encontro = False
    esteen = 0
  End If
  End If
i = i + 1
Wend
If encontro = False Then
    xProductoRegalo.AppendRows
    ultimo = xProductoRegalo.Count(1) - 1 'Llena con la ultima posicion disponible
    xProductoRegalo(ultimo, 0) = Codigo
    xProductoRegalo(ultimo, 1) = Descripcion
    xProductoRegalo(ultimo, 2) = 0
    xProductoRegalo(ultimo, 3) = FlagFraccion
    xProductoRegalo(ultimo, 4) = Precio
    xProductoRegalo(ultimo, 5) = TipoVenta
    xProductoRegalo(ultimo, 6) = Regalo
    xProductoRegalo(ultimo, 7) = Dato1
    xProductoRegalo(ultimo, 8) = Dato2
    xProductoRegalo(ultimo, 9) = Dato3
    xProductoRegalo(ultimo, 10) = Dato4
    xProductoRegalo(ultimo, 11) = Dato5
    xProductoRegalo(ultimo, 12) = IndicadorReceta
    xProductoRegalo(ultimo, 13) = pctComision
    xProductoRegalo(ultimo, 14) = NroLote
    xProductoRegalo(ultimo, 15) = FchVmto
    xProductoRegalo(ultimo, 16) = flgFarmaco
    xProductoRegalo(ultimo, 17) = PrecPublic
    xProductoRegalo(ultimo, 18) = CadenaPromocion
    xProductoRegalo(ultimo, 19) = CadenaPromocionDcto
    xProductoRegalo(ultimo, 20) = FracEnabled
    xProductoRegalo(ultimo, 21) = CantPuntos
    xProductoRegalo(ultimo, 22) = CantPuntosRedimir
    xProductoRegalo(ultimo, 23) = cantidad
    If ultimo = 0 Then
        xProductoRegalo(ultimo, 24) = cantidad
        Else
        If xProductoRegalo(ultimo - 1, 23) < cantidad Then
        xProductoRegalo(ultimo, 24) = cantidad
        Else
        xProductoRegalo(ultimo, 24) = xProductoRegalo(ultimo - 1, 24)
        End If
    End If
    'xProductoRegalo(ultimo, 25) = Precio / cantidad
    xProductoRegalo(ultimo, 25) = Precio / IIf(cantidad = 0, 1, cantidad) 'ECASTILLO 04.07.2020
Else
    xProductoRegalo(esteen, 0) = Codigo
    xProductoRegalo(esteen, 1) = Descripcion
    xProductoRegalo(esteen, 2) = cantidad
    xProductoRegalo(esteen, 3) = FlagFraccion
    xProductoRegalo(esteen, 4) = Precio
    xProductoRegalo(esteen, 5) = TipoVenta
    xProductoRegalo(esteen, 6) = Regalo
    xProductoRegalo(esteen, 7) = Dato1
    xProductoRegalo(esteen, 8) = Dato2
    xProductoRegalo(esteen, 9) = Dato3
    xProductoRegalo(esteen, 10) = Dato4
    xProductoRegalo(esteen, 11) = Dato5
    xProductoRegalo(esteen, 12) = IndicadorReceta
    xProductoRegalo(esteen, 13) = pctComision
    xProductoRegalo(esteen, 14) = NroLote
    xProductoRegalo(esteen, 15) = FchVmto
    xProductoRegalo(esteen, 16) = flgFarmaco
    xProductoRegalo(esteen, 17) = PrecPublic
    xProductoRegalo(esteen, 18) = CadenaPromocion
    xProductoRegalo(esteen, 19) = CadenaPromocionDcto
    xProductoRegalo(esteen, 20) = FracEnabled
    xProductoRegalo(esteen, 21) = CantPuntos
    xProductoRegalo(esteen, 22) = CantPuntosRedimir
    xProductoRegalo(esteen, 23) = cantidad
    If esteen = 0 Then
        xProductoRegalo(esteen, 24) = cantidad
        Else
        
        If xProductoRegalo(esteen - 1, 23) < cantidad Then
        xProductoRegalo(esteen, 24) = cantidad
        Else
        xProductoRegalo(esteen, 24) = xProductoRegalo(ultimo - 1, 24)
        End If
    End If
    'xProductoRegalo(esteen, 25) = Precio / cantidad
    xProductoRegalo(esteen, 25) = Precio / IIf(cantidad = 0, 1, cantidad) 'ECASTILLO 04.07.2020
End If
CUANTAMAXIMA = IIf(cantidad > CUANTAMAXIMA, cantidad, CUANTAMAXIMA)
AgregaCuentaMaxima CadenaPromocion, cantidad 'CUANTAMAXIMA 'ECASTILLO 04.07.2020
End Function
Function ObtieneCuentaMaxima(strPromocion As String) 'ECASTILLO 04.07.2020
    Dim x As Integer
    Dim encontro As Boolean
    Dim cuenta As Integer
    While x < xCuantaMaxima.Count(1) And encontro = False
        If xCuantaMaxima(x, 0) = strPromocion Then
            encontro = True
            cuenta = xCuantaMaxima(x, 1)
        End If
        x = x + 1
    Wend
    ObtieneCuentaMaxima = cuenta
End Function
Function AgregaCuentaMaxima(strPromocion As String, cantidad As Integer) 'ECASTILLO 04.07.2020
    Dim ultimo As Integer
    Dim i As Integer
    Dim encontro As Boolean
    Dim estaen As Integer
    ultimo = xCuantaMaxima.Count(1) 'Llena con la ultima posicion disponible
    encontro = False
    While i < ultimo
        If encontro = False Then
          If xCuantaMaxima(i, 0) = strPromocion Then
            estaen = i
            encontro = True
          Else
            encontro = False
            estaen = 0
          End If
        End If
        i = i + 1
    Wend
    If encontro = False Then
        xCuantaMaxima.AppendRows
        ultimo = xCuantaMaxima.Count(1) - 1 'Llena con la ultima posicion disponible
        xCuantaMaxima(ultimo, 0) = strPromocion
        xCuantaMaxima(ultimo, 1) = cantidad
    Else
        If xCuantaMaxima(estaen, 1) <> cantidad Then
            xCuantaMaxima(estaen, 0) = strPromocion
            xCuantaMaxima(estaen, 1) = cantidad
        End If
    End If
End Function
Function AGREGAREGALOBK(Codigo As String, _
                               Descripcion As String, _
                               cantidad As Integer, _
                               FlagFraccion As String, _
                               Precio As Double, _
                               TipoVenta As TipoVenta, _
                               Regalo As eProdRegalo, _
                               Optional Dato1 As String = "", _
                               Optional Dato2 As String = "", _
                               Optional Dato3 As String = "", _
                               Optional Dato4 As String = "", _
                               Optional Dato5 As String = "", _
                               Optional IndicadorReceta As String = "", _
                               Optional pctComision As Double = 0, _
                               Optional NroLote As String = "", _
                               Optional FchVmto As String = "", _
                               Optional flgFarmaco As String = "0", _
                               Optional PrecPublic As Double, Optional CadenaPromocion As String = "", Optional CadenaPromocionDcto As String = "", Optional FracEnabled As String = "", Optional CantPuntos As String, Optional CantPuntosRedimir As Double)
    Dim JU As Integer
    
    
    
    If xProductoRegaloBK.Count(1) = 0 Then
        
    JU = xProductoRegaloBK.Count(1)
    xProductoRegaloBK.AppendRows
        xProductoRegaloBK(JU, 0) = Codigo
    xProductoRegaloBK(JU, 1) = Descripcion
    xProductoRegaloBK(JU, 2) = cantidad
    xProductoRegaloBK(JU, 3) = FlagFraccion
    xProductoRegaloBK(JU, 4) = Precio
    xProductoRegaloBK(JU, 5) = TipoVenta
    xProductoRegaloBK(JU, 6) = Regalo
    xProductoRegaloBK(JU, 7) = Dato1
    xProductoRegaloBK(JU, 8) = Dato2
    xProductoRegaloBK(JU, 9) = Dato3
    xProductoRegaloBK(JU, 10) = Dato4
    xProductoRegaloBK(JU, 11) = Dato5
    xProductoRegaloBK(JU, 12) = IndicadorReceta
    xProductoRegaloBK(JU, 13) = pctComision
    xProductoRegaloBK(JU, 14) = NroLote
    xProductoRegaloBK(JU, 15) = FchVmto
    xProductoRegaloBK(JU, 16) = flgFarmaco
    xProductoRegaloBK(JU, 17) = PrecPublic
    xProductoRegaloBK(JU, 18) = CadenaPromocion
    xProductoRegaloBK(JU, 19) = CadenaPromocionDcto
    xProductoRegaloBK(JU, 20) = FracEnabled
    xProductoRegaloBK(JU, 21) = CantPuntos
    xProductoRegaloBK(JU, 22) = CantPuntosRedimir
    xProductoRegaloBK(JU, 23) = cantidad

Else
JU = 0
Dim HAY As Integer
HAY = -1
'HAY = 0
    While JU < xProductoRegaloBK.Count(1)
    
    
    If xProductoRegaloBK(JU, 0) = Codigo And xProductoRegaloBK(JU, 6) = Regalo Then
       HAY = JU
       JU = 10000
    End If
    JU = JU + 1
    Wend
    
    JU = HAY
    If HAY < 0 Then
       
        JU = xProductoRegaloBK.Count(1)
        xProductoRegaloBK.AppendRows
    End If
    xProductoRegaloBK(JU, 0) = Codigo
    xProductoRegaloBK(JU, 1) = Descripcion
    xProductoRegaloBK(JU, 2) = cantidad
    xProductoRegaloBK(JU, 3) = FlagFraccion
    xProductoRegaloBK(JU, 4) = Precio
    xProductoRegaloBK(JU, 5) = TipoVenta
    xProductoRegaloBK(JU, 6) = Regalo
    xProductoRegaloBK(JU, 7) = Dato1
    xProductoRegaloBK(JU, 8) = Dato2
    xProductoRegaloBK(JU, 9) = Dato3
    xProductoRegaloBK(JU, 10) = Dato4
    xProductoRegaloBK(JU, 11) = Dato5
    xProductoRegaloBK(JU, 12) = IndicadorReceta
    xProductoRegaloBK(JU, 13) = pctComision
    xProductoRegaloBK(JU, 14) = NroLote
    xProductoRegaloBK(JU, 15) = FchVmto
    xProductoRegaloBK(JU, 16) = flgFarmaco
    xProductoRegaloBK(JU, 17) = PrecPublic
    xProductoRegaloBK(JU, 18) = CadenaPromocion
    xProductoRegaloBK(JU, 19) = CadenaPromocionDcto
    xProductoRegaloBK(JU, 20) = FracEnabled
    xProductoRegaloBK(JU, 21) = CantPuntos
    xProductoRegaloBK(JU, 22) = CantPuntosRedimir
    xProductoRegaloBK(JU, 23) = cantidad

    End If
    
End Function

Function AgregaMetodoSegmento(ByVal localCode As String, _
                         ByVal typeCode As String, _
                         ByVal amount As String, _
                         ByVal day As String, _
                         ByVal message As String, _
                         ByVal deliveryTime As String, _
                         ByVal Value As String, _
                         ByVal valueEnd As String, _
                         ByVal time As String, _
                         ByVal starHour As String, _
                         ByVal endHour As String) As XArrayDB
                         
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If xMetodoSegmento.Count(1) < 0 Then Exit Function
    'I.ECASTILLO 17.12.2020
    Dim flg_2e_reserva
    Dim noAgrega2Metodo As Boolean
    Dim ii As Integer
    Dim yy As Integer
    Dim localPosu As String
    localPosu = objLocal.GetCodPosu(localCode)
    '
    If xMetodoSegmento.Count(1) > 0 Then
        noAgrega2Metodo = True
        'yy = 0
        'While yy < xFormaPago.Count(1)
        '    If (xFormaPago(yy, 0) = "002" Or _
        '       xFormaPago(yy, 0) = "001") And _
        '       (CodigoFormaPago <> "011" And CodigoFormaPago <> "008") Then
        '       noAgrega2Metodo = True
        '       GoTo continuaFlujoFP
        '    Else
        '       noAgrega2Metodo = False
        '    End If
        '    yy = yy + 1
        'Wend
    End If
    'F.ECASTILLO 17.12.2020
continuaFlujoFP:
    Dim i As Integer
    Dim encontro As Boolean
    aux = xMetodoSegmento.Count(1)
    While i < aux
        If xMetodoSegmento(i, 0) = localPosu And _
            xMetodoSegmento(i, 1) = typeCode And _
            xMetodoSegmento(i, 2) = day And _
            xMetodoSegmento(i, 3) = Value Then
            ultimo = i
            encontro = True
GoTo j
        Else
            encontro = False
            ultimo = xMetodoSegmento.Count(1) 'Llena con la ultima posicion disponible
        End If
        i = i + 1
    Wend
    If encontro = False And noAgrega2Metodo = False Then
        xMetodoSegmento.AppendRows
    End If
j:
    If (noAgrega2Metodo = True And encontro = False) Then Exit Function
    
    If xMetodoSegmento.Count(1) = 0 Then ultimo = 0: xMetodoSegmento.AppendRows
    xMetodoSegmento(ultimo, 0) = localPosu
    xMetodoSegmento(ultimo, 1) = typeCode
    xMetodoSegmento(ultimo, 2) = day
    xMetodoSegmento(ultimo, 3) = Value
    xMetodoSegmento(ultimo, 4) = valueEnd
    xMetodoSegmento(ultimo, 5) = message
    xMetodoSegmento(ultimo, 6) = deliveryTime
    xMetodoSegmento(ultimo, 7) = amount
    xMetodoSegmento(ultimo, 8) = time
    xMetodoSegmento(ultimo, 9) = starHour
    xMetodoSegmento(ultimo, 10) = endHour
            
    Set AgregaMetodoSegmento = xMetodoSegmento
End Function

Function validaCapacidad() As String
    Dim resp As Dictionary
    Dim i As Integer
    Dim x As Integer
    validaCapacidad = "0"
    If xMetodoSegmento.Count(1) < 0 Then Exit Function
    
    i = xMetodoSegmento.Count(1)
    While x < i And x <= 1
        If xMetodoSegmento(x, 1) = "RET" Then validaCapacidad = "1": Exit Function
        Set resp = objWS.validaSegmento(xMetodoSegmento(x, 2), _
                                        xMetodoSegmento(x, 3), _
                                        xMetodoSegmento(x, 0), _
                                        xMetodoSegmento(x, 1), _
                                        "CALL", _
                                        xMetodoSegmento(x, 9), _
                                        xMetodoSegmento(x, 10))
        If Not resp Is Nothing Then
            If IsObject(resp("data")) = False Then
                Exit Function
            End If
            If resp("data")("code") = "1" Then
                validaCapacidad = "1"
            Else
                validaCapacidad = "0"
            End If
        End If
        x = x + 1
    Wend
End Function

