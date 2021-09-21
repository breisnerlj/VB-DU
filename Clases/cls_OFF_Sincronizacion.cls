VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_OFF_Sincronizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_strConexionLocal As String
Private m_strIni As String
Private m_AppPath As String
Private m_BKPath As String
Private m_oradb As OraDatabase

'*************************************************************************************
' Cuando se inicializa la clase se crean todas las variables
'*************************************************************************************
Private Sub Class_Initialize()
    'Path de ubicacion de los archivos de contingencia a eviar
    m_AppPath = App.Path & IIf(right(App.Path, 1) = "\", "", "\")
    
    'm_BKPath = m_AppPath & "bkContingencia"
    m_BKPath = m_AppPath & "backup"
    
    'Cadena de conexion para abrir los archivos de contingencia
    m_strConexionLocal = "Provider = Microsoft.Jet.OLEDB.4.0;" & _
                         "Data Source=" & m_AppPath & "; " & _
                         "Extended Properties=""Text;FMT=Delimited"" "
                         
        'conexion con oracle
    Set m_oradb = godbVentas
End Sub

Public Function ObtenerListaUsuarios() As oraDynaset
    Dim strExec As String, i As Integer
    
    On Error GoTo ErrorHandler
    
    'Falso positivo
    Set ObtenerListaUsuarios = Nothing
    
    'Limpiar parametros existentes
    For i = m_oradb.Parameters.Count - 1 To 0 Step -1
        m_oradb.Parameters.Remove i
    Next
    
    'Agregar parametros
    
    m_oradb.Parameters.Add "A_SEC_ARCHIVO", "", ORAPARM_INPUT, ORATYPE_VARCHAR2
    m_oradb.Parameters.Add "A_COD_LOCAL", objUsuario.CodigoLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2
    
    'Armar cadena del SP
    strExec = "BEGIN " & _
              "  :RESULT := BTLPROD.PKG_CONTINGENCIA.FN_LISTA_USUARIO ( " & _
                                            ":A_SEC_ARCHIVO, " & _
                                            ":A_COD_LOCAL); END;"
    
    'Ejecutar SP y devolver resultado
    Set ObtenerListaUsuarios = m_oradb.CreatePlsqlDynaset(strExec, "RESULT", 0&)
    
    Exit Function
      
ErrorHandler:
    Set ObtenerListaUsuarios = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function EnviarArchivos() As oraDynaset
    Dim rsDetVta As ADODB.Recordset
    Dim rsFPago As ADODB.Recordset
    Dim byteNumElementosDetVta As Integer
    Dim byteNumElementosFormaPago As Integer
    Dim row As Integer, i As Integer
    Dim strExec As String
    'Dim oFSO As Scripting.FileSystemObject
    
    'Variables detalles de venta
    Dim SecVenta As OracleInProcServer.OraParamArray
    Dim CodCia As OracleInProcServer.OraParamArray
    Dim CodLocal As OracleInProcServer.OraParamArray
    Dim CodTipoDocumento As OracleInProcServer.OraParamArray
    Dim CodMaquina As OracleInProcServer.OraParamArray
    Dim NumDocumento As OracleInProcServer.OraParamArray
    Dim numItem As OracleInProcServer.OraParamArray
    Dim CodProducto As OracleInProcServer.OraParamArray
    Dim desProducto As OracleInProcServer.OraParamArray
    Dim CtdProducto As OracleInProcServer.OraParamArray
    Dim flgFraccion As OracleInProcServer.OraParamArray
    Dim PctDescuento As OracleInProcServer.OraParamArray
    Dim PrcUnitario As OracleInProcServer.OraParamArray
    Dim PrcOriginal As OracleInProcServer.OraParamArray
    Dim MtoIgv As OracleInProcServer.OraParamArray
    Dim MtoExonerado As OracleInProcServer.OraParamArray
    Dim MtoSubTotal As OracleInProcServer.OraParamArray
    Dim usuEmision As OracleInProcServer.OraParamArray
    Dim fchEmision As OracleInProcServer.OraParamArray
    Dim usuAnulacion As OracleInProcServer.OraParamArray
    Dim fchAnulacion As OracleInProcServer.OraParamArray
    Dim FlgModPrecio As OracleInProcServer.OraParamArray
    Dim FlgRegalo As OracleInProcServer.OraParamArray
    Dim CodEstado As OracleInProcServer.OraParamArray
    Dim tipCliente As OracleInProcServer.OraParamArray
    Dim numRuc As OracleInProcServer.OraParamArray
    Dim dirCliente As OracleInProcServer.OraParamArray
    Dim nomCliente As OracleInProcServer.OraParamArray
    Dim PctIgv As OracleInProcServer.OraParamArray
    Dim PartidaArancelaria As OracleInProcServer.OraParamArray
    Dim CtdFracciona As OracleInProcServer.OraParamArray
    Dim UsuModPrecio As OracleInProcServer.OraParamArray
    Dim tipMovimiento As OracleInProcServer.OraParamArray
    Dim CodSerieRel As OracleInProcServer.OraParamArray


    'Variables detalles formas de pago
    Dim secArchivoFP As OracleInProcServer.OraParamArray
    Dim secVentaFP As OracleInProcServer.OraParamArray
    Dim numItemFP As OracleInProcServer.OraParamArray
    Dim codPagoFP As OracleInProcServer.OraParamArray
    Dim codMonedaFP As OracleInProcServer.OraParamArray
    Dim mtoImporteFP As OracleInProcServer.OraParamArray
    Dim mtoSolesFP As OracleInProcServer.OraParamArray
    Dim mtoVueltoFP As OracleInProcServer.OraParamArray
    Dim numTarjetaFP As OracleInProcServer.OraParamArray
    Dim fchVencimientoFP As OracleInProcServer.OraParamArray
    Dim numCuotasFP As OracleInProcServer.OraParamArray
    Dim tipoCambioFP As OracleInProcServer.OraParamArray
    Dim pctIgvFP As OracleInProcServer.OraParamArray
    Dim partidaArancelariaFP As OracleInProcServer.OraParamArray
    Dim ctdFraccionaFP As OracleInProcServer.OraParamArray
    Dim usuModPrecioFP As OracleInProcServer.OraParamArray
    Dim tipMovimientoFP As OracleInProcServer.OraParamArray
    Dim NumAutorizacionFP As OracleInProcServer.OraParamArray
    On Error GoTo ErrorHandler
    
    'Falso positivo
    Set EnviarArchivos = Nothing
    
'    Set oFSO = New Scripting.FileSystemObject
'    If oFSO.FileExists(m_AppPath & "\detalleventa.txt") Then Exit Function
'    If Len(Dir$(m_AppPath & "detalleventa.txt", vbHidden)) = 0 Then Exit Function
    If Len(Dir$(strDetalleVentaXML, vbHidden)) = 0 Then Exit Function
        
    'Ejecuta la consulta sql para llenar el recordset de detalle venta
    Set rsDetVta = New ADODB.Recordset
    With rsDetVta
'        .ActiveConnection = m_strConexionLocal 'conn
        .ActiveConnection = gstrConexion
'        .Source = "SELECT * FROM detalleventa.txt"
        .Source = strDetalleVentaXML
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open
        .ActiveConnection = Nothing
    End With
    
    
    
    
    'Variables para el numero de elementos en el detalle de ventas a enviar
    byteNumElementosDetVta = IIf(rsDetVta.RecordCount < 0, 1, rsDetVta.RecordCount)
    If byteNumElementosDetVta = 0 Then Exit Function
'        Err.Raise vbObjectError + 1, _
'                  "cls_OFF_Sincronizacion.EnviarArchivos()", _
'                  "El archivo de contingencia 'detalleventa.txt' está vacío."
    
'    If oFSO.FileExists(m_AppPath & "\pagoventa.txt") Then Exit Function
'    If Len(Dir$(m_AppPath & "pagoventa.txt", vbHidden)) = 0 Then Exit Function
    If Len(Dir$(strPagoVentaXML, vbHidden)) = 0 Then Exit Function
    
    'Ejecuta la consulta sql para llenar el recordset de detalle de formas de pago
    Set rsFPago = New ADODB.Recordset
    With rsFPago
'        .ActiveConnection = m_strConexionLocal 'conn
'        .Source = "SELECT * FROM pagoventa.txt"
        .ActiveConnection = gstrConexion
        .Source = strPagoVentaXML
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open
        .ActiveConnection = Nothing
    End With
    'Variables para el numero de elementos en el detalle de ventas a enviar
    byteNumElementosFormaPago = IIf(rsFPago.RecordCount < 0, 1, rsFPago.RecordCount)
    If byteNumElementosFormaPago = 0 Then Exit Function
'        Err.Raise vbObjectError + 1, _
'                  "cls_OFF_Sincronizacion.EnviarArchivos()", _
'                  "El archivo de contingencia 'pagoventa.txt' está vacío."

    'Limpiar parametros existentes
    For i = m_oradb.Parameters.Count - 1 To 0 Step -1
        m_oradb.Parameters.Remove i
    Next
    
    'Agregar parametros compañia
    m_oradb.Parameters.Add "A_COD_LOCAL", objUsuario.CodigoLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Compañia
    m_oradb.Parameters.Add "A_COD_USUARIO", objUsuario.Codigo, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Compañia
    m_oradb.Parameters.Add "A_DES_NOMBRE", objUsuario.Nombre, ORAPARM_INPUT, ORATYPE_VARCHAR2 'Compañia
    
    'Agregar parametros detalle venta
    m_oradb.Parameters.AddTable "A_CAD_SEC_VENTA", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set SecVenta = m_oradb.Parameters("A_CAD_SEC_VENTA")
    m_oradb.Parameters.AddTable "A_CAD_COD_CIA", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set CodCia = m_oradb.Parameters("A_CAD_COD_CIA")
    m_oradb.Parameters.AddTable "A_CAD_COD_LOCAL", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set CodLocal = m_oradb.Parameters("A_CAD_COD_LOCAL")
    m_oradb.Parameters.AddTable "A_CAD_COD_TIPO_DOCUMENTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set CodTipoDocumento = m_oradb.Parameters("A_CAD_COD_TIPO_DOCUMENTO")
    m_oradb.Parameters.AddTable "A_CAD_COD_MAQUINA", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set CodMaquina = m_oradb.Parameters("A_CAD_COD_MAQUINA")
    m_oradb.Parameters.AddTable "A_CAD_NUM_DOCUMENTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set NumDocumento = m_oradb.Parameters("A_CAD_NUM_DOCUMENTO")
    m_oradb.Parameters.AddTable "A_CAD_NUM_ITEM", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set numItem = m_oradb.Parameters("A_CAD_NUM_ITEM")
    m_oradb.Parameters.AddTable "A_CAD_COD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set CodProducto = m_oradb.Parameters("A_CAD_COD_PRODUCTO")
    m_oradb.Parameters.AddTable "A_CAD_DES_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set desProducto = m_oradb.Parameters("A_CAD_DES_PRODUCTO")
    m_oradb.Parameters.AddTable "A_CAD_CTD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set CtdProducto = m_oradb.Parameters("A_CAD_CTD_PRODUCTO")
    m_oradb.Parameters.AddTable "A_CAD_FLG_FRACCION", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set flgFraccion = m_oradb.Parameters("A_CAD_FLG_FRACCION")
    m_oradb.Parameters.AddTable "A_CAD_PCT_DESCUENTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set PctDescuento = m_oradb.Parameters("A_CAD_PCT_DESCUENTO")
    m_oradb.Parameters.AddTable "A_CAD_PRC_UNITARIO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set PrcUnitario = m_oradb.Parameters("A_CAD_PRC_UNITARIO")
    m_oradb.Parameters.AddTable "A_CAD_PRC_ORIGINAL", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set PrcOriginal = m_oradb.Parameters("A_CAD_PRC_ORIGINAL")
    m_oradb.Parameters.AddTable "A_CAD_MTO_IGV", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set MtoIgv = m_oradb.Parameters("A_CAD_MTO_IGV")
    m_oradb.Parameters.AddTable "A_CAD_MTO_EXONERADO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set MtoExonerado = m_oradb.Parameters("A_CAD_MTO_EXONERADO")
    m_oradb.Parameters.AddTable "A_CAD_MTO_SUBTOTAL", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set MtoSubTotal = m_oradb.Parameters("A_CAD_MTO_SUBTOTAL")
    m_oradb.Parameters.AddTable "A_CAD_USU_EMISION", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set usuEmision = m_oradb.Parameters("A_CAD_USU_EMISION")
    m_oradb.Parameters.AddTable "A_CAD_FCH_EMISION", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set fchEmision = m_oradb.Parameters("A_CAD_FCH_EMISION")
    m_oradb.Parameters.AddTable "A_CAD_USU_ANULACION", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set usuAnulacion = m_oradb.Parameters("A_CAD_USU_ANULACION")
    m_oradb.Parameters.AddTable "A_CAD_FCH_ANULACION", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set fchAnulacion = m_oradb.Parameters("A_CAD_FCH_ANULACION")
    m_oradb.Parameters.AddTable "A_CAD_FLG_MODPRECIO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set FlgModPrecio = m_oradb.Parameters("A_CAD_FLG_MODPRECIO")
    m_oradb.Parameters.AddTable "A_CAD_FLG_REGALO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set FlgRegalo = m_oradb.Parameters("A_CAD_FLG_REGALO")
    m_oradb.Parameters.AddTable "A_CAD_COD_ESTADO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set CodEstado = m_oradb.Parameters("A_CAD_COD_ESTADO")
    m_oradb.Parameters.AddTable "A_CAD_TIP_CLIENTE", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set tipCliente = m_oradb.Parameters("A_CAD_TIP_CLIENTE")
    m_oradb.Parameters.AddTable "A_CAD_NUM_RUC", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set numRuc = m_oradb.Parameters("A_CAD_NUM_RUC")
    m_oradb.Parameters.AddTable "A_CAD_DIR_CLIENTE", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set dirCliente = m_oradb.Parameters("A_CAD_DIR_CLIENTE")
    m_oradb.Parameters.AddTable "A_CAD_NOM_CLIENTE", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set nomCliente = m_oradb.Parameters("A_CAD_NOM_CLIENTE")
    m_oradb.Parameters.AddTable "A_CAD_PCT_IGV", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set PctIgv = m_oradb.Parameters("A_CAD_PCT_IGV")
    m_oradb.Parameters.AddTable "A_CAD_PARTIDA_ARANCELARIA", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set PartidaArancelaria = m_oradb.Parameters("A_CAD_PARTIDA_ARANCELARIA")
    m_oradb.Parameters.AddTable "A_CAD_ctd_fracciona", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set CtdFracciona = m_oradb.Parameters("A_CAD_CTD_FRACCIONA")
    m_oradb.Parameters.AddTable "A_CAD_usu_mod_precio", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set UsuModPrecio = m_oradb.Parameters("A_CAD_USU_MOD_PRECIO")
    m_oradb.Parameters.AddTable "A_CAD_TIP_MOVIMIENTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set tipMovimiento = m_oradb.Parameters("A_CAD_TIP_MOVIMIENTO")
    
    m_oradb.Parameters.AddTable "A_CAD_COD_SERIEL_REL", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosDetVta, 200
    Set CodSerieRel = m_oradb.Parameters("A_CAD_COD_SERIEL_REL")
    
    


    'Agregar parametros detalle formas de pago
    m_oradb.Parameters.AddTable "A_CA_FP_SEC_ARCHIVO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosFormaPago, 200
    Set secArchivoFP = m_oradb.Parameters("A_CA_FP_SEC_ARCHIVO")
    m_oradb.Parameters.AddTable "A_CA_FP_SEC_VENTA", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosFormaPago, 200
    Set secVentaFP = m_oradb.Parameters("A_CA_FP_SEC_VENTA")
    m_oradb.Parameters.AddTable "A_CA_FP_NUM_ITEM", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosFormaPago, 200
    Set numItemFP = m_oradb.Parameters("A_CA_FP_NUM_ITEM")
    m_oradb.Parameters.AddTable "A_CA_FP_COD_PAGO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosFormaPago, 200
    Set codPagoFP = m_oradb.Parameters("A_CA_FP_COD_PAGO")
    m_oradb.Parameters.AddTable "A_CA_FP_COD_MONEDA", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosFormaPago, 200
    Set codMonedaFP = m_oradb.Parameters("A_CA_FP_COD_MONEDA")
    m_oradb.Parameters.AddTable "A_CA_FP_MTO_IMPORTE", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosFormaPago, 200
    Set mtoImporteFP = m_oradb.Parameters("A_CA_FP_MTO_IMPORTE")
    m_oradb.Parameters.AddTable "A_CA_FP_MTO_SOLES", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosFormaPago, 200
    Set mtoSolesFP = m_oradb.Parameters("A_CA_FP_MTO_SOLES")
    m_oradb.Parameters.AddTable "A_CA_FP_MTO_VUELTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosFormaPago, 200
    Set mtoVueltoFP = m_oradb.Parameters("A_CA_FP_MTO_VUELTO")
    m_oradb.Parameters.AddTable "A_CA_FP_NUM_TARJETA", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosFormaPago, 200
    Set numTarjetaFP = m_oradb.Parameters("A_CA_FP_NUM_TARJETA")
    m_oradb.Parameters.AddTable "A_CA_FP_FCH_VENCIMIENTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosFormaPago, 200
    Set fchVencimientoFP = m_oradb.Parameters("A_CA_FP_FCH_VENCIMIENTO")
    m_oradb.Parameters.AddTable "A_CA_FP_NUM_CUOTAS", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosFormaPago, 200
    Set numCuotasFP = m_oradb.Parameters("A_CA_FP_NUM_CUOTAS")
    m_oradb.Parameters.AddTable "A_CA_FP_TIPO_CAMBIO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosFormaPago, 200
    Set tipoCambioFP = m_oradb.Parameters("A_CA_FP_TIPO_CAMBIO")
    
    m_oradb.Parameters.AddTable "A_CA_FP_NUM_AUTORIZACION", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosFormaPago, 200
    Set NumAutorizacionFP = m_oradb.Parameters("A_CA_FP_NUM_AUTORIZACION")
    
    
    'Recorrer el archivo detalleventa y establecer valores de parametros
    row = 0
    While Not rsDetVta.EOF
        SecVenta(row) = Trim(rsDetVta.Fields("SEC_VENTA").Value)
        CodCia(row) = Trim(rsDetVta.Fields("COD_CIA").Value)
        CodLocal(row) = Trim(rsDetVta.Fields("COD_LOCAL").Value)
        CodTipoDocumento(row) = Trim(rsDetVta.Fields("COD_TIPO_DOCUMENTO").Value)
        CodMaquina(row) = Trim(rsDetVta.Fields("COD_MAQUINA").Value)
        NumDocumento(row) = Trim(rsDetVta.Fields("NUM_DOCUMENTO").Value)
        numItem(row) = Trim(rsDetVta.Fields("NUM_ITEM").Value)
        CodProducto(row) = Trim(rsDetVta.Fields("COD_PRODUCTO").Value)
        desProducto(row) = Trim(rsDetVta.Fields("DES_PRODUCTO").Value)
        CtdProducto(row) = Trim(rsDetVta.Fields("CTD_PRODUCTO").Value)
        flgFraccion(row) = Trim(rsDetVta.Fields("FLG_FRACCION").Value)
        PctDescuento(row) = Trim(rsDetVta.Fields("PCT_DESCUENTO").Value)
        PrcUnitario(row) = Trim(rsDetVta.Fields("PRC_UNITARIO").Value)
        PrcOriginal(row) = Trim(rsDetVta.Fields("PRC_ORIGINAL").Value)
        MtoIgv(row) = Trim(rsDetVta.Fields("MTO_IGV").Value)
        MtoExonerado(row) = Trim(rsDetVta.Fields("MTO_EXONERADO").Value)
        MtoSubTotal(row) = Trim(rsDetVta.Fields("MTO_SUBTOTAL").Value)
        usuEmision(row) = Trim(rsDetVta.Fields("USU_EMISION").Value)
        fchEmision(row) = Trim(rsDetVta.Fields("FCH_EMISION").Value)
        usuAnulacion(row) = Trim(rsDetVta.Fields("USU_ANULACION").Value)
        fchAnulacion(row) = Trim(rsDetVta.Fields("FCH_ANULACION").Value)
        FlgModPrecio(row) = Trim(rsDetVta.Fields("FLG_MODPRECIO").Value)
        FlgRegalo(row) = Trim(rsDetVta.Fields("FLG_REGALO").Value)
        CodEstado(row) = Trim(rsDetVta.Fields("COD_ESTADO").Value)
        tipCliente(row) = Trim(rsDetVta.Fields("TIP_CLIENTE").Value)
        numRuc(row) = Trim(rsDetVta.Fields("NUM_RUC").Value)
        dirCliente(row) = Trim(rsDetVta.Fields("DIR_CLIENTE").Value)
        nomCliente(row) = Trim(rsDetVta.Fields("NOM_CLIENTE").Value)
        PctIgv(row) = Trim(rsDetVta.Fields("PCT_IGV").Value)
        PartidaArancelaria(row) = Trim(rsDetVta.Fields("PARTIDA_ARANCELARIA").Value)
        CtdFracciona(row) = Trim(rsDetVta.Fields("CTD_FRACCIONA").Value)
        UsuModPrecio(row) = Trim(rsDetVta.Fields("USU_MOD_PRECIO").Value)
        tipMovimiento(row) = Trim(rsDetVta.Fields("TIP_MOVIMIENTO").Value)
        CodSerieRel(row) = Trim(rsDetVta.Fields("COD_SERIE_REL").Value)
                    
        row = row + 1
        rsDetVta.MoveNext
    Wend

    'Recorrer archivo pagoventa y establecer valores de parametros
    row = 0
    While Not rsFPago.EOF
        secVentaFP(row) = Trim(rsFPago.Fields("SEC_VENTA").Value)
        numItemFP(row) = Trim(rsFPago.Fields("NUM_ITEM").Value)
        codPagoFP(row) = Trim(rsFPago.Fields("COD_PAGO").Value)
        codMonedaFP(row) = Trim(rsFPago.Fields("COD_MONEDA").Value)
        mtoImporteFP(row) = Trim(rsFPago.Fields("MTO_IMPORTE").Value)
        mtoSolesFP(row) = Trim(rsFPago.Fields("MTO_SOLES").Value)
        mtoVueltoFP(row) = Trim(rsFPago.Fields("MTO_VUELTO").Value)
        numTarjetaFP(row) = Trim(rsFPago.Fields("NUM_TARJETA").Value)
        fchVencimientoFP(row) = Trim(rsFPago.Fields("FCH_VENCIMIENTO").Value)
        numCuotasFP(row) = Trim(rsFPago.Fields("NUM_CUOTAS").Value)
        tipoCambioFP(row) = Trim(rsFPago.Fields("TIPO_CAMBIO").Value)
        NumAutorizacionFP(row) = Trim(rsFPago.Fields("NUM_AUTORIZACION").Value)
        row = row + 1
        rsFPago.MoveNext
    Wend
    
    'Armar cadena del SP
    strExec = "BEGIN " & _
              "  :RESULT := BTLPROD.PKG_CONTINGENCIA.SP_GRABA (:A_COD_LOCAL,               :A_COD_USUARIO,              :A_DES_NOMBRE,          " & _
                                                              ":A_CAD_sec_venta,           :A_CAD_cod_cia,              :A_CAD_cod_local,       " & _
                                                              ":A_CAD_cod_tipo_documento,  :A_CAD_cod_maquina,          :A_CAD_num_documento,   " & _
                                                              ":A_CAD_num_item,            :A_CAD_cod_producto,         :A_CAD_des_producto,    " & _
                                                              ":A_CAD_ctd_producto,        :A_CAD_flg_fraccion,         :A_CAD_pct_descuento,   " & _
                                                              ":A_CAD_prc_unitario,        :A_CAD_prc_original,         :A_CAD_mto_igv,         " & _
                                                              ":A_CAD_mto_exonerado,       :A_CAD_mto_subtotal,         :A_CAD_usu_emision,     " & _
                                                              ":A_CAD_fch_emision,         :A_CAD_usu_anulacion,        :A_CAD_fch_anulacion,   " & _
                                                              ":A_CAD_flg_modprecio,       :A_CAD_flg_regalo,           :A_CAD_cod_estado,      " & _
                                                              ":A_CAD_tip_cliente,         :A_CAD_num_ruc,              :A_CAD_dir_cliente,     " & _
                                                              ":A_CAD_nom_cliente,         :A_CAD_pct_igv,              :A_CAD_partida_arancelaria, " & _
                                                              ":A_CAD_ctd_fracciona ,      :A_CAD_usu_mod_precio,       :A_CAD_tip_movimiento,      " & _
                                                              ":A_CA_FP_sec_archivo,       :A_CA_FP_sec_venta,          :A_CA_FP_num_item,      :A_CA_FP_cod_pago,  " & _
                                                              ":A_CA_FP_cod_moneda,        :A_CA_FP_mto_importe,        :A_CA_FP_mto_soles,     :A_CA_FP_mto_vuelto," & _
                                                              ":A_CA_FP_num_tarjeta,       :A_CA_FP_fch_vencimiento,    :A_CA_FP_num_cuotas,    :A_CA_FP_tipo_cambio, :A_CA_FP_NUM_AUTORIZACION, :A_CAD_COD_SERIEL_REL); " & _
              "END;"
    
    'Ejecutar SP y devolver resultado
    Set EnviarArchivos = m_oradb.CreatePlsqlDynaset(strExec, "RESULT", 0&)
    
    'Backup archivos actuales
    Call RespaldarArchivos

    'Vaciar archivos actuales
'    Call VaciarArchivo(m_AppPath & "\detalleventa.txt")
'    Call VaciarArchivo(m_AppPath & "\pagoventa.txt")
    
final:
    'Cerrar y eliminar cursores
    If rsDetVta.State = adStateOpen Then rsDetVta.Close
    If rsFPago.State = adStateOpen Then rsFPago.Close
    Set rsDetVta = Nothing
    Set rsFPago = Nothing
'    Set oFSO = Nothing
    
    'Eliminar Variables detalles de venta
    Set SecVenta = Nothing
    Set CodCia = Nothing
    Set CodLocal = Nothing
    Set CodTipoDocumento = Nothing
    Set CodMaquina = Nothing
    Set NumDocumento = Nothing
    Set numItem = Nothing
    Set CodProducto = Nothing
    Set desProducto = Nothing
    Set CtdProducto = Nothing
    Set flgFraccion = Nothing
    Set PctDescuento = Nothing
    Set PrcUnitario = Nothing
    Set PrcOriginal = Nothing
    Set MtoIgv = Nothing
    Set MtoExonerado = Nothing
    Set MtoSubTotal = Nothing
    Set usuEmision = Nothing
    Set fchEmision = Nothing
    Set usuAnulacion = Nothing
    Set fchAnulacion = Nothing
    Set FlgModPrecio = Nothing
    Set FlgRegalo = Nothing
    Set CodEstado = Nothing
    Set tipCliente = Nothing
    Set numRuc = Nothing
    Set dirCliente = Nothing
    Set nomCliente = Nothing
    Set PctIgv = Nothing
    Set PartidaArancelaria = Nothing
    Set CtdFracciona = Nothing
    Set UsuModPrecio = Nothing
    Set tipMovimiento = Nothing

    'Eliminar Variables detalles formas de pago
    Set secArchivoFP = Nothing
    Set secVentaFP = Nothing
    Set numItemFP = Nothing
    Set codPagoFP = Nothing
    Set codMonedaFP = Nothing
    Set mtoImporteFP = Nothing
    Set mtoSolesFP = Nothing
    Set mtoVueltoFP = Nothing
    Set numTarjetaFP = Nothing
    Set fchVencimientoFP = Nothing
    Set numCuotasFP = Nothing
    Set tipoCambioFP = Nothing
    Set pctIgvFP = Nothing
    Set partidaArancelariaFP = Nothing
    Set ctdFraccionaFP = Nothing
    Set usuModPrecioFP = Nothing
    Set tipMovimientoFP = Nothing
  
    Exit Function
      
ErrorHandler:
    Set EnviarArchivos = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo final
End Function

Public Function ProcesarLiquidacion(CodUsuario As String, _
                                    CodLiquidacion As String, _
                                    SecArchivo As String) As Boolean

    Dim i As Integer, strExec As String
    
    On Error GoTo ErrorHandler
    
    ProcesarLiquidacion = False
    
    'Limpiar parametros existentes
    For i = m_oradb.Parameters.Count - 1 To 0 Step -1
        m_oradb.Parameters.Remove i
    Next
    
    'Agregar parametros compañia
    m_oradb.Parameters.Add "A_SEC_ARCHIVO", SecArchivo, ORAPARM_INPUT, ORATYPE_VARCHAR2
    m_oradb.Parameters.Add "A_COD_LIQUIDACION", CodLiquidacion, ORAPARM_INPUT, ORATYPE_VARCHAR2
    m_oradb.Parameters.Add "A_COD_USUARIO", CodUsuario, ORAPARM_INPUT, ORATYPE_VARCHAR2
    
    'Armar cadena del SP
    strExec = "BEGIN " & _
              "     BTLPROD.PKG_CONTINGENCIA.SP_GRABA_DOCUMENTOS(:A_SEC_ARCHIVO, :A_COD_LIQUIDACION, :A_COD_USUARIO); " & _
              "END;"
    
    'Ejecutar SP y devolver resultado
    Call m_oradb.ExecuteSQL(strExec)
    
    ProcesarLiquidacion = True

final:
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo final
End Function

Private Function RespaldarArchivos() As Boolean
    'Backup de archivos enviados
'    Dim FileHandler As Scripting.FileSystemObject
'    Dim oFile As Scripting.File
    Dim strNow As String
    Dim strSourceDetVenta As String, strDestDetVenta As String
    Dim strSourceFPago As String, strDestFPago As String
    
    On Error GoTo ErrorHandler
    'falso positivo
    RespaldarArchivos = False
    
    'Nombres de archivos
    strNow = Format(Now, "ddmmyyyyhhnnss")
    'strSourceDetVenta = m_AppPath & "detalleventa.txt"
    strSourceDetVenta = strDetalleVentaXML
    'strDestDetVenta = m_BKPath & "\detalleventa" & strNow & ".txt"
    strDestDetVenta = m_BKPath & "\detalleventa" & strNow & ".xml"
    'strSourceFPago = m_AppPath & "pagoventa.txt"
    strSourceFPago = strPagoVentaXML
    'strDestFPago = m_BKPath & "\pagoventa" & strNow & ".txt"
    strDestFPago = m_BKPath & "\pagoventa" & strNow & ".xml"
    
'    Set FileHandler = New Scripting.FileSystemObject
    
    'Si carpeta no existe, crearla
'    If FileHandler.FolderExists(m_BKPath) = False Then Call FileHandler.CreateFolder(m_BKPath)
    If Len(Dir$(m_BKPath, vbDirectory)) = 0 Then Call MkDir(m_BKPath)
    
    'Backup DetalleVenta
'    Set oFile = FileHandler.GetFile(strSourceDetVenta)
'    Call oFile.Copy(strDestDetVenta)
'    Call oFile.Delete(True)
    Call FileCopy(strSourceDetVenta, strDestDetVenta)
    Call Kill(strSourceDetVenta)
    
    'DetalleFormasPago
'    Set oFile = FileHandler.GetFile(strSourceFPago)
'    Call oFile.Copy(strDestFPago)
'    Call oFile.Delete(True)
    Call FileCopy(strSourceFPago, strDestFPago)
    Call Kill(strSourceFPago)
    
    RespaldarArchivos = True
    
final:
'    Set oFile = Nothing
'    Set FileHandler = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo final
End Function


'Private Function VaciarArchivo(fName As String) As Boolean
'    Dim oFSO As Scripting.FileSystemObject
'    Dim oFTS As Scripting.TextStream
'    Dim sLine As String
'
'    On Error GoTo ErrorHandler
'
'    VaciarArchivo = False
'
'    Set oFSO = New Scripting.FileSystemObject
'
'    'Leer la cabecera del archivo
'    If oFSO.FileExists(fName) = True Then
'        Set oFTS = oFSO.OpenTextFile(fName)
'        Do While Not oFTS.AtEndOfStream
'            sLine = oFTS.ReadLine
'            Exit Do
'        Loop
'    End If
'    oFTS.Close
'
'    'Reemplazar el archivo y escribir la cabecera
'    Set oFTS = oFSO.CreateTextFile(fName, True, True)
'    oFTS.Write sLine
'
'    VaciarArchivo = True
'
'final:
'    oFTS.Close
'    Set oFTS = Nothing
'    Set oFSO = Nothing
'    Exit Function
'
'ErrorHandler:
'        Err.Raise Err.Number, Err.Source, Err.Description
'        GoTo final
'End Function

Private Sub Class_Terminate()
    Set m_oradb = Nothing
End Sub

Public Function ListaErrores(ByVal Secuencia As String) As oraDynaset
    On Error GoTo CtrlErr
        Set ListaErrores = gclsOracle.FN_Cursor("BTLPROD.PKG_CONTINGENCIA.FN_LISTA_ERROR", 0, Secuencia)
        
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "cls_OFF_Sincronizacion.ListaErrores", Err.Description
End Function
