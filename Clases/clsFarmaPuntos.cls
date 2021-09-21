VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFarmaPuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarTarjetaBean As clsTarjetaBean
Private mvarAfiliado As clsAfiliado
Private mvardireccionWSPuntos As String
Private mvarIdLocal As String
Private mvarIdCaja As String
Private mvarIdEmpleado As String
Private mvarCodigoAutorizacion As String
Private mvarWSCPuntos As clsOrbisFarma
Private oFarmaPuntosConstante As clsFPConstante
    
Private Sub Class_Initialize()
    Set mvarWSCPuntos = New clsOrbisFarma
    Set oFarmaPuntosConstante = New clsFPConstante
End Sub

Private Function LoadRows(Lista() As String) As XArrayDB
    Dim arr As XArrayDB
    Dim i As Integer
    
    Set arr = New XArrayDB
    arr.ReDim LBound(Lista), UBound(Lista), 0, 0
    
    For i = LBound(Lista) To UBound(Lista)
        arr.Value(i, 0) = Lista(i)
    Next i

    Set LoadRows = arr
End Function

Private Function pGetCadenaDeList(ByVal vLista As XArrayDB) As String
    Dim vValor As String, i As Integer
    For i = vLista.LowerBound(1) To i < vLista.UpperBound(1)
        If ((i + 1) = vLista.UpperBound(1)) Then
            vValor = vValor & vLista.Value(i, 0)
        Else
            vValor = vValor & vLista.Value(i, 0) & "|"
        End If
    Next i

    pGetCadenaDeList = vValor
End Function

'''Public Sub FarmaPuntos(idLocal As String, _
'''                       idCaja As String, _
'''                       direccionWSPuntos As String, _
'''                       CodigoAutorizacion As String, _
'''                       TIMEOUT As Integer)
'''    mvarIdLocal = idLocal
'''    mvarIdCaja = idCaja
'''    mvarCodigoAutorizacion = CodigoAutorizacion
'''    mvardireccionWSPuntos = direccionWSPuntos
'''End Sub

'''Public Property Get TarjetaBean() As clsTarjetaBean
'''    Set TarjetaBean = mvarTarjetaBean
'''End Property
'''
'''Public Sub EliminarTarjetaBean()
'''    Set mvarTarjetaBean = Nothing
'''End Sub
    
'''Public Function RegistrarAfiliado(ByVal NumeroTarjeta As String, _
'''                                  ByVal Afiliado As clsAfiliado, _
'''                                  ByVal idEmpleado As String) As clsTarjetaBean
'''    Dim stCardAffiliation As stCardAffiliation
'''
'''    If mvarTarjetaBean Is Nothing Then
'''        mvarTarjetaBean = New clsTarjetaBean
'''    End If
'''    'si la ejecucion es exitosa, no modificar el estado de la tarjeta
'''    ' si no hubo conexion con orbis estado de tarjeta = sin tarjeta
'''    On Error GoTo Control
'''
'''    If (NumeroTarjeta <> "" And Not Afiliado Is Nothing And idEmpleado <> "") Then
'''        Set stCardAffiliation = _
'''            mvarWSCPuntos.setCardAffiliation(mvarWSCPuntos.getFormatoDocIdentidad(NumeroTarjeta), _
'''                                             mvarWSCPuntos.getFormatoDocIdentidad(Afiliado.DNI), _
'''                                             Afiliado.Nombre, _
'''                                             Afiliado.Apellidos, _
'''                                             mvarWSCPuntos.getFormatearFecha(Afiliado.FechaNacimiento), _
'''                                             Trim$(Afiliado.Genero), _
'''                                             "", _
'''                                             "", _
'''                                             Trim$(Afiliado.Email), _
'''                                             Trim$(Afiliado.Telefono), _
'''                                             Trim$(Afiliado.Celular), _
'''                                             Afiliado.DireccionNormalizada, _
'''                                             "", _
'''                                             idEmpleado)
'''
'''        mvarTarjetaBean.NumeroTarjeta = stCardAffiliation.CardNumber
'''        mvarTarjetaBean.DNI = stCardAffiliation.AccountNumber
'''        mvarTarjetaBean.EstadoOperacion = stCardAffiliation.ErrorId
'''        mvarTarjetaBean.Mensaje = stCardAffiliation.Message
'''        mvarTarjetaBean.Fecha = stCardAffiliation.TransactionDate
'''    Else
'''        mvarTarjetaBean.EstadoOperacion = oFarmaPuntosConstante.PARAMETRO_INCOMPLETO
'''    End If
'''
'''    Set RegistrarAfiliado = mvarTarjetaBean
'''    Exit Function
'''Control:
'''    mvarTarjetaBean.IdTransaccion = "0"
'''    mvarTarjetaBean.DNI = Afiliado.DNI
'''    mvarTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.SIN_ESTADO
'''    mvarTarjetaBean.EstadoOperacion = oFarmaPuntosConstante.NO_CONEXION_ORBIS
'''    mvarTarjetaBean.Mensaje = Err.Description
'''    Set RegistrarAfiliado = mvarTarjetaBean
'''End Function
Public Function ObtenerAfiliadoInka(ByVal NroDocumento As String, _
                                         ByVal TipoDocumento As String) As clsAfiliado
    On Error GoTo Control
    Set mvarAfiliado = Nothing
    gvarTNSNAME2 = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "CNXINKCLB")
    gclsOracle.Cerrar 'Cierra RAC
    If MainInkaClub = False Then Exit Function 'Abre Inka
    If (NroDocumento <> "" And TipoDocumento <> "") Then
        Dim obj As oraDynaset
        'Set obj = gclsOracle.FN_Cursor("BTLPROD.FN_VALIDA_CLIENTE_INKA", 0, nroDocumento, tipoDocumento)
        Set obj = gclsOracle.FN_Cursor("ECVENTA.UNIFICADO_INKACLUB.CONSULTA_CLIENTE", 0, NroDocumento, TipoDocumento, "")
        While Not obj.EOF
            Set mvarAfiliado = New clsAfiliado
            mvarAfiliado.CodCliente = obj("CO_CLIENTE_TITULAR").Value
            mvarAfiliado.Nombre = obj("NO_CLIENTE").Value
            mvarAfiliado.ApParterno = obj("AP_PATERNO").Value
            mvarAfiliado.ApMarterno = obj("AP_MATERNO").Value
            mvarAfiliado.puntosDisponibles = obj("PUNTOS_DISPONIBLES").Value
            obj.MoveNext
        Wend
    End If
    Set ObtenerAfiliadoInka = mvarAfiliado
    gclsOracle.Cerrar 'Cierra Inka
    If MainX(bk_usuario, bk_password) = False Then Exit Function 'Abre RAC
    Exit Function
Control:
    gclsOracle.Cerrar 'Cierra Inka
    If MainX(bk_usuario, bk_password) = False Then Exit Function 'Abre RAC
    Set ObtenerAfiliadoInka = Nothing
End Function

Public Function ValidaValeFidInka(ByVal num As String) As String
    On Error GoTo Control
    gclsOracle.Cerrar 'Cierra RAC
    If MainInkaClub = False Then Exit Function 'Abre Inka
    Dim Bool As String
    Dim Bool2 As Boolean
    Dim msg As String
    Dim Vale As String
    If (num <> "" And frmPedido.pstrCodCliente_Ink <> "") Then
        Dim obj As oraDynaset
        Set obj = gclsOracle.FN_Cursor("ECVENTA.UNIFICADO_INKACLUB.LISTA_CHEQUE", 0, frmPedido.pstrCodCliente_Ink)
        While Not obj.EOF
'            Debug.Print obj("NU_CHEQUE").Value
'            Debug.Print right(obj("NU_CHEQUE").Value, 14)
'            Debug.Print Mid(obj("NU_CHEQUE").Value, 26)
'            Debug.Print num
            If num = Mid(obj("NU_CHEQUE").Value, 26) And Bool2 = False Then Bool2 = True: Vale = obj("NU_CHEQUE").Value
            obj.MoveNext
        Wend
        If Bool2 = True Then
            Bool = gclsOracle.FN_Valor("ECVENTA.UNIFICADO_INKACLUB.VALIDA_VALE", Vale, frmPedido.pstrCodCliente_Ink)
        Else
            Bool = "20009"
        End If
    End If
    If Bool = "20007" Then msg = "Vale no valido, ya ha sido canjeado."
    If Bool = "20008" Then msg = "Vale no valido, se encuentra anulado."
    If Bool = "20009" Then msg = "Vale no valido, es inexistente."
    If Bool = "20010" Then msg = "Cliente Invalido."
    If Bool = "20011" Then msg = "Vale no valido, se encuentra vencido."
    ValidaValeFidInka = IIf(Bool = "OK", Vale, "False")
    If ValidaValeFidInka = "False" Then MsgBox msg, vbCritical, App.ProductName
    gclsOracle.Cerrar 'Cierra Inka
    If MainX(bk_usuario, bk_password) = False Then Exit Function 'Abre RAC
    Exit Function
Control:
    gclsOracle.Cerrar 'Cierra Inka
    If MainX(bk_usuario, bk_password) = False Then Exit Function 'Abre RAC
    ValidaValeFidInka = "False"
End Function
Public Function ObtenerDatosAfiliado(ByVal NumeroTarjeta As String, _
                                     ByVal idEmpleado As String) As clsAfiliado
    Set ObtenerDatosAfiliado = ObtenerDatosAfiliadoSinTarjeta(NumeroTarjeta, idEmpleado)
End Function

Public Function ObtenerDatosAfiliadoSinTarjeta(ByVal NumeroTarjeta As String, _
                                               ByVal idEmpleado As String) As clsAfiliado
    Dim Tarjetas As New XArrayDB
    Dim oAffiliationSearch As New stAffiliationSearch
    Dim vtarjetas() As String
    Dim i As Integer
    
    On Error GoTo Control

    Set mvarAfiliado = New clsAfiliado

    If (NumeroTarjeta <> "" And idEmpleado <> "") Then
        Set oAffiliationSearch = _
            mvarWSCPuntos.getAffiliationSearch(mvarWSCPuntos.getFormatoDocIdentidad(NumeroTarjeta), _
                                               idEmpleado)
        If (Not (oAffiliationSearch.errorid = oFarmaPuntosConstante.EstadoTarjeta.INACTIVA Or _
                 oAffiliationSearch.errorid = oFarmaPuntosConstante.EstadoTarjeta.INVALIDA)) Then
            mvarAfiliado.Nombre = oAffiliationSearch.cardname
            mvarAfiliado.ApParterno = oAffiliationSearch.cardlastname
            mvarAfiliado.ApMarterno = oAffiliationSearch.cardlastname
            mvarAfiliado.TipoDni = mvarWSCPuntos.getTipoDocIdentidad(oAffiliationSearch.accountnumber)
            mvarAfiliado.DNI = mvarWSCPuntos.getFormatoDocIdentidadInversa(oAffiliationSearch.accountnumber)
            mvarAfiliado.FechaNacimiento = mvarWSCPuntos.getFormatearFechaFV(oAffiliationSearch.cardbirthdate)
            mvarAfiliado.Genero = IIf(oAffiliationSearch.cardgender <> "M", "F", "M")
            If (Not ("0" = oAffiliationSearch.cardemail)) Then
                mvarAfiliado.Email = oAffiliationSearch.cardemail
            End If
            If (Not ("0" = oAffiliationSearch.cardphone)) Then
                mvarAfiliado.Telefono = oAffiliationSearch.cardphone
            End If
            If (Not ("0" = oAffiliationSearch.cardcellphone)) Then
                mvarAfiliado.Celular = oAffiliationSearch.cardcellphone
            End If
            'IdTipoAvenida|Dirección|IdDepartamento|IdProvincia|IdDistrito|IdTipoDirección|Referencias
            If (oAffiliationSearch.cardaddress <> "") Then
                Dim a() As String, a_len As Byte
                a = Split(oAffiliationSearch.cardaddress, "|")
                a_len = UBound(a) - LBound(a) + 1
                If (a_len >= 1) Then
                    mvarAfiliado.TipoLugar = a(0)
                End If
                If (a_len >= 2) Then
                    mvarAfiliado.Direccion = a(1)
                End If
                If (a_len >= 3) Then
                    mvarAfiliado.Departamento = a(2)
                End If
                If (a_len >= 4) Then
                    mvarAfiliado.Provincia = a(3)
                End If
                If (a_len >= 5) Then
                    mvarAfiliado.Distrito = a(4)
                End If
                If (a_len >= 6) Then
                    mvarAfiliado.TipoDireccion = a(5)
                End If
                If (a_len >= 7) Then
                    mvarAfiliado.Referencias = a(6)
                End If
            End If

            If (oAffiliationSearch.cardrelated <> "") Then
                vtarjetas = Split(oAffiliationSearch.cardrelated, ",")
                mvarAfiliado.Tarjetas = LoadRows(vtarjetas)
            End If
        Else  ' si no lo encuenta en el WS envia NULL
            Set mvarAfiliado = Nothing 'New clsAfiliado
        End If
    End If

    Set ObtenerDatosAfiliadoSinTarjeta = mvarAfiliado
    Exit Function
Control:
    Set ObtenerDatosAfiliadoSinTarjeta = Nothing
End Function

Public Function ValidarTarjetaAsociada(ByVal NumeroTarjeta As String, _
                                       ByVal idEmpleado As String) As clsTarjetaBean
    Dim vTarjetaBean As New clsTarjetaBean
    Dim Lista() As String, n_lista As Integer
    Dim msj As String
    Dim oStTransactionInit As stTransactionInit
    Dim i As Integer
    Dim vEstado As String
    Dim listatmp() As String
    
    On Error GoTo Control
      
    Set oStTransactionInit = _
        mvarWSCPuntos.setTransactionInit(mvarWSCPuntos.getFormatoDocIdentidad(NumeroTarjeta), _
                                         idEmpleado)
    vEstado = oStTransactionInit.errorid

    If (oStTransactionInit.CardItems <> "") Then
        Lista = Split(oStTransactionInit.CardItems, "|")
        n_lista = UBound(Lista) - LBound(Lista) + 1
    End If

    'Inicializa el estado de tarjeta como SIN ESTADO
    vTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.SIN_ESTADO

    'validar estado tarjeta
    If (vEstado = oFarmaPuntosConstante.EstadoTarjeta.INVALIDA Or _
        vEstado = oFarmaPuntosConstante.EstadoTarjeta.INACTIVA Or _
        vEstado = oFarmaPuntosConstante.EstadoTarjeta.BLOQUEADA) Then
        vTarjetaBean.EstadoTarjeta = oStTransactionInit.errorid
    End If
    
    'Obtener si tarjeta esta bloqueada para redimir
    If (vEstado = oFarmaPuntosConstante.EXITO) Then
        If ("BLOCKEDREDEMPTION" = Lista(0)) Then
            vTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.BLOQUEADA_REDIMIR
        Else
            vTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.ACTIVA
        End If
    End If
            
    vTarjetaBean.NumeroTarjeta = oStTransactionInit.cardnumber
    vTarjetaBean.IdTransaccion = oStTransactionInit.TransactionId
    vTarjetaBean.DNI = mvarWSCPuntos.getFormatoDocIdentidadInversa(oStTransactionInit.accountnumber)
    vTarjetaBean.NombreCompleto = oStTransactionInit.cardname
    vTarjetaBean.PuntosTotalAcumulados = CDbl(val(oStTransactionInit.CardBalance))

    If (n_lista > 0) Then
    'productos inscritos
        If (n_lista >= 2) Then
            If (Lista(1) <> "" And Len(Lista(1)) > 0) Then
                listatmp = Split(Lista(1), ",")
                vTarjetaBean.ListaInscritos = LoadRows(listatmp)
            End If
        End If
        'Productos pendiente de canje
        If (n_lista >= 3) Then
            If (Lista(2) <> "" And Len(Lista(2)) > 0) Then
                listatmp = Split(Lista(2), ",")
                vTarjetaBean.ListaInscritos = LoadRows(listatmp)
            End If
        End If
        'Productos pendiente de canje
        If (n_lista >= 4) Then
            If (Lista(3) <> "" And Len(Lista(3)) > 0) Then
                vTarjetaBean.AhorroTotal = CDbl(val(Lista(3)))
            End If
        End If

        vTarjetaBean.EstadoOperacion = oStTransactionInit.errorid
        vTarjetaBean.Mensaje = oStTransactionInit.message
        vTarjetaBean.Fecha = oStTransactionInit.transactiondate

    End If
        
    Set ValidarTarjetaAsociada = vTarjetaBean
    Exit Function
Control:
    vTarjetaBean.IdTransaccion = "0"
    vTarjetaBean.NumeroTarjeta = NumeroTarjeta
    vTarjetaBean.EstadoOperacion = oFarmaPuntosConstante.NO_CONEXION_ORBIS
    vTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.SIN_ESTADO
    vTarjetaBean.Mensaje = Err.Description
    Set ValidarTarjetaAsociada = vTarjetaBean
End Function

Public Function GetEstadoTarjeta(ByVal NumeroTarjeta As String, _
                                 ByVal idEmpleado As String) As String
    Dim vEstadoTarjeta As String
    Dim Lista() As String
    Dim tmpT1 As Long, tmpT2 As Long
    Dim msj As String
    Dim oStTransactionInit As stTransactionInit
    Dim vEstado As String
    
    'tmpT1 = System.currentTimeMillis()
        
    On Error GoTo Control

    'tmpT1 = System.currentTimeMillis()
    Set oStTransactionInit = mvarWSCPuntos.setTransactionInit(mvarWSCPuntos.getFormatoDocIdentidad(NumeroTarjeta), _
                                                          idEmpleado)
    'tmpT2 = System.currentTimeMillis()
    'System.out.println("tiempo de demora 1 stTransactionQuote: "+(tmpT2-tmpT1))
    vEstado = oStTransactionInit.errorid

    If (oStTransactionInit.CardItems <> "") Then
        Lista = Split(oStTransactionInit.CardItems, "|")
    End If
            
    'Inicializa el estado de tarjeta como SIN ESTADO
    vEstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.SIN_ESTADO

    'validar estado tarjeta
    If (vEstado = oFarmaPuntosConstante.EstadoTarjeta.INVALIDA Or _
        vEstado = oFarmaPuntosConstante.EstadoTarjeta.INACTIVA Or _
        vEstado = oFarmaPuntosConstante.EstadoTarjeta.BLOQUEADA) Then
        vEstadoTarjeta = oStTransactionInit.errorid
    End If
    
    'Obtener si tarjeta esta bloqueada para redimir
    If (vEstado = oFarmaPuntosConstante.EXITO) Then
        If ("BLOCKEDREDEMPTION" = Lista(0)) Then
            vEstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.BLOQUEADA_REDIMIR
        Else
            vEstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.ACTIVA
        End If
    End If
        
    GetEstadoTarjeta = vEstadoTarjeta & "@" & _
                       mvarWSCPuntos.getFormatoDocIdentidadInversa(oStTransactionInit.accountnumber) & "@" & _
                       CDbl(val(oStTransactionInit.CardBalance))
    Exit Function
Control:
    Debug.Print Err.Description
    GetEstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.SIN_ESTADO & "@@0"
End Function

'''Public Function ObtenerCotizacion(ByVal listaProducto As XArrayDB, _
'''                                  ByVal idEmpleado As String) As clsTarjetaBean
'''    'no construye solo seteas los datos
'''    Dim vProductos As String
'''    Dim oStTransactionQuote As stTransactionQuote
'''    Dim vEstado As String
'''    Dim Lista() As String
'''
'''    On Error GoTo Control
'''
'''    vProductos = pGetCadenaDeList(listaProducto)
'''
'''    Set oStTransactionQuote = _
'''            mvarWSCPuntos.setTransactionQuote(mvarWSCPuntos.getFormatoDocIdentidad(mvarTarjetaBean.NumeroTarjeta), _
'''                                              idEmpleado, _
'''                                              mvarTarjetaBean.IdTransaccion, _
'''                                              vProductos)
'''
'''    vEstado = oStTransactionQuote.ErrorId
'''    If (oStTransactionQuote.TransactionItems <> "") Then
'''        Lista = Split(oStTransactionQuote.TransactionItems, "|")
'''        mvarTarjetaBean.ListaBonificados = LoadRows(Lista)
'''    End If
'''
'''    'Solo actualizar en el caso de Bloqueada, Invalida o inactiva,
'''    'porque el INI puede obtener el estado Bloqueda Redime
'''    If (vEstado = oFarmaPuntosConstante.EstadoTarjeta.INVALIDA Or _
'''        vEstado = oFarmaPuntosConstante.EstadoTarjeta.INACTIVA Or _
'''        vEstado = oFarmaPuntosConstante.EstadoTarjeta.BLOQUEADA) Then
'''        mvarTarjetaBean.EstadoTarjeta = oStTransactionQuote.ErrorId
'''    End If
'''
'''    'Bloqueada para redimir
'''    If (mvarTarjetaBean.EstadoTarjeta <> oFarmaPuntosConstante.EstadoTarjeta.BLOQUEADA_REDIMIR) Then
'''        mvarTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.ACTIVA
'''    End If
'''
'''    mvarTarjetaBean.IdTransaccion = oStTransactionQuote.TransactionId
'''    mvarTarjetaBean.NumeroTarjeta = oStTransactionQuote.CardNumber
'''    mvarTarjetaBean.DNI = mvarWSCPuntos.getFormatoDocIdentidadInversa(oStTransactionQuote.AccountNumber)
'''    mvarTarjetaBean.EstadoOperacion = oStTransactionQuote.ErrorId
'''    mvarTarjetaBean.Mensaje = oStTransactionQuote.Message
'''
'''    Set ObtenerCotizacion = mvarTarjetaBean
'''    Exit Function
'''Control:
'''    'no se setea el idTransaccion
'''    mvarTarjetaBean.EstadoOperacion = oFarmaPuntosConstante.NO_CONEXION_ORBIS
'''    mvarTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.SIN_ESTADO
'''    mvarTarjetaBean.Mensaje = Err.Description
'''    Set ObtenerCotizacion = mvarTarjetaBean
'''End Function

'''Public Function RechazarBonificado(ByVal listaProducto As XArrayDB, _
'''                                   ByVal idEmpleado As String) As clsTarjetaBean
'''    Dim vProductos As String
'''    Dim i As Integer
'''    Dim oStTransactionUpdate As stTransactionUpdate
'''    Dim Lista() As String
'''
'''    On Error GoTo Control
'''
'''    If (Not listaProducto Is Nothing) Then
'''        If (listaProducto.Size() > 0) Then
'''            For i = listaProducto.LowerBound(1) To listaProducto.UpperBound(1)
'''                If (i + 1 = listaProducto.UpperBound(1)) Then
'''                    vProductos = vProductos & listaProducto.Value(i, 0)
'''                Else
'''                    vProductos = vProductos & listaProducto.Value(i, 0) & "|"
'''                End If
'''            Next i
'''        End If
'''    End If
'''
'''    Set oStTransactionUpdate = _
'''        mvarWSCPuntos.setTransactionUpdate(mvarWSCPuntos.getFormatoDocIdentidad(mvarTarjetaBean.NumeroTarjeta), _
'''                                           idEmpleado, _
'''                                           mvarTarjetaBean.IdTransaccion, _
'''                                           vProductos)
'''
'''    mvarTarjetaBean.IdTransaccion = oStTransactionUpdate.TransactionId
'''    mvarTarjetaBean.EstadoOperacion = oStTransactionUpdate.ErrorId
'''    mvarTarjetaBean.Mensaje = oStTransactionUpdate.Message
'''
'''    If (oStTransactionUpdate.TransactionItems <> "") Then
'''        Lista = Split(oStTransactionUpdate.TransactionItems, "|")
'''        mvarTarjetaBean.ListaBonificados = LoadRows(Lista)
'''    End If
'''
'''    Set RechazarBonificado = mvarTarjetaBean
'''    Exit Function
'''Control:
'''    mvarTarjetaBean.EstadoOperacion = oFarmaPuntosConstante.NO_CONEXION_ORBIS
'''    mvarTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.SIN_ESTADO
'''    mvarTarjetaBean.Mensaje = Err.Description
'''    Set RechazarBonificado = mvarTarjetaBean
'''End Function

'''Public Function RegistrarVentaOnline(ByVal listaProducto As XArrayDB, _
'''                                     ByVal saldoRedimido As Double, _
'''                                     ByVal numeroPedido As String, _
'''                                     ByVal fechaPedido As Date, _
'''                                     ByVal importePedido As Double, _
'''                                     ByVal idEmpleado) As clsTarjetaBean
'''    Dim vProductos As String, vEstado As String
'''    Dim oStTransactionSale As stTransactionSale
'''
'''    On Error GoTo Control
'''
'''    vProductos = pGetCadenaDeList(listaProducto)
'''
'''    Set oStTransactionSale = _
'''        mvarWSCPuntos.setTransactionSale(mvarWSCPuntos.getFormatoDocIdentidad(mvarTarjetaBean.NumeroTarjeta), _
'''                                         idEmpleado, _
'''                                         mvarTarjetaBean.IdTransaccion, _
'''                                         vProductos, _
'''                                         saldoRedimido, _
'''                                         numeroPedido, _
'''                                         fechaPedido, _
'''                                         importePedido)
'''    vEstado = oStTransactionSale.ErrorId
'''
'''    'Inicializa el estado de tarjeta como SIN ESTADO
'''    mvarTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.SIN_ESTADO
'''
'''    'validar estado tarjeta
'''    If (vEstado = oFarmaPuntosConstante.EstadoTarjeta.INVALIDA Or _
'''        vEstado = oFarmaPuntosConstante.EstadoTarjeta.INACTIVA Or _
'''        vEstado = oFarmaPuntosConstante.EstadoTarjeta.BLOQUEADA) Then
'''        mvarTarjetaBean.EstadoTarjeta = oStTransactionSale.ErrorId
'''    End If
'''    mvarTarjetaBean.IdTransaccion = oStTransactionSale.TransactionId
'''    mvarTarjetaBean.NumeroAutororizacion = oStTransactionSale.SaleAuthNumber
'''    mvarTarjetaBean.PuntosTotalAcumulados = CDbl(Val(oStTransactionSale.CardBalance))
'''    mvarTarjetaBean.DNI = mvarWSCPuntos.getFormatoDocIdentidadInversa(oStTransactionSale.AccountNumber)
'''    mvarTarjetaBean.EstadoOperacion = oStTransactionSale.ErrorId
'''    mvarTarjetaBean.Mensaje = oStTransactionSale.Message
'''
'''    Set RegistrarVentaOnline = mvarTarjetaBean
'''    Exit Function
'''Control:
'''    mvarTarjetaBean.IdTransaccion = "0"
'''    mvarTarjetaBean.EstadoOperacion = oFarmaPuntosConstante.NO_CONEXION_ORBIS
'''    mvarTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.SIN_ESTADO
'''    mvarTarjetaBean.Mensaje = Err.Description
'''
'''    Set RegistrarVentaOnline = mvarTarjetaBean
'''End Function

Public Function RegistrarVentaOffline(ByVal NumeroTarjeta As String, _
                                      ByVal listaProducto As XArrayDB, _
                                      ByVal saldoRedimido As Double, _
                                      ByVal NumeroPedido As String, _
                                      ByVal fechaPedido As Date, _
                                      ByVal importePedido As Double, _
                                      ByVal idEmpleado As String) As clsTarjeta
    Dim vProductos As String, vEstado As String
    Dim vTarjetaBean As New clsTarjetaBean
    Dim oStTransactionSaleOffline As stTransactionSaleOffline
    
    On Error GoTo Control
    
    vProductos = pGetCadenaDeList(listaProducto)
    
    Set oStTransactionSaleOffline = _
        mvarWSCPuntos.setTransactionSaleOffline(mvarWSCPuntos.getFormatoDocIdentidad(NumeroTarjeta), _
                                                idEmpleado, _
                                                "", _
                                                vProductos, _
                                                saldoRedimido, _
                                                NumeroPedido, _
                                                fechaPedido, _
                                                importePedido)

    vEstado = oStTransactionSaleOffline.errorid
    
    If (vEstado = oFarmaPuntosConstante.EstadoTarjeta.INVALIDA Or _
        vEstado = oFarmaPuntosConstante.EstadoTarjeta.INACTIVA Or _
        vEstado = oFarmaPuntosConstante.EstadoTarjeta.BLOQUEADA) Then
        vTarjetaBean.EstadoTarjeta = vEstado
    End If
    
    vTarjetaBean.NumeroTarjeta = NumeroTarjeta
    vTarjetaBean.IdTransaccion = oStTransactionSaleOffline.TransactionId
    vTarjetaBean.NumeroAutororizacion = oStTransactionSaleOffline.saleauthnumber
    vTarjetaBean.PuntosTotalAcumulados = CDbl(val(oStTransactionSaleOffline.CardBalance))
    vTarjetaBean.EstadoOperacion = oStTransactionSaleOffline.errorid
    vTarjetaBean.Mensaje = oStTransactionSaleOffline.message
    vTarjetaBean.DNI = oStTransactionSaleOffline.accountnumber
    
    Set RegistrarVentaOffline = vTarjetaBean
    Exit Function
Control:
    vTarjetaBean.IdTransaccion = "0"
    vTarjetaBean.EstadoOperacion = oFarmaPuntosConstante.NO_CONEXION_ORBIS
    vTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.SIN_ESTADO
    vTarjetaBean.Mensaje = Err.Description
    
    Set RegistrarVentaOffline = vTarjetaBean
End Function

Public Function RecuperarPuntos(ByVal NumeroTarjeta As String, _
                                ByVal listaProducto As XArrayDB, _
                                ByVal NumeroPedido As String, _
                                ByVal fechaPedido As Date, _
                                ByVal importePedido As Double, _
                                ByVal idEmpleado As String) As clsTarjetaBean
    
    'se enviara con el registra venta offline en el mismo momento
    Set RecuperarPuntos = Me.RegistrarVentaOffline(NumeroTarjeta, _
                                                   listaProducto, _
                                                   0#, _
                                                   NumeroPedido, _
                                                   fechaPedido, _
                                                   importePedido, _
                                                   idEmpleado)

End Function

'''Public Function ValidarAnulacion(ByVal NumeroTarjeta As String, _
'''                                 ByVal IdTransaccion As String, _
'''                                 ByVal numeroAutorizacion As String, _
'''                                 ByVal listaArticulo As XArrayDB, _
'''                                 ByVal idEmpleado As String) As clsTarjetaBean
'''    Dim vProductos As String, vEstado As String
'''    Dim vTarjetaBean As New clsTarjetaBean
'''    Dim oStcheckSaleReturn As stcheckSaleReturn
'''
'''    vProductos = pGetCadenaDeList(listaArticulo)
'''
'''    On Error GoTo Control
'''
'''    Set oStcheckSaleReturn = _
'''        mvarWSCPuntos.checkSaleReturn(mvarWSCPuntos.getFormatoDocIdentidad(NumeroTarjeta), _
'''                                      idEmpleado, _
'''                                      IdTransaccion, _
'''                                      numeroAutorizacion, _
'''                                      vProductos)
'''
'''    'Inicializa el estado de tarjeta como SIN ESTADO
'''    vTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.SIN_ESTADO
'''    vEstado = oStcheckSaleReturn.ErrorId
'''
'''    'validar estado tarjeta
'''    If (vEstado = oFarmaPuntosConstante.EstadoTarjeta.INVALIDA Or _
'''        vEstado = oFarmaPuntosConstante.EstadoTarjeta.INACTIVA Or _
'''        vEstado = oFarmaPuntosConstante.EstadoTarjeta.BLOQUEADA) Then
'''        vTarjetaBean.EstadoTarjeta = oStcheckSaleReturn.ErrorId
'''    Else
'''        If (vEstado = oFarmaPuntosConstante.EXITO) Then
'''            vTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.ACTIVA
'''        End If
'''    End If
'''
'''    vTarjetaBean.IdTransaccion = oStcheckSaleReturn.TransactionId
'''    vTarjetaBean.Mensaje = oStcheckSaleReturn.Message
'''    vTarjetaBean.EstadoOperacion = oStcheckSaleReturn.ErrorId
'''
'''    Set ValidarAnulacion = vTarjetaBean
'''    Exit Function
'''Control:
'''    vTarjetaBean.IdTransaccion = "0"
'''    vTarjetaBean.EstadoOperacion = oFarmaPuntosConstante.NO_CONEXION_ORBIS
'''    vTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.SIN_ESTADO
'''    vTarjetaBean.Mensaje = Err.Description
'''
'''    Set ValidarAnulacion = vTarjetaBean
'''End Function

'''Public Function AnularPedido(ByVal NumeroTarjeta As String, _
'''                             ByVal listaArticulo As XArrayDB, _
'''                             ByVal numeroPedido As String, _
'''                             ByVal saldoRedimido As Double, _
'''                             ByVal fechaPedido As Date, _
'''                             ByVal importePedido As Double, _
'''                             ByVal idEmpleado As String) As clsTarjetaBean
'''    Dim vProductos As String, vEstado As String
'''    Dim oStSaleReturn As stSaleReturn
'''
'''    vProductos = pGetCadenaDeList(listaArticulo)
'''
'''    On Error GoTo Control
'''
'''    Set oStSaleReturn = _
'''        mvarWSCPuntos.setSaleReturn(mvarWSCPuntos.getFormatoDocIdentidad(NumeroTarjeta), _
'''                                    idEmpleado, _
'''                                    mvarTarjetaBean.IdTransaccion, _
'''                                    vProductos, _
'''                                    saldoRedimido, _
'''                                    numeroPedido, _
'''                                    fechaPedido, _
'''                                    importePedido)
'''
'''    'Inicializa el estado de tarjeta como SIN ESTADO
'''    mvarTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.SIN_ESTADO
'''    vEstado = oStSaleReturn.ErrorId
'''
'''    'validar estado tarjeta
'''    If (vEstado = oFarmaPuntosConstante.EstadoTarjeta.INVALIDA Or _
'''        vEstado = oFarmaPuntosConstante.EstadoTarjeta.INACTIVA Or _
'''        vEstado = oFarmaPuntosConstante.EstadoTarjeta.BLOQUEADA) Then
'''        mvarTarjetaBean.EstadoTarjeta = oStSaleReturn.ErrorId
'''    Else
'''        If (vEstado = oFarmaPuntosConstante.EXITO) Then
'''            mvarTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.ACTIVA
'''        End If
'''    End If
'''
'''    mvarTarjetaBean.Mensaje = oStSaleReturn.Message
'''    mvarTarjetaBean.EstadoOperacion = oStSaleReturn.ErrorId
'''    mvarTarjetaBean.NumeroAutororizacion = oStSaleReturn.SaleAuthNumber
'''    mvarTarjetaBean.IdTransaccion = oStSaleReturn.TransactionId
'''
'''    Set AnularPedido = mvarTarjetaBean
'''    Exit Function
'''Control:
'''    mvarTarjetaBean.IdTransaccion = "0"
'''    mvarTarjetaBean.EstadoOperacion = oFarmaPuntosConstante.NO_CONEXION_ORBIS
'''    mvarTarjetaBean.EstadoTarjeta = oFarmaPuntosConstante.EstadoTarjeta.SIN_ESTADO
'''    mvarTarjetaBean.Mensaje = Err.Description
'''
'''    Set AnularPedido = mvarTarjetaBean
'''End Function

Public Function ConsultarSaldo(ByVal NumeroTarjeta As String, _
                               ByVal idEmpleado As String) As clsTarjetaBean

    Set ConsultarSaldo = Me.ValidarTarjetaAsociada(NumeroTarjeta, idEmpleado)

End Function

''''    public function logeoUsuario(String usuario, String clave) as clsAfiliado
''''        logeoUsuario = mvarAfiliado
''''    end function
''''
''''    Public Property Get IdLocal() as String
''''        IdLocal = mvaridLocal
''''    end property
''''
''''    Public Property Get IdCaja() as String
''''        IdCaja = mvaridCaja
''''    end property
''''
''''    Public Property Get IdEmpleado() as string
''''        IdEmpleado = mvaridEmpleado
''''    end property
''''
''''    Public Property Get CodigoAutorizacion() as string
''''        CodigoAutorizacion = mvarcodigoAutorizacion
''''    end property
