VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vIndice As Integer
Dim xArrayProducto As New XArrayDB
Private m_strPedido As String
Private m_strCodProducto As String
Private m_strDescripcion As String
Private m_strUnidades As String
Private m_strFracciones As String
Private m_strLote As String
Private m_strFchVenc As String
Private m_strStock As String
Private m_strCtdPedido As String
Private m_strEsFraccion As String
Private m_strCtdFraccion As String
Private m_strCtdPendiente As String
Private m_strCtdAtendido As String

Private m_strCtdUnd As String
Private m_strCtdFrac As String

Public Property Get varrIndice() As Integer
    varrIndice = vIndice
End Property

Public Property Let varrIndice(ByVal vNewValue As Integer)
    vIndice = vNewValue
End Property

Public Property Get Pedido() As String
    Pedido = m_strPedido
End Property

Public Property Let Pedido(ByVal StrPedido As String)
    m_strPedido = StrPedido
End Property

Public Property Get Producto() As String
    Producto = m_strCodProducto
End Property

Public Property Let Producto(ByVal strProducto As String)
    m_strCodProducto = strProducto
End Property

Public Property Get Descripcion() As String
    Descripcion = m_strDescripcion
End Property

Public Property Let Descripcion(ByVal strDescripcion As String)
    m_strDescripcion = strDescripcion
End Property

Public Property Get Unidades() As String
    Unidades = m_strUnidades
End Property

Public Property Let Unidades(ByVal strUnidades As String)
    m_strUnidades = strUnidades
End Property

Public Property Get Fracciones() As String
    Fracciones = m_strFracciones
End Property

Public Property Let Fracciones(ByVal strFracciones As String)
    m_strFracciones = strFracciones
End Property

Public Property Get Lote() As String
    Lote = m_strLote
End Property

Public Property Let Lote(ByVal strLote As String)
    m_strLote = strLote
End Property

Public Property Get FchVenc() As String
    FchVenc = m_strFchVenc
End Property

Public Property Let FchVenc(ByVal strFchVenc As String)
    m_strFchVenc = strFchVenc
End Property

Public Property Get Stock() As String
    Stock = m_strStock
End Property

Public Property Let Stock(ByVal strStock As String)
    m_strStock = strStock
End Property

Public Property Get CtdPedido() As String
    CtdPedido = m_strCtdPedido
End Property

Public Property Let CtdPedido(ByVal strCtdPedido As String)
    m_strCtdPedido = strCtdPedido
End Property

Public Property Get EsFraccion() As String
    EsFraccion = m_strEsFraccion
End Property

Public Property Let EsFraccion(ByVal strEsFraccion As String)
    m_strEsFraccion = strEsFraccion
End Property

Public Property Get CtdFraccion() As String
    CtdFraccion = m_strCtdFraccion
End Property

Public Property Let CtdFraccion(ByVal strCtdFraccion As String)
    m_strCtdFraccion = strCtdFraccion
End Property

Public Property Get CtdPendiente() As String
    CtdPendiente = m_strCtdPendiente
End Property

Public Property Let CtdPendiente(ByVal strCtdPendiente As String)
    m_strCtdPendiente = strCtdPendiente
End Property

Public Property Get CtdAtendido() As String
    CtdAtendido = m_strCtdAtendido
End Property

Public Property Let CtdAtendido(ByVal strCtdAtendido As String)
    m_strCtdAtendido = strCtdAtendido
End Property

Public Property Get ctdUnd() As String
    ctdUnd = m_strCtdUnd
End Property

Public Property Let ctdUnd(ByVal strCtdUnd As String)
    m_strCtdUnd = strCtdUnd
End Property

Public Property Get ctdFrac() As String
    ctdFrac = m_strCtdFrac
End Property

Public Property Let ctdFrac(ByVal strCtdFrac As String)
    m_strCtdFrac = strCtdFrac
End Property

Public Function ListaEstado(ByVal CodigoEstado As String, ByVal Texto As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaEstado = gclsOracle.FN_Cursor("BTLCERO.PKG_PEDIDO.FN_LISTA_ESTADOS", 0, CodigoEstado, Texto)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsPedido.ListaEstado", Err.Description
End Function

Public Function Lista(ByVal strLocal As String, _
                      ByVal StrPedido As String, _
                      ByVal strFchInicio As String, _
                      ByVal strFchFin As String, _
                      ByVal strEstado As String, _
                      ByVal strCliente As String) As oraDynaset
On Error GoTo CtrlErr
      Set Lista = gclsOracle.FN_Cursor("BTLCERO.PKG_PEDIDO.FN_LISTA_PEDIDOS_LOCAL", 0, strLocal, StrPedido, _
                                        strEstado, strFchInicio, strFchFin, strCliente)
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsPedido.Lista", Err.Description
End Function

Public Function Cabecera_Pedido(ByVal strLocal As String, _
                                ByVal StrPedido As String) As oraDynaset
On Error GoTo CtrlErr
      Set Cabecera_Pedido = gclsOracle.FN_Cursor("BTLCERO.PKG_PEDIDO.FN_DEV_CABECERA_PEDIDO", 0, strLocal, StrPedido)
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsPedido.Cabecera_Pedido", Err.Description
End Function

Public Function Detalle_Pedido(ByVal strLocal As String, _
                               ByVal StrPedido As String) As oraDynaset
On Error GoTo CtrlErr
      Set Detalle_Pedido = gclsOracle.FN_Cursor("BTLCERO.PKG_PEDIDO.FN_LISTA_DETALLE", 0, strLocal, StrPedido)
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsPedido.Detalle_Pedido", Err.Description
End Function

Public Function GrabaPedido(ByRef RetNumDocumento As String, ByRef RetImpTotal As Double, _
                            ByVal vstrCia As String, ByVal vstrCodLocal As String, _
                            ByVal vstrCodMaquina As String, ByVal vstrCodUsuario As String, _
                            ByVal vstrNumPedido As String, ByVal vstrCodLiquidacion As String, _
                            ByVal varrProducto As Variant, ByVal varrUnidades As Variant, ByVal varrFracciones As Variant, _
                            ByVal varrLotes As Variant, ByVal varrFchVcto As Variant, ByVal StrDireccionEntrega As String, _
                            ByVal strOC As String, ByVal strPoliza As String, ByVal strNombreClienteAux As String) As String
On Error GoTo CtrlErr

    Dim lstrSql As String
    Dim i As Integer
    Dim lintFilas As Integer

    Dim lparProducto As OracleInProcServer.OraParamArray
    Dim lparUnidades As OracleInProcServer.OraParamArray
    Dim lparFracciones As OracleInProcServer.OraParamArray
    Dim lparLotes As OracleInProcServer.OraParamArray
    Dim lparFchVcto As OracleInProcServer.OraParamArray
    Dim lparSubTotal As OracleInProcServer.OraParamArray

    Dim arrUnidades As Variant
    Dim arrFracciones As Variant
    Dim arrLotes As Variant
    Dim arrFVencimiento As Variant
    Dim arrUnidadesX As Variant
    Dim o As Integer
    Dim ox As Integer
    lintFilas = UBound(varrProducto) - LBound(varrProducto) + 1
    gclsOracle.LimpiaParametros



    lintFilas = 0
    For i = 0 To xArrayProducto.Count(1) - 1
            If Not Trim(xArrayProducto(i, 13) & xArrayProducto(i, 14)) = "" Then
    
            arrUnidadesX = Split(xArrayProducto(i, 13), "|")
            
            ox = 0
            
            While ox < UBound(arrUnidadesX)
                lintFilas = lintFilas + 1
                ox = ox + 1
            Wend
            End If
    Next i
    
    If lintFilas = 0 Then
        GrabaPedido = "No existen productos Despachados"
        Exit Function
    End If
    
    With gclsOracle.ODataBase.Parameters
        .AddTable "A_CAD_COD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas, 10
        .AddTable "A_CAD_CTD_UND_ATE", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas, 10
        .AddTable "A_CAD_CTD_FRAC_ATE", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas, 10
        .AddTable "A_CAD_NUM_LOTE", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas, 10
        .AddTable "A_CAD_FCH_VCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas, 10
        .AddTable "A_CAD_IMP_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas, 10
    End With

    Set lparProducto = gclsOracle.ODataBase.Parameters("A_CAD_COD_PRODUCTO")
    Set lparUnidades = gclsOracle.ODataBase.Parameters("A_CAD_CTD_UND_ATE")
    Set lparFracciones = gclsOracle.ODataBase.Parameters("A_CAD_CTD_FRAC_ATE")
    Set lparLotes = gclsOracle.ODataBase.Parameters("A_CAD_NUM_LOTE")
    Set lparFchVcto = gclsOracle.ODataBase.Parameters("A_CAD_FCH_VCTO")
    Set lparSubTotal = gclsOracle.ODataBase.Parameters("A_CAD_IMP_PRODUCTO")
    
    Dim k As Integer
    k = 0
    For i = 0 To xArrayProducto.Count(1) - 1
        If Not Trim(xArrayProducto(i, 13) & xArrayProducto(i, 14)) = "" Then

        arrUnidades = Split(xArrayProducto(i, 13), "|")
        arrFracciones = Split(xArrayProducto(i, 14), "|")
        arrLotes = Split(xArrayProducto(i, 15), "|")
        arrFVencimiento = Split(xArrayProducto(i, 16), "|")
        
        o = 0
       
        While o < UBound(arrUnidades)
            lparProducto(k) = xArrayProducto(i, 0)
            lparUnidades(k) = arrUnidades(o)
            lparFracciones(k) = arrFracciones(o)
            lparLotes(k) = arrLotes(o)
            lparFchVcto(k) = arrFVencimiento(o)
            lparSubTotal(k) = xArrayProducto(i, 23)
            k = k + 1
            o = o + 1
        Wend
        End If
    Next i
    

    With gclsOracle.ODataBase.Parameters
         .Add "RET_NUM_GUIA", RetNumDocumento, ORAPARM_OUTPUT, ORATYPE_VARCHAR2
         .Add "RET_IMP_TOTAL", RetImpTotal, ORAPARM_OUTPUT, ORATYPE_FLOAT
         .Add "A_CIA", vstrCia, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_COD_LOCAL", vstrCodLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_COD_MAQUINA", vstrCodMaquina, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_COD_USUARIO", vstrCodUsuario, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_NUM_PEDIDO", vstrNumPedido, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_COD_LIQUIDACION", vstrCodLiquidacion, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_DES_DIRECCION_CLI", StrDireccionEntrega, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_NUM_ORDCOM", strOC, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_cod_poliza", strPoliza, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_nom_cliente_poliza", strNombreClienteAux, ORAPARM_INPUT, ORATYPE_VARCHAR2
    End With

    lstrSql = "BEGIN " & _
              "BTLCERO.PKG_PEDIDO.SP_GRABA_GUIA(RET_NUM_GUIA         => :RET_NUM_GUIA," & _
              "                                 RET_IMP_TOTAL        => :RET_IMP_TOTAL," & _
              "                                 A_CIA                => :A_CIA," & _
              "                                 A_COD_LOCAL          => :A_COD_LOCAL," & _
              "                                 A_COD_MAQUINA        => :A_COD_MAQUINA," & _
              "                                 A_COD_USUARIO        => :A_COD_USUARIO," & _
              "                            A_CAD_COD_PRODUCTO   => :A_CAD_COD_PRODUCTO," & _
              "                            A_CAD_CTD_UND_ATE    => :A_CAD_CTD_UND_ATE," & _
              "                            A_CAD_CTD_FRAC_ATE   => :A_CAD_CTD_FRAC_ATE," & _
              "                            A_CAD_NUM_LOTE       => :A_CAD_NUM_LOTE," & _
              "                            A_CAD_FCH_VCTO       => :A_CAD_FCH_VCTO," & _
              "                            A_CAD_IMP_PRODUCTO       => :A_CAD_IMP_PRODUCTO," & _
              "                                 A_NUM_PEDIDO         => :A_NUM_PEDIDO," & _
              "                            A_DES_DIRECCION_CLI       => :A_DES_DIRECCION_CLI," & _
              "                                 A_COD_LIQUIDACION    => :A_COD_LIQUIDACION," & _
              "                                 A_NUM_ORDCOM         => :A_NUM_ORDCOM," & _
              "                                 A_cod_poliza    => :A_cod_poliza," & _
              "                                 A_nom_cliente_poliza    => :A_nom_cliente_poliza);" & _
              "END;"

    gclsOracle.ODataBase.ExecuteSQL (lstrSql)

    RetNumDocumento = Trim("" & gclsOracle.ODataBase.Parameters("RET_NUM_GUIA").Value)

    RetImpTotal = Val(Trim("" & gclsOracle.ODataBase.Parameters("RET_IMP_TOTAL").Value))
    
   Exit Function

CtrlErr:
    Err.Raise Err.Number, "clsPedido.GrabaPedido", Err.Description
End Function

Public Function fn_Estan_cantidades(ByVal flgFraccion As String, ByVal strCtdUnd As Long, _
                                     ByVal strCtdFrac As Long, ByVal strCtdFraccion As Integer) As Long
On Error GoTo Control
    
    If flgFraccion = "1" Then
       fn_Estan_cantidades = strCtdFrac + (strCtdUnd * strCtdFraccion)
    Else
       fn_Estan_cantidades = strCtdUnd
    End If

   Exit Function

Control:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error : " & Err.Number
End Function



Public Function fnConvierte(ByVal flgFraccion As String, ByVal strCtdUnd As Long, _
                                     ByVal strCtdFrac As Long, ByVal strCtdFraccion As Integer) As String
On Error GoTo Control
    Dim Cantidad As Long
    If flgFraccion = "1" Then
       Cantidad = strCtdFrac + (strCtdUnd * strCtdFraccion)
       fnConvierte = (Cantidad \ strCtdFraccion) & IIf((Cantidad Mod strCtdFraccion) = 0, "", "F" & (Cantidad Mod strCtdFraccion))
    Else
       fnConvierte = strCtdUnd
    End If

   Exit Function

Control:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error : " & Err.Number
End Function
Private Sub Class_Initialize()
    xArrayProducto.ReDim 0, -1, 0, 25
End Sub
Public Function AddProducto(ByVal Codigo As String, ByVal Descripcion As String, ByVal Laboratorio As String, _
                            ByVal Linea As String, ByVal Pedido As Long, ByVal Atendido As Long, _
                            ByVal Stock As String, ByVal Observacion As String, ByVal flgFraccion As String, _
                            ByVal CantidadFraccionamiento As String, ByVal Precio As String)
    Dim i As Integer
    xArrayProducto.AppendRows
    i = xArrayProducto.Count(1) - 1
    xArrayProducto(i, 0) = Codigo
    xArrayProducto(i, 1) = Descripcion
    xArrayProducto(i, 2) = Laboratorio
    xArrayProducto(i, 3) = Linea
    xArrayProducto(i, 4) = Me.fnConvierte(flgFraccion, IIf(flgFraccion = "0", Pedido, 0), IIf(flgFraccion = "1", Pedido, 0), CantidadFraccionamiento)
    xArrayProducto(i, 5) = Me.fnConvierte(flgFraccion, IIf(flgFraccion = "0", Atendido, 0), IIf(flgFraccion = "1", Atendido, 0), CantidadFraccionamiento)
    xArrayProducto(i, 6) = Stock
    xArrayProducto(i, 7) = Observacion
    xArrayProducto(i, 8) = Pedido ''para el calculo
    xArrayProducto(i, 9) = Atendido ''para el calculo
    xArrayProducto(i, 10) = (Val(fncGuion(Stock, 0, "F")) * CantidadFraccionamiento) + Val(IIf(InStr(1, Stock, "F") = 0, 0, fncGuion(Stock, 1, "F"))) ''para el calculo
    xArrayProducto(i, 11) = flgFraccion ''
    xArrayProducto(i, 12) = CantidadFraccionamiento ''
    xArrayProducto(i, 13) = "" ''Cantidad unidades
    xArrayProducto(i, 14) = "" ''Cantidad fracciones
    xArrayProducto(i, 15) = "" ''Lote
    xArrayProducto(i, 16) = "" ''FechaVencimiento
    xArrayProducto(i, 17) = "" ''real Cantidad unidades
    xArrayProducto(i, 18) = "" ''real Cantidad fracciones
    xArrayProducto(i, 19) = "" ''real Lote
    xArrayProducto(i, 20) = "" ''real FechaVencimiento
    xArrayProducto(i, 21) = Atendido ''para el calculo original
    xArrayProducto(i, 22) = (Val(fncGuion(Stock, 0, "F")) * CantidadFraccionamiento) + Val(IIf(InStr(1, Stock, "F") = 0, 0, fncGuion(Stock, 1, "F"))) ''para el calculo original
    xArrayProducto(i, 23) = Precio
     xArrayProducto(i, 24) = Atendido 'real atendido
     xArrayProducto(i, 25) = CStr((Val(fncGuion(Stock, 0, "F")) * CantidadFraccionamiento) + Val(IIf(InStr(1, Stock, "F") = 0, 0, fncGuion(Stock, 1, "F")))) 'real stock
End Function

Public Function addAtendido(ByVal Codigo As String, ByVal CtdUnidades As String, ByVal Ctdfracciones As String, ByVal NumLote As String, FchVencimiento As String, Optional ByVal Elimina As Boolean = False, Optional ByVal indice As String = "")
    Dim i As Long
    Dim Atendiendo As Double
    Dim AtendidoNew As Double
    Dim arrUnidades As Variant
    Dim arrFracciones As Variant

    i = xArrayProducto.Find(0, 0, Codigo, XORDER_ASCEND)
    ''validaciones
    Atendiendo = fn_Estan_cantidades(xArrayProducto(i, 11), CtdUnidades, Ctdfracciones, xArrayProducto(i, 12))
    If Elimina = False Then
        If Val(xArrayProducto(i, 10)) = 0 Then
            MsgBox "No hay Productos en Stock", vbCritical, App.ProductName
            Exit Function
        End If
    
        If Atendiendo = 0 Then
            MsgBox "La cantidad atendida no puede ser cero", vbCritical, App.ProductName
            Exit Function
        End If
    
        If Val(xArrayProducto(i, 10)) - Val(Atendiendo) < 0 Then
            MsgBox "La cantidad solicitada es mayor al stock", vbCritical, App.ProductName
            Exit Function
        End If
        If (Val(xArrayProducto(i, 9)) + Val(Atendiendo)) > Val(xArrayProducto(i, 8)) Then
            MsgBox "La cantidad atendida excede al pedido", vbCritical, App.ProductName
            Exit Function
        End If

        xArrayProducto(i, 13) = xArrayProducto(i, 13) & CtdUnidades & "|"  ''Cantidad unidades
        xArrayProducto(i, 14) = xArrayProducto(i, 14) & Ctdfracciones & "|"    ''Cantidad fracciones
        xArrayProducto(i, 15) = xArrayProducto(i, 15) & NumLote & "|"  ''Lote
        xArrayProducto(i, 16) = xArrayProducto(i, 16) & FchVencimiento & "|"  ''FechaVencimiento
    Else
        Dim arrLote As Variant
        Dim arrFchVencimiento As Variant
        Dim xTemp As New XArrayDB
        xTemp.ReDim 0, -1, 0, 5
        arrUnidades = Split(xArrayProducto(i, 13), "|")
        arrFracciones = Split(xArrayProducto(i, 14), "|")
        arrLote = Split(xArrayProducto(i, 15), "|")
        arrFchVencimiento = Split(xArrayProducto(i, 16), "|")
        Dim f As Integer
        While (f < UBound(arrUnidades))
            xTemp.AppendRows
            xTemp(f, 0) = arrUnidades(f)
            xTemp(f, 1) = arrFracciones(f)
            xTemp(f, 2) = arrLote(f)
            xTemp(f, 3) = arrFchVencimiento(f)
            f = f + 1
        Wend
        xTemp.Delete 1, indice
        f = 0
        xArrayProducto(i, 13) = ""
        xArrayProducto(i, 14) = ""
        xArrayProducto(i, 15) = ""
        xArrayProducto(i, 16) = ""
        
        While f < xTemp.Count(1)
            
        xArrayProducto(i, 13) = xArrayProducto(i, 13) & xTemp(f, 0) & "|"
        xArrayProducto(i, 14) = xArrayProducto(i, 14) & xTemp(f, 1) & "|"
        xArrayProducto(i, 15) = xArrayProducto(i, 15) & xTemp(f, 2) & "|"
        xArrayProducto(i, 16) = xArrayProducto(i, 16) & xTemp(f, 3) & "|"
            
            f = f + 1
        Wend
    End If
    Dim u As Integer
    
        arrUnidades = Split(xArrayProducto(i, 13), "|")
        arrFracciones = Split(xArrayProducto(i, 14), "|")
        AtendidoNew = 0
        While u < UBound(arrUnidades)
            AtendidoNew = AtendidoNew + Val((arrUnidades(u) * xArrayProducto(i, 12)) + arrFracciones(u))
            u = u + 1
        Wend
    
    Atendiendo = AtendidoNew
    
    xArrayProducto(i, 9) = Val(xArrayProducto(i, 21)) + Atendiendo
    xArrayProducto(i, 10) = Val(xArrayProducto(i, 22)) - Atendiendo
    
    xArrayProducto(i, 5) = Me.fnConvierte(xArrayProducto(i, 11), IIf(xArrayProducto(i, 11) = "0", xArrayProducto(i, 9), 0), IIf(xArrayProducto(i, 11) = "1", xArrayProducto(i, 9), 0), xArrayProducto(i, 12))
    xArrayProducto(i, 6) = Me.fnConvierte(xArrayProducto(i, 11), IIf(xArrayProducto(i, 11) = "0", xArrayProducto(i, 10), 0), IIf(xArrayProducto(i, 11) = "1", xArrayProducto(i, 10), 0), xArrayProducto(i, 12))
    
End Function

Public Property Get ListaProducto() As XArrayDB
    Set ListaProducto = xArrayProducto 'Devuelve la variable de la clase en forma de XArray
End Property


Public Function GrabaDetalle(ByVal indice As Integer)
    xArrayProducto(indice, 17) = xArrayProducto(indice, 13)
    xArrayProducto(indice, 18) = xArrayProducto(indice, 14)
    xArrayProducto(indice, 19) = xArrayProducto(indice, 15)
    xArrayProducto(indice, 20) = xArrayProducto(indice, 16)
    xArrayProducto(indice, 24) = xArrayProducto(indice, 9)
    xArrayProducto(indice, 25) = xArrayProducto(indice, 10)
End Function

Public Function CancelaDetalle(ByVal indice As Integer)
    xArrayProducto(indice, 13) = xArrayProducto(indice, 17)
    xArrayProducto(indice, 14) = xArrayProducto(indice, 18)
    xArrayProducto(indice, 15) = xArrayProducto(indice, 19)
    xArrayProducto(indice, 16) = xArrayProducto(indice, 20)
    xArrayProducto(indice, 9) = Val(xArrayProducto(indice, 24))
    xArrayProducto(indice, 10) = Val(xArrayProducto(indice, 25))
    xArrayProducto(indice, 5) = Me.fnConvierte(xArrayProducto(indice, 11), IIf(xArrayProducto(indice, 11) = "0", xArrayProducto(indice, 9), 0), IIf(xArrayProducto(indice, 11) = "1", xArrayProducto(indice, 9), 0), xArrayProducto(indice, 12))
    xArrayProducto(indice, 6) = Me.fnConvierte(xArrayProducto(indice, 11), IIf(xArrayProducto(indice, 11) = "0", xArrayProducto(indice, 10), 0), IIf(xArrayProducto(indice, 11) = "1", xArrayProducto(indice, 10), 0), xArrayProducto(indice, 12))
End Function
Function listaGuias(ByVal NumeroPedido As String, ByVal EstadoGuia As String, ByVal strLocal As String)

On Error GoTo CtrlErr
      Set listaGuias = gclsOracle.FN_Cursor("BTLCERO.PKG_PEDIDO.FN_LISTA_GUIAS", 0, NumeroPedido, EstadoGuia, strLocal)
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsPedido.ListaGuias", Err.Description

End Function

Public Function GeneraFactura(ByVal StrPedido As String, _
                      ByVal strLocal As String, _
                      ByVal strUsuario As String, _
                      ByVal CadNumGuia As String, _
                      ByVal strMaquina As String, _
                      ByVal FlgGuia As String, _
                      ByRef strTipoDoc As String, _
                      ByRef strNumDoc As String, _
                      ByVal strConVenta As String) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant

    gvarValores = Array(StrPedido, _
                         strLocal, _
                         strUsuario, _
                         CadNumGuia, _
                         strMaquina, _
                         FlgGuia, _
                         strTipoDoc, _
                         strNumDoc, _
                         strConVenta)

    gvarIO = Array(entrada, _
                    entrada, _
                    entrada, _
                    entrada, _
                    entrada, _
                    entrada, _
                    Entrada_Salida, _
                    Salida, _
                    entrada)
                    
    GeneraFactura = gclsOracle.SP("BTLCERO.PKG_PEDIDO.SP_GRABA_DOCUMENTO", _
                              gvarValores, _
                              gvarIO)
strTipoDoc = gvarValores(6)
strNumDoc = gvarValores(7)
    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsPedido.GeneraFactura", Err.Description
End Function


Public Function Dev_CondVenta(ByVal vstrDatos) As oraDynaset
    On Error GoTo CtrlErr
        Set Dev_CondVenta = gclsOracle.FN_Cursor("BTLCERO.PKG_DOC.FN_DEV_CONVTA", 0, vstrDatos)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento::Dev_CondVenta", Err.Description
End Function

Public Function ListaCondVenta() As oraDynaset
    On Error GoTo CtrlErr
        Set ListaCondVenta = gclsOracle.FN_Cursor("BTLCERO.PKG_DOC.FN_LISTA_FORPAG", 0)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento::ListaCondVenta", Err.Description
End Function


Public Function DevuelveIPServer(ByVal CodigoLocal As String) As String
    On Error GoTo CtrlError
Dim StrCodigoZona As String
StrCodigoZona = gclsOracle.FN_Valor("BTLPROD.FN_DEV_IP_FV", CodigoLocal)
DevuelveIPServer = StrCodigoZona
Exit Function
CtrlError:
    Err.Raise Err.Number, "clsZona.DevuelveIPServer", Err.Description
End Function

