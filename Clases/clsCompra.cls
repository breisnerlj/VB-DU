VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Lista(ByVal pstrLocal As String, ByVal pstrDesde As String, ByVal pstrHasta As String, Optional ByVal pstrProveedor As String) As oraDynaset
    On Error GoTo pase
    Set Lista = gclsOracle.FN_Cursor("NUEVO.PKG_DOC_PROV.FN_RECEPCIONADOS", 0, pstrLocal, pstrDesde, pstrHasta, pstrProveedor)
    Exit Function
pase:
    Err.Raise Err.Number, "clsCompra", Err.Description
End Function

Public Function OrdenCompra(ByVal pstrOCompra As String, ByVal pstrLocal As String) As oraDynaset
On Error GoTo pase
    Set OrdenCompra = gclsOracle.FN_Cursor("CMR.PKG_ORDEN_COMPRA.FN_DATOS", 0, pstrOCompra, pstrLocal)
    Exit Function
pase:
    Err.Raise Err.Number, "clsCompra", Err.Description
End Function

Public Function OrdenCompraDetalle(ByVal pstrOCompra As String, Optional pFlgEdit As String, Optional pCodLocal As String) As oraDynaset
On Error GoTo pase
    Set OrdenCompraDetalle = gclsOracle.FN_Cursor("CMR.PKG_ORDEN_COMPRA.FN_DETALLE", 0, pstrOCompra, pFlgEdit, pCodLocal)
    Exit Function
pase:
    Err.Raise Err.Number, "clsCompra", Err.Description
End Function

Public Function OrdenCompraLista(ByVal pstrLocal As String) As oraDynaset
On Error GoTo pase
    Set OrdenCompraLista = gclsOracle.FN_Cursor("CMR.PKG_ORDEN_COMPRA.FN_LISTA_PEN", 0, pstrLocal)
    Exit Function
pase:
    Err.Raise Err.Number, "clsCompra", Err.Description
End Function

Public Function ListaMotivo(ByVal pstrTipo As String) As oraDynaset
On Error GoTo pase
    Set ListaMotivo = gclsOracle.FN_Cursor("CMR.PKG_ORDEN_COMPRA.FN_LISTA_TIPO_OC", 0, pstrTipo)
    Exit Function
pase:
    Err.Raise Err.Number, "clsCompra", Err.Description
End Function

Public Function ListaDocumento() As oraDynaset
On Error GoTo pase
    Set ListaDocumento = gclsOracle.FN_Cursor("NUEVO.PKG_DOC_PROV.FN_LISTA_TIPODOC", 0)
    Exit Function
pase:
    Err.Raise Err.Number, "clsCompra", Err.Description
End Function

Public Sub GrabaRecepcion(ByVal pstrCia As String, _
                          ByVal pstrOCompra As String, ByVal pstrMotivo As String, _
                          ByVal pstrFormaPago As String, ByVal pstrLocal As String, _
                          ByVal pstrProveedor As String, ByVal pstrDocumento As String, _
                          ByVal pstrNumDocumento As String, ByVal pstrFecha As String, _
                          ByVal pstrModulo As String, ByVal pstrUsuario As String, _
                          ByVal pdblAfecto As Double, ByVal pdblImpuesto As Double, _
                          ByVal pdblInafecto As Double, ByVal pdblTotal As Double, _
                          ByVal pdblRedondeo As Double, ByVal pstrFlgBloqueo As String, _
                          ByVal pstrFlgContabilizar As String, ByRef pstrGuiaDevolucion As String, _
                          ByRef pstrRegCompra As String, _
                          ByRef pparProductoProv As Variant, _
                          ByRef pparCantidad As Variant, _
                          ByRef pparCantidadFrac As Variant, _
                          ByRef pparProducto As Variant, _
                          ByRef pparCantidadRec As Variant, _
                          ByRef pparCantidadFracRec As Variant, _
                          ByRef pparPrecio As Variant, _
                          ByRef pparLote As Variant, _
                          ByRef pparFechaVencimiento As Variant)

'On Error GoTo PASE

    Dim lintFilas As Integer
    
    Dim lparProductoProv As OracleInProcServer.OraParamArray
    Dim lparCantidad As OracleInProcServer.OraParamArray
    Dim lparCantidadFrac As OracleInProcServer.OraParamArray
    Dim lparProducto As OracleInProcServer.OraParamArray
    Dim lparCantidadRec As OracleInProcServer.OraParamArray
    Dim lparCantidadFracRec As OracleInProcServer.OraParamArray
    Dim lparPrecio As OracleInProcServer.OraParamArray
    Dim lparLote As OracleInProcServer.OraParamArray
    Dim lparFechaVcmto As OracleInProcServer.OraParamArray
    
    lintFilas = UBound(pparProductoProv) - LBound(pparProductoProv) + 1
    
    gclsOracle.LimpiaParametros
    
    gclsOracle.ODataBase.Parameters.Add "A_CIA", pstrCia, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_NUM_ORDEN_COMPRA", pstrOCompra, ORAPARM_BOTH, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_COD_TIPO_ORD_COM", pstrMotivo, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_COD_FORMA_PAGO", pstrFormaPago, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_COD_LOCAL", pstrLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_RUC_PROVEEDOR", pstrProveedor, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_NUM_DOCUMENTO", pstrNumDocumento, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_TIP_DOCUMENTO", pstrDocumento, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_FCH_EMISION", pstrFecha, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_COD_MODULO", pstrModulo, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_COD_USUARIO", pstrUsuario, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_MTO_IMPUESTO_MAN", pdblImpuesto, ORAPARM_INPUT, ORATYPE_NUMBER
    gclsOracle.ODataBase.Parameters.Add "A_MTO_BASE_IMP_MAN", pdblAfecto, ORAPARM_INPUT, ORATYPE_NUMBER
    gclsOracle.ODataBase.Parameters.Add "A_MTO_INAFECTO_MAN", pdblInafecto, ORAPARM_INPUT, ORATYPE_NUMBER
    gclsOracle.ODataBase.Parameters.Add "A_MTO_TOTAL_MAN", pdblTotal, ORAPARM_INPUT, ORATYPE_NUMBER
    gclsOracle.ODataBase.Parameters.Add "A_MTO_REDONDEO_DOC", pdblRedondeo, ORAPARM_INPUT, ORATYPE_NUMBER
    gclsOracle.ODataBase.Parameters.Add "A_FLG_BLOQUEO_X_PRECIO", pstrFlgBloqueo, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_FLG_CONTABILIZAR", pstrFlgContabilizar, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_NUM_GUIA_DEV", pstrGuiaDevolucion, ORAPARM_OUTPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_COD_REG", pstrRegCompra, ORAPARM_OUTPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_FLG_OFF_LINE", "", ORAPARM_OUTPUT, ORATYPE_VARCHAR2

    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_COD_PRODUCTO_PROV", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas, 20
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_CTD_PRODUCTO_DOC", ORAPARM_INPUT, ORATYPE_FLOAT, lintFilas
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_CTD_PRODUCTO_FRAC_DOC", ORAPARM_INPUT, ORATYPE_FLOAT, lintFilas
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_COD_PRODUCTO", ORAPARM_INPUT, ORATYPE_CHAR, lintFilas, 5
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_CTD_PRODUCTO_REC", ORAPARM_INPUT, ORATYPE_FLOAT, lintFilas
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_CTD_PRODUCTO_FRAC_REC", ORAPARM_INPUT, ORATYPE_FLOAT, lintFilas
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_MTO_SUBTOTAL", ORAPARM_INPUT, ORATYPE_FLOAT, lintFilas
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_NUM_LOTE", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas, 40
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_FCH_VEN", ORAPARM_INPUT, ORATYPE_CHAR, lintFilas, 10
    
    Set lparProductoProv = gclsOracle.ODataBase.Parameters("A_CAD_COD_PRODUCTO_PROV")
    Set lparCantidad = gclsOracle.ODataBase.Parameters("A_CAD_CTD_PRODUCTO_DOC")
    Set lparCantidadFrac = gclsOracle.ODataBase.Parameters("A_CAD_CTD_PRODUCTO_FRAC_DOC")
    Set lparProducto = gclsOracle.ODataBase.Parameters("A_CAD_COD_PRODUCTO")
    Set lparCantidadRec = gclsOracle.ODataBase.Parameters("A_CAD_CTD_PRODUCTO_REC")
    Set lparCantidadFracRec = gclsOracle.ODataBase.Parameters("A_CAD_CTD_PRODUCTO_FRAC_REC")
    Set lparPrecio = gclsOracle.ODataBase.Parameters("A_CAD_MTO_SUBTOTAL")
    Set lparLote = gclsOracle.ODataBase.Parameters("A_CAD_NUM_LOTE")
    Set lparFechaVcmto = gclsOracle.ODataBase.Parameters("A_CAD_FCH_VEN")
    
    Dim i As Integer
    Dim lstrSql As String
    
    lintFilas = 0
    For i = LBound(pparProductoProv) To UBound(pparProductoProv)
        lparProductoProv(lintFilas) = pparProductoProv(i)
        lparProducto(lintFilas) = pparProducto(i)
        lparPrecio(lintFilas) = pparPrecio(i)
        lparLote(lintFilas) = pparLote(i)
        lparFechaVcmto(lintFilas) = pparFechaVencimiento(i)
        lparCantidad(lintFilas) = pparCantidad(i)
        lparCantidadFrac(lintFilas) = pparCantidadFrac(i)
        lparCantidadRec(lintFilas) = pparCantidadRec(i)
        lparCantidadFracRec(lintFilas) = pparCantidadFracRec(i)
        lintFilas = lintFilas + 1
    Next
    
    'Cadena Sql
    lstrSql = "BEGIN " & _
              "NUEVO.PKG_DOC_PROV.SP_GRABA(A_CIA                       => :A_CIA," & _
              "                            A_NUM_ORDEN_COMPRA          => :A_NUM_ORDEN_COMPRA," & _
              "                            A_COD_TIPO_ORD_COM          => :A_COD_TIPO_ORD_COM," & _
              "                            A_COD_FORMA_PAGO            => :A_COD_FORMA_PAGO," & _
              "                            A_COD_LOCAL                 => :A_COD_LOCAL," & _
              "                            A_RUC_PROVEEDOR             => :A_RUC_PROVEEDOR," & _
              "                            A_NUM_DOCUMENTO             => :A_NUM_DOCUMENTO," & _
              "                            A_TIP_DOCUMENTO             => :A_TIP_DOCUMENTO," & _
              "                            A_FCH_EMISION               => :A_FCH_EMISION," & _
              "                            A_COD_MODULO                => :A_COD_MODULO," & _
              "                            A_COD_USUARIO               => :A_COD_USUARIO," & _
              "                            A_MTO_IMPUESTO_MAN          => :A_MTO_IMPUESTO_MAN," & _
              "                            A_MTO_BASE_IMP_MAN          => :A_MTO_BASE_IMP_MAN," & _
              "                            A_MTO_INAFECTO_MAN          => :A_MTO_INAFECTO_MAN,"
    
    lstrSql = lstrSql & _
              "                            A_MTO_TOTAL_MAN             => :A_MTO_TOTAL_MAN," & _
              "                            A_MTO_REDONDEO_DOC          => :A_MTO_REDONDEO_DOC," & _
              "                            A_FLG_BLOQUEO_X_PRECIO      => :A_FLG_BLOQUEO_X_PRECIO," & _
              "                            A_CAD_COD_PRODUCTO_PROV     => :A_CAD_COD_PRODUCTO_PROV," & _
              "                            A_CAD_CTD_PRODUCTO_DOC      => :A_CAD_CTD_PRODUCTO_DOC," & _
              "                            A_CAD_CTD_PRODUCTO_FRAC_DOC => :A_CAD_CTD_PRODUCTO_FRAC_DOC," & _
              "                            A_CAD_COD_PRODUCTO          => :A_CAD_COD_PRODUCTO," & _
              "                            A_CAD_CTD_PRODUCTO_REC      => :A_CAD_CTD_PRODUCTO_REC," & _
              "                            A_CAD_CTD_PRODUCTO_FRAC_REC => :A_CAD_CTD_PRODUCTO_FRAC_REC," & _
              "                            A_CAD_MTO_SUBTOTAL          => :A_CAD_MTO_SUBTOTAL," & _
              "                            A_CAD_NUM_LOTE              => :A_CAD_NUM_LOTE," & _
              "                            A_CAD_FCH_VEN               => :A_CAD_FCH_VEN," & _
              "                            A_NUM_GUIA_DEV              => :A_NUM_GUIA_DEV," & _
              "                            A_COD_REG                   => :A_COD_REG," & _
              "                            A_FLG_CONTABILIZAR          => :A_FLG_CONTABILIZAR, A_FLG_OFF_LINE => :A_FLG_OFF_LINE);" & _
              "END;"

    'Invocar el package de Razuri
    gclsOracle.ODataBase.ExecuteSQL (lstrSql)
        
    pstrGuiaDevolucion = Trim("" & gclsOracle.ODataBase.Parameters("A_NUM_GUIA_DEV").Value)
    pstrRegCompra = Trim("" & gclsOracle.ODataBase.Parameters("A_COD_REG").Value)
    
'    Exit Sub
'PASE:
'    Err.Raise Err.Number, "clsCompra", Err.Description
End Sub

Public Function ParteRecepcion(ByVal pstrProveedor As String, _
                               ByVal pstrNumDocumento As String, _
                               ByVal pstrTipDocumento As String) As oraDynaset
On Error GoTo pase

    Set ParteRecepcion = gclsOracle.FN_Cursor("NUEVO.PKG_RECEPCION.FN_PARTE_CAB", 0, pstrProveedor, pstrNumDocumento, pstrTipDocumento)
    
Exit Function
pase:
    Err.Raise Err.Number, "clsCompra", Err.Description
End Function

Public Function ParteRecepcionDetalle(ByVal pstrProveedor As String, _
                                      ByVal pstrNumDocumento As String, _
                                      ByVal pstrTipDocumento As String) As oraDynaset
On Error GoTo pase

    Set ParteRecepcionDetalle = gclsOracle.FN_Cursor("NUEVO.PKG_RECEPCION.FN_PARTE_DET", 0, pstrProveedor, pstrNumDocumento, pstrTipDocumento)
    
Exit Function
pase:
    Err.Raise Err.Number, "clsCompra", Err.Description
End Function

Public Function ParteRecepcionDevolucion(ByVal pstrProveedor As String, _
                                         ByVal pstrNumDocumento As String, _
                                         ByVal pstrTipDocumento As String) As oraDynaset
On Error GoTo pase

    Set ParteRecepcionDevolucion = gclsOracle.FN_Cursor("NUEVO.PKG_RECEPCION.FN_DEVOLUCION", 0, pstrProveedor, pstrNumDocumento, pstrTipDocumento)

Exit Function
pase:
    Err.Raise Err.Number, "clsCompra", Err.Description
End Function

Public Function ParteRecepcionItems(ByVal pstrProveedor As String, _
                                    ByVal pstrNumDocumento As String, _
                                    ByVal pstrTipDocumento As String) As oraDynaset
On Error GoTo pase

    Set ParteRecepcionItems = gclsOracle.FN_Cursor("NUEVO.PKG_RECEPCION.FN_DETALLE", 0, pstrProveedor, pstrNumDocumento, pstrTipDocumento)
    
Exit Function
pase:
    Err.Raise Err.Number, "clsCompra", Err.Description
End Function

Public Function Anula(ByVal pstrCia As String, _
                      ByVal pstrLocal As String, _
                      ByVal pstrSecDocumento As String, _
                      ByVal pstrUsuario As String, _
                      ByVal pstrFlgUsuInv As String, _
                      Optional pstrFlgCommit As String, _
                      Optional pstrFlgOffLine As String) As String
On Error GoTo pase

    Dim varValores As Variant
    Dim varIO As Variant

    varValores = Array(pstrCia, pstrLocal, pstrSecDocumento, pstrUsuario, pstrFlgUsuInv, pstrFlgCommit, pstrFlgOffLine)
    varIO = Array(entrada, entrada, entrada, entrada, entrada, entrada, entrada)
    
    Anula = gclsOracle.SP("NUEVO.PKG_DOC_PROV.SP_ANULA", varValores, varIO)

Exit Function
pase:
    Err.Raise Err.Number, "clsCompra", Err.Description
End Function

