VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSPVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private strCodLocal As String
Private strNumero As String
Private strCodUsuario As String
Private arrDetalle As New XArrayDB

Public Property Get Numero() As String
    Numero = strNumero
End Property

Public Property Let Numero(ByVal vNewValue As String)
    strNumero = vNewValue
End Property

Public Property Get CodLocal() As String
    CodLocal = strCodLocal
End Property

Public Property Let CodLocal(ByVal vNewValue As String)
    strCodLocal = vNewValue
End Property

Public Property Get CodUsuario() As String
    CodUsuario = strCodUsuario
End Property

Public Property Let CodUsuario(ByVal vNewValue As String)
    strCodUsuario = vNewValue
End Property

'Public Property Get Estado() As clsEstadoSPVM
'    Set Estado = objEstado
'End Property
'
'Public Property Set Estado(ByVal vNewValue As clsEstadoSPVM)
'    Set objEstado = vNewValue
'End Property

Public Property Get Detalle() As XArrayDB
    Set Detalle = arrDetalle
End Property

Public Property Set Detalle(ByVal vNewValue As XArrayDB)
    Set arrDetalle = vNewValue
End Property

Public Function Lista(ByVal vstrDesde As String, _
               ByVal vstrHasta As String, _
               ByVal vstrCodLocal As String, _
               ByVal vstrCodEstado As String, _
               ByVal vstrNumero As String, _
               ByVal vstrCodProd As String) As oraDynaset
    
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_SPVM.FN_LISTA_CAB", 0, _
                    vstrDesde, _
                    vstrHasta, _
                    vstrCodLocal, _
                    vstrCodEstado, _
                    vstrNumero, _
                    vstrCodProd)
End Function


Public Function ListaDet(ByVal vstrCodLocal As String, _
                  ByVal vstrNumero As String) As oraDynaset
    
    Set ListaDet = gclsOracle.FN_Cursor("BTLPROD.PKG_SPVM.FN_LISTA_DET", 0, _
                                        vstrCodLocal, _
                                        vstrNumero)
End Function

Public Function Filtro(ByVal vstrCodLocal As String, _
                       ByVal vstrCodProd As String, _
                       ByVal vstrCodLaboratorio As String, _
                       ByVal vstrCodLinea As String) As oraDynaset
    
    Set Filtro = gclsOracle.FN_Cursor("BTLPROD.PKG_SPVM.FN_FILTRO", 0, _
                       vstrCodLocal, _
                       vstrCodProd, _
                       vstrCodLaboratorio, _
                       vstrCodLinea)
End Function

'''Public Sub Grabar()
'''    Dim i As Integer
'''    Dim intDetalle As Integer
'''    Dim strSQL As String
'''
'''    strNumero = ""
'''
'''    intDetalle = Detalle.Count(1)
''''    If objOrdPic.Bandeja.Codigo = "" Then
''''        Err.Raise 1, "clsTransferencia.Grabar", "Falta Dato de Bandeja"
''''    End If
''''    If objOrdPic.Numero = "" Then
''''        Err.Raise 1, "clsTransferencia.Grabar", "Falta Dato de Numero de OP"
''''    End If
''''
'''    If intDetalle = 0 Then
'''        Err.Raise 1, "Detalle de Solicitud", "No existen items para grabar Solicitud"
'''    End If
'''
'''    gclsOracle.LimpiaParametros
'''
'''    With gclsOracle.ODataBase.Parameters
'''        .AddTable "V_COD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 10
'''        '.AddTable "V_COD_MOTIVO_SPVM", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 5
'''        '.AddTable "V_CTD_DIAS_INVENTARIO", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 10
'''        .AddTable "v_ctd_pvm_solicitado", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 10
'''        .AddTable "v_ctd_pvm_registro", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 10
'''        '.AddTable "V_CTD_DIAS_QUIEBRE", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 10
'''        '.AddTable "V_DES_OBSERVACION", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 100
'''        '.AddTable "V_CTD_DIAS_AGOTADO", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 10
'''        .AddTable "V_COD_EST_ABAST", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 5
'''        .AddTable "V_FLG_SELECCIONADO", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 1
'''    End With
'''
'''    Dim arrTabla As New Collection
'''    Dim parArray As OraParamArray
'''
'''    For Each parArray In gclsOracle.ODataBase.Parameters
'''        arrTabla.Add parArray, parArray.name
'''    Next
'''
'''    For i = Detalle.LowerBound(1) To Detalle.UpperBound(1)
'''        arrTabla("V_COD_PRODUCTO")(i) = Detalle(i, 0)
'''        'arrTabla("V_COD_MOTIVO_SPVM")(i) = Detalle(i, 11)
'''        'arrTabla("V_CTD_DIAS_INVENTARIO")(i) = Detalle(i, 13)
'''        arrTabla("v_ctd_pvm_solicitado")(i) = Detalle(i, 9)
'''        arrTabla("v_ctd_pvm_registro")(i) = Detalle(i, 8)
'''        'arrTabla("V_CTD_DIAS_QUIEBRE")(i) = Detalle(i, 8)
'''        'arrTabla("V_DES_OBSERVACION")(i) = Detalle(i, 12)
'''        'arrTabla("V_CTD_DIAS_AGOTADO")(i) = Detalle(i, 6)
'''        arrTabla("V_COD_EST_ABAST")(i) = Detalle(i, 5)
'''        arrTabla("V_FLG_SELECCIONADO")(i) = Detalle(i, 4)
'''    Next i
'''
'''    gclsOracle.ODataBase.Parameters.Add "A_COD_LOCAL", CodLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2
'''    gclsOracle.ODataBase.Parameters.Add "A_COD_USUARIO", CodUsuario, ORAPARM_INPUT, ORATYPE_VARCHAR2
'''    gclsOracle.ODataBase.Parameters.Add "A_NUM_SPVM", Space(13), ORAPARM_OUTPUT, ORATYPE_VARCHAR2
'''
'''
'''    strSQL = "BEGIN" & _
'''            "       BTLPROD.PKG_SPVM.SP_REGISTRA (A_COD_LOCAL             => :A_COD_LOCAL," & _
'''            "                                    A_COD_USUARIO           => :A_COD_USUARIO," & _
'''            "                                    V_COD_PRODUCTO          => :V_COD_PRODUCTO," & _
'''            "                                    v_ctd_pvm_solicitado => :v_ctd_pvm_solicitado," & _
'''            "                                    v_ctd_pvm_registro   => :v_ctd_pvm_registro," & _
'''            "                                    V_COD_EST_ABAST         => :V_COD_EST_ABAST," & _
'''            "                                    V_FLG_SELECCIONADO      => :V_FLG_SELECCIONADO," & _
'''            "                                    A_NUM_SPVM               => :A_NUM_SPVM);" & _
'''            "END;"
'''
'''
'''
'''    gclsOracle.Execute strSQL
'''
'''    Numero = "" & gclsOracle.ODataBase.Parameters("A_NUM_SPVM")
'''
'''
'''End Sub

Public Sub Anular(ByVal vstrCodLocal As String, _
           ByVal vstrNumero As String, _
           ByVal vstrCodUsu As String)

    Dim varrValores As Variant
    Dim varrIO As Variant
    Dim strError As String

    varrValores = Array(vstrCodLocal, vstrNumero, vstrCodUsu)
    varrIO = Array(entrada, entrada, entrada)
    
    strError = gclsOracle.SP("BTLPROD.PKG_SPVM.SP_ANULA", varrValores, varrIO)
    
    If strError <> "" Then
        Err.Raise -1, "", strError
    End If

End Sub


Public Function ListaProductosPVM(ByVal vstrCodLocal As String, _
                                  ByVal codLaboratorio As String, _
                                  ByVal codLinea As String, _
                                  ByVal codProducto As String) As oraDynaset
    
    Set ListaProductosPVM = gclsOracle.FN_Cursor("BTLPROD.PKG_SPVM.fn_lista_productos", 0, _
                                        vstrCodLocal, codLaboratorio, codProducto, codLinea)
End Function

Public Sub AsignarCantidadPVM(ByVal vstrCodLocal As String, _
           ByVal codProducto As String, _
           ByVal usuario As String, _
           ByVal cantidad As Integer)

    Dim varrValores As Variant
    Dim varrIO As Variant
    Dim strError As String

    varrValores = Array(vstrCodLocal, codProducto, usuario, cantidad)
    varrIO = Array(entrada, entrada, entrada, entrada)
    
    strError = gclsOracle.SP("BTLPROD.PKG_SPVM.sp_inserta_det", varrValores, varrIO)
    
    If strError <> "" Then
        Err.Raise -1, "", strError
    End If
    
End Sub

Public Function HistorialPVM(ByVal vstrCodLocal As String, ByVal vstrCodProducto As String) As oraDynaset
    
    Set HistorialPVM = gclsOracle.FN_Cursor("BTLPROD.PKG_SPVM.fn_det_pvm", 0, _
                                        vstrCodLocal, vstrCodProducto)
End Function

Public Function LabelMeses() As oraDynaset
    Set LabelMeses = gclsOracle.FN_Cursor("btlprod.fn_dev_eti_meses", 0)
End Function



