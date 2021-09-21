VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPedEspecial"
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
    
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_PEDESPECIAL.FN_LISTA_CAB", 0, _
                    vstrDesde, _
                    vstrHasta, _
                    vstrCodLocal, _
                    vstrCodEstado, _
                    vstrNumero, _
                    vstrCodProd)
End Function


Public Function ListaDet(ByVal vstrCodLocal As String, _
                  ByVal vstrNumero As String) As oraDynaset
    
    Set ListaDet = gclsOracle.FN_Cursor("BTLPROD.PKG_PEDESPECIAL.FN_LISTA_DET", 0, _
                                        vstrCodLocal, _
                                        vstrNumero)
End Function

Public Function Filtro(ByVal vstrCodLocal As String, _
                       ByVal vstrCodProd As String, _
                       ByVal vstrCodLaboratorio As String, _
                       ByVal vstrCodLinea As String) As oraDynaset
    
    Set Filtro = gclsOracle.FN_Cursor("BTLPROD.PKG_PEDESPECIAL.FN_FILTRO", 0, _
                       vstrCodLocal, _
                       vstrCodProd, _
                       vstrCodLaboratorio, _
                       vstrCodLinea)
End Function

Public Sub Grabar()
Dim i As Integer
Dim intDetalle As Integer
Dim strSQL As String
    
    Numero = ""
    
    intDetalle = Detalle.Count(1)

    If intDetalle = 0 Then
        Err.Raise 1, "Detalle de Solicitud", "No existen items para grabar Pedido"
    End If
    
    gclsOracle.LimpiaParametros
    
    With gclsOracle.ODataBase.Parameters
        .AddTable "V_COD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 10
        .AddTable "v_ctd_solicitado", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 10
        .AddTable "V_COD_EST_ABAST", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 5
        .AddTable "V_FLG_SELECCIONADO", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 1
    End With
    
    Dim arrTabla As New Collection
    Dim parArray As OraParamArray
           
    For Each parArray In gclsOracle.ODataBase.Parameters
        arrTabla.Add parArray, parArray.name
    Next
    
    For i = Detalle.LowerBound(1) To Detalle.UpperBound(1)
        arrTabla("V_COD_PRODUCTO")(i) = Detalle(i, 0)
        arrTabla("v_ctd_solicitado")(i) = Detalle(i, 9)
        arrTabla("V_COD_EST_ABAST")(i) = Detalle(i, 5)
        arrTabla("V_FLG_SELECCIONADO")(i) = Detalle(i, 4)
    Next i
    
    gclsOracle.ODataBase.Parameters.Add "A_COD_LOCAL", CodLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_COD_USUARIO", CodUsuario, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_NUM_PEDESPECIAL", Space(13), ORAPARM_OUTPUT, ORATYPE_VARCHAR2
    
    strSQL = "BEGIN" & _
            "       BTLPROD.PKG_PEDESPECIAL.SP_REGISTRA (A_COD_LOCAL             => :A_COD_LOCAL," & _
            "                                            A_COD_USUARIO           => :A_COD_USUARIO," & _
            "                                            V_COD_PRODUCTO          => :V_COD_PRODUCTO," & _
            "                                            v_ctd_solicitado        => :v_ctd_solicitado," & _
            "                                            V_COD_EST_ABAST         => :V_COD_EST_ABAST," & _
            "                                            V_FLG_SELECCIONADO      => :V_FLG_SELECCIONADO," & _
            "                                            A_NUM_PEDESPECIAL       => :A_NUM_PEDESPECIAL);" & _
            "END;"


    
    gclsOracle.Execute strSQL
    
    Numero = "" & gclsOracle.ODataBase.Parameters("A_NUM_PEDESPECIAL")
    
End Sub

Public Sub Actualizar()
Dim i As Integer
Dim intDetalle As Integer
Dim strSQL As String
      
    intDetalle = Detalle.Count(1)

    If intDetalle = 0 Then
        Err.Raise 1, "Detalle de Solicitud", "No existen items para grabar Pedido"
    End If
    
    gclsOracle.LimpiaParametros
    
    With gclsOracle.ODataBase.Parameters
        .AddTable "V_COD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 10
        .AddTable "v_ctd_solicitado", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 10
        .AddTable "V_COD_EST_ABAST", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 5
        .AddTable "V_FLG_SELECCIONADO", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 1
    End With
    
    Dim arrTabla As New Collection
    Dim parArray As OraParamArray
           
    For Each parArray In gclsOracle.ODataBase.Parameters
        arrTabla.Add parArray, parArray.name
    Next
    
    For i = Detalle.LowerBound(1) To Detalle.UpperBound(1)
        arrTabla("V_COD_PRODUCTO")(i) = Detalle(i, 0)
        arrTabla("v_ctd_solicitado")(i) = Detalle(i, 9)
        arrTabla("V_COD_EST_ABAST")(i) = Detalle(i, 5)
        arrTabla("V_FLG_SELECCIONADO")(i) = Detalle(i, 4)
    Next i
    
    gclsOracle.ODataBase.Parameters.Add "A_COD_LOCAL", CodLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "C_NUM_PEDESPECIAL", Numero, ORAPARM_INPUT, ORATYPE_VARCHAR2
    
    strSQL = "BEGIN" & _
            "       BTLPROD.PKG_PEDESPECIAL.SP_EDITA    (A_COD_LOCAL             => :A_COD_LOCAL," & _
            "                                            V_COD_PRODUCTO          => :V_COD_PRODUCTO," & _
            "                                            v_ctd_solicitado        => :v_ctd_solicitado," & _
            "                                            V_COD_EST_ABAST         => :V_COD_EST_ABAST," & _
            "                                            V_FLG_SELECCIONADO      => :V_FLG_SELECCIONADO," & _
            "                                            C_NUM_PEDESPECIAL       => :C_NUM_PEDESPECIAL);" & _
            "END;"
            
    gclsOracle.Execute strSQL
    
End Sub

Public Sub Anular(ByVal vstrCodLocal As String, _
           ByVal vstrNumero As String, _
           ByVal vstrCodUsu As String)

    Dim varrValores As Variant
    Dim varrIO As Variant
    Dim strError As String

    varrValores = Array(vstrCodLocal, vstrNumero, vstrCodUsu)
    varrIO = Array(entrada, entrada, entrada)
    
    strError = gclsOracle.SP("BTLPROD.PKG_PEDESPECIAL.SP_ANULA", varrValores, varrIO)
    
    If strError <> "" Then
        Err.Raise -1, "", strError
    End If

End Sub





