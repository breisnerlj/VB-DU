VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSMM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private strCodLocal As String
Private strNumero As String
Private strCodUsuario As String
'Private objEstado As New clsEstadoSMM
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

'Public Property Get Estado() As clsEstadoSMM
'    Set Estado = objEstado
'End Property
'
'Public Property Set Estado(ByVal vNewValue As clsEstadoSMM)
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
    
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_SMM.FN_LISTA", 0, _
                    vstrDesde, _
                    vstrHasta, _
                    vstrCodLocal, _
                    vstrCodEstado, _
                    vstrNumero, _
                    vstrCodProd)
End Function


Public Function ListaDet(ByVal vstrCodLocal As String, _
                  ByVal vstrNumero As String) As oraDynaset
    
    Set ListaDet = gclsOracle.FN_Cursor("BTLPROD.PKG_SMM.FN_LISTA", 0, _
                                        vstrCodLocal, _
                                        vstrNumero)
End Function

Public Function Filtro(ByVal vstrCodLocal As String, _
                       ByVal vstrCodProd As String, _
                       ByVal vstrCodLaboratorio As String, _
                       ByVal vstrCodLinea As String) As oraDynaset
    
    Set Filtro = gclsOracle.FN_Cursor("BTLPROD.PKG_SMM.FN_FILTRO", 0, _
                       vstrCodLocal, _
                       vstrCodProd, _
                       vstrCodLaboratorio, _
                       vstrCodLinea)
End Function

Public Sub Grabar()
Dim i As Integer
Dim intDetalle As Integer
Dim StrSql As String
    
    Numero = ""
    
    intDetalle = Detalle.Count(1)
'    If objOrdPic.Bandeja.Codigo = "" Then
'        Err.Raise 1, "clsTransferencia.Grabar", "Falta Dato de Bandeja"
'    End If
'    If objOrdPic.Numero = "" Then
'        Err.Raise 1, "clsTransferencia.Grabar", "Falta Dato de Numero de OP"
'    End If
'
    If intDetalle = 0 Then
        Err.Raise 1, "Detalle de Solicitud", "No existen items para grabar Solicitud"
    End If
    
    gclsOracle.LimpiaParametros
    
    With gclsOracle.ODataBase.Parameters
        .AddTable "V_COD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 10
        .AddTable "V_COD_MOTIVO_SMM", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 5
        .AddTable "V_CTD_DIAS_INVENTARIO", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 10
        .AddTable "V_CTD_MAXIMO_SOLICITADO", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 10
        .AddTable "V_CTD_MAXIMO_REGISTRO", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 10
        .AddTable "V_CTD_DIAS_QUIEBRE", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 10
        .AddTable "V_DES_OBSERVACION", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 100
        .AddTable "V_CTD_DIAS_AGOTADO", ORAPARM_INPUT, ORATYPE_VARCHAR2, intDetalle, 10
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
        arrTabla("V_COD_MOTIVO_SMM")(i) = Detalle(i, 11)
        arrTabla("V_CTD_DIAS_INVENTARIO")(i) = Detalle(i, 13)
        arrTabla("V_CTD_MAXIMO_SOLICITADO")(i) = Detalle(i, 10)
        arrTabla("V_CTD_MAXIMO_REGISTRO")(i) = Detalle(i, 9)
        arrTabla("V_CTD_DIAS_QUIEBRE")(i) = Detalle(i, 8)
        arrTabla("V_DES_OBSERVACION")(i) = Detalle(i, 12)
        arrTabla("V_CTD_DIAS_AGOTADO")(i) = Detalle(i, 6)
        arrTabla("V_COD_EST_ABAST")(i) = Detalle(i, 5)
        arrTabla("V_FLG_SELECCIONADO")(i) = Detalle(i, 4)
    Next i
    
    gclsOracle.ODataBase.Parameters.Add "A_COD_LOCAL", CodLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_COD_USUARIO", CodUsuario, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_NUM_SMM", Space(13), ORAPARM_OUTPUT, ORATYPE_VARCHAR2
    
    
    StrSql = "BEGIN" & _
            "       BTLPROD.PKG_SMM.SP_REGISTRA (A_COD_LOCAL             => :A_COD_LOCAL," & _
            "                                    A_COD_USUARIO           => :A_COD_USUARIO," & _
            "                                    V_COD_PRODUCTO          => :V_COD_PRODUCTO," & _
            "                                    V_COD_MOTIVO_SMM        => :V_COD_MOTIVO_SMM," & _
            "                                    V_CTD_DIAS_INVENTARIO   => :V_CTD_DIAS_INVENTARIO," & _
            "                                    V_CTD_MAXIMO_SOLICITADO => :V_CTD_MAXIMO_SOLICITADO," & _
            "                                    V_CTD_MAXIMO_REGISTRO   => :V_CTD_MAXIMO_REGISTRO," & _
            "                                    V_CTD_DIAS_QUIEBRE      => :V_CTD_DIAS_QUIEBRE," & _
            "                                    V_DES_OBSERVACION       => :V_DES_OBSERVACION," & _
            "                                    V_CTD_DIAS_AGOTADO      => :V_CTD_DIAS_AGOTADO," & _
            "                                    V_COD_EST_ABAST         => :V_COD_EST_ABAST," & _
            "                                    V_FLG_SELECCIONADO      => :V_FLG_SELECCIONADO," & _
            "                                    A_NUM_SMM               => :A_NUM_SMM);" & _
            "END;"


    
    gclsOracle.Execute StrSql
    
    Numero = "" & gclsOracle.ODataBase.Parameters("A_NUM_SMM")

        
End Sub



Public Sub Anular(ByVal vstrCodLocal As String, _
           ByVal vstrNumero As String, _
           ByVal vstrCodUsu As String)

    Dim varrValores As Variant
    Dim varrIO As Variant
    Dim strError As String

    varrValores = Array(vstrCodLocal, vstrNumero, vstrCodUsu)
    varrIO = Array(entrada, entrada, entrada)
    
    strError = gclsOracle.SP("BTLPROD.PKG_SMM.SP_ANULA", varrValores, varrIO)
    
    If strError <> "" Then
        Err.Raise -1, "", strError
    End If

End Sub

