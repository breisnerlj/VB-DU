VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ListaTipo() As oraDynaset
    On Error GoTo CtrlErr
        Set ListaTipo = gclsOracle.FN_Cursor("CMR.PKG_AUTORIZACION.FN_LISTA", 0)
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAutorizacion.Lista", Err.Description
End Function

Public Function Grabar(ByVal Cia As String, _
                       ByVal CodigoLocal As String, _
                       ByVal TipoAjuste As String, _
                       ByVal CodigoUsuario As String, _
                       ByVal ObservacionAjuste As String _
                       ) As String
                       
    On Error GoTo Pase
    
    Dim arrValores As Variant
    Dim arrDireccion As Variant
    Dim A_CAD_COD_PRODUCTO As String
    Dim A_CAD_CTD_PRODUCTO As String
    Dim A_CAD_CTD_PRODUCTO_FRAC As String
    Dim A_CAD_IMP_PRODUCTO As String
    Dim p As Integer
    While p < objVenta.Producto.Count(1)
        A_CAD_COD_PRODUCTO = A_CAD_COD_PRODUCTO & objVenta.Producto(p, 0) & "|"
        If objVenta.Producto(p, 0) = "0" Then
            A_CAD_CTD_PRODUCTO = A_CAD_CTD_PRODUCTO & objVenta.Producto(p, 0) & "|"
            A_CAD_CTD_PRODUCTO_FRAC = A_CAD_CTD_PRODUCTO_FRAC & "0" & "|"
        Else
            A_CAD_CTD_PRODUCTO = A_CAD_CTD_PRODUCTO & objVenta.Producto(p, 0) & "|"
            A_CAD_CTD_PRODUCTO_FRAC = A_CAD_CTD_PRODUCTO_FRAC & "0" & "|"
        End If
        A_CAD_IMP_PRODUCTO = A_CAD_IMP_PRODUCTO & objVenta.Producto(p, 4) & "|"
        p = p + 1
    Wend

    arrValores = Array("", Cia, CodigoLocal, TipoAjuste, CodigoUsuario, ObservacionAjuste, A_CAD_COD_PRODUCTO, A_CAD_CTD_PRODUCTO, A_CAD_CTD_PRODUCTO_FRAC, A_CAD_IMP_PRODUCTO)
    arrDireccion = Array(Salida, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada)
    Grabar = gclsOracle.SP("BTLPROD.PKG_AJUSTE.SP_GRABA", arrValores, arrDireccion)
    
    If Grabar = "" Then MsgBox "Se Grabo Satisfactorimente", vbExclamation, App.ProductName: Exit Function
    MsgBox Grabar
    Exit Function
Pase:
    Err.Raise Err.Number, "clsCobranza", Err.Description
End Function

Public Function ListaTipoAju() As oraDynaset
    On Error GoTo CtrlErr
      Set ListaTipoAju = gclsOracle.FN_Cursor("BTLPROD.PKG_AJUSTE.FN_LISTA_TIPO_AJU", 0)
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAjuste.ListaTipoAju", Err.Description
End Function

Public Function ListaMotivoTipoAju(ByVal vstrCodTipoAju As String, _
                                   Optional vstrCodLocal As String = vbNullString) As oraDynaset
    On Error GoTo CtrlErr
      Set ListaMotivoTipoAju = gclsOracle.FN_Cursor("BTLPROD.PKG_AJUSTE.FN_LISTA_MOTIVO_AJU", 0, vstrCodTipoAju, vstrCodLocal)
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAjuste.ListaMotivoTipoAju", Err.Description
End Function

Public Function ListaCabAjuste(ByVal NumAjuste As String, Optional ByVal CodAjuste As String = "") As oraDynaset
    On Error GoTo CtrlErr
        Set ListaCabAjuste = gclsOracle.FN_Cursor("BTLPROD.PKG_AJUSTE.FN_LISTA_CAB_AJUSTE", 0, NumAjuste, CodAjuste)
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAjuste.ListaCabAjuste", Err.Description
End Function

Public Function ListaDetAjuste(ByVal NumAjuste As String, ByVal CodZona As String, ByVal Signo As String) As oraDynaset
    On Error GoTo CtrlErr
        Set ListaDetAjuste = gclsOracle.FN_Cursor("BTLPROD.PKG_AJUSTE.FN_LISTA_DET_AJUSTE", 0, NumAjuste, CodZona, Signo)
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAjuste.ListaDetAjuste", Err.Description
End Function

Public Function ListaTipoAjuste(ByVal NumAjuste As String, ByVal CodZona As String) As oraDynaset
    On Error GoTo CtrlErr
        Set ListaTipoAjuste = gclsOracle.FN_Cursor("BTLPROD.PKG_AJUSTE.FN_LISTA_TIPO", 0, NumAjuste, CodZona)
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAjuste.ListaTipoAjuste", Err.Description
End Function

Public Function ListaMotivo(ByVal NumAjuste As String, ByVal CodZona As String) As oraDynaset
    On Error GoTo CtrlErr
        Set ListaMotivo = gclsOracle.FN_Cursor("BTLPROD.PKG_AJUSTE.FN_LISTA_MOTIVO", 0, NumAjuste, CodZona)
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAjuste.ListaMotivo", Err.Description
End Function

Public Function DevValor(ByVal vstrCodLocal As String, ByVal vstrCodProducto As String, _
                         ByVal vstrCtdUnd As String, ByVal vstrCtdFrac As String) As Double
    On Error GoTo CtrlErr
        DevValor = gclsOracle.FN_Valor("BTLPROD.PKG_SOLICITUD_AJUSTE.FN_DEV_VALOR", vstrCodLocal, vstrCodProducto, vstrCtdUnd, vstrCtdFrac)
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAjuste.DevValor", Err.Description
End Function

Public Function GrabaSolicitud(ByRef RetNumDocumento As String, ByVal vstrCia As String, ByVal vstrCodLocal As String, ByVal vstrCodUsuario As String, _
                               ByVal vstrObservaciones As String, ByVal varrTipo As Variant, ByVal varrMotivo As Variant, _
                               ByVal vstrProducto As Variant, ByVal varrCtdUnd As Variant, ByVal varrCtdFrac As Variant) As String
                          

On Error GoTo Control
    Dim lstrsql As String
    Dim i As Integer
    Dim lintFilas As Integer

    Dim lparTipo As OracleInProcServer.OraParamArray
    Dim lparMotivo As OracleInProcServer.OraParamArray
    Dim lparProducto As OracleInProcServer.OraParamArray
    Dim lparCtdUnd As OracleInProcServer.OraParamArray
    Dim lparCtdFrac As OracleInProcServer.OraParamArray

    lintFilas = UBound(varrTipo) - LBound(varrTipo) + 1
    gclsOracle.LimpiaParametros

    With gclsOracle.ODataBase.Parameters
        .AddTable "A_CAD_COD_TIPO", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas, 8
        .AddTable "A_CAD_COD_MOTIVO", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas, 8
        .AddTable "A_CAD_COD_PRODUCTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas, 8
        .AddTable "A_CAD_CTD_UND_FIS", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas, 8
        .AddTable "A_CAD_CTD_FRAC_FIS", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas, 8
    End With

    Set lparTipo = gclsOracle.ODataBase.Parameters("A_CAD_COD_TIPO")
    Set lparMotivo = gclsOracle.ODataBase.Parameters("A_CAD_COD_MOTIVO")
    Set lparProducto = gclsOracle.ODataBase.Parameters("A_CAD_COD_PRODUCTO")
    Set lparCtdUnd = gclsOracle.ODataBase.Parameters("A_CAD_CTD_UND_FIS")
    Set lparCtdFrac = gclsOracle.ODataBase.Parameters("A_CAD_CTD_FRAC_FIS")

    For i = LBound(varrTipo) To UBound(varrTipo)
            lparTipo(i) = varrTipo(i)
            lparMotivo(i) = varrMotivo(i)
            lparProducto(i) = vstrProducto(i)
            lparCtdUnd(i) = varrCtdUnd(i)
            lparCtdFrac(i) = varrCtdFrac(i)
    Next i

    With gclsOracle.ODataBase.Parameters
         .Add "RET_NUM_SOLICITUD", RetNumDocumento, ORAPARM_OUTPUT, ORATYPE_VARCHAR2
         .Add "A_CIA", vstrCia, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_COD_LOCAL", vstrCodLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_COD_USUARIO", vstrCodUsuario, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_OBSERVACION", vstrObservaciones, ORAPARM_INPUT, ORATYPE_VARCHAR2
    End With

    lstrsql = "BEGIN " & _
              "BTLPROD.PKG_SOLICITUD_AJUSTE.SP_GRABA_SOLICITUD(RET_NUM_SOLICITUD    => :RET_NUM_SOLICITUD," & _
              "                                                A_CIA                => :A_CIA," & _
              "                                                A_COD_LOCAL          => :A_COD_LOCAL," & _
              "                                                A_COD_USUARIO        => :A_COD_USUARIO," & _
              "                                                A_OBSERVACION        => :A_OBSERVACION," & _
              "                                                A_CAD_COD_TIPO       => :A_CAD_COD_TIPO," & _
              "                                                A_CAD_COD_MOTIVO     => :A_CAD_COD_MOTIVO," & _
              "                                                A_CAD_COD_PRODUCTO   => :A_CAD_COD_PRODUCTO," & _
              "                                                A_CAD_CTD_UND_FIS    => :A_CAD_CTD_UND_FIS," & _
              "                                                A_CAD_CTD_FRAC_FIS   => :A_CAD_CTD_FRAC_FIS);" & _
              "END;"

    gclsOracle.ODataBase.ExecuteSQL (lstrsql)

    RetNumDocumento = Trim("" & gclsOracle.ODataBase.Parameters("RET_NUM_SOLICITUD").Value)

   Exit Function

Control:
    Err.Raise Err.Number, "clsAjuste.GrabaSolicitud", Err.Description
End Function

Public Function DevNroSolicitud() As String
    On Error GoTo CtrlErr
        DevNroSolicitud = gclsOracle.FN_Valor("BTLPROD.PKG_SOLICITUD_AJUSTE.FN_DEV_NROSOLICITUD")
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAjuste.DevNroSolicitud", Err.Description
End Function

Public Function Lista(ByVal CodigoLocal As String, _
                      ByVal FechaInicio As String, _
                      ByVal FechaFin As String, _
                      ByVal Estado As String) As oraDynaset
On Error GoTo CtrlErr
      Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_SOLICITUD_AJUSTE.LISTA_SOL_LOCAL", 0, CodigoLocal, FechaInicio, FechaFin, Estado)
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAjuste.Lista", Err.Description
End Function

Public Function ListaDetalle(ByVal NumeroSolicitud As String) As oraDynaset
On Error GoTo CtrlErr
      Set ListaDetalle = gclsOracle.FN_Cursor("BTLPROD.PKG_SOLICITUD_AJUSTE.LISTA_SOL_DETALLE", 0, NumeroSolicitud)
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAjuste.ListaDetalle", Err.Description
End Function

Public Function ListaProductos(ByVal vstrLocal As String, ByVal vstrBusca As String) As oraDynaset
On Error GoTo CtrlErr
      Set ListaProductos = gclsOracle.FN_Cursor("BTLPROD.PKG_SOLICITUD_AJUSTE.FN_LISTA_PRODUCTOS", 0, vstrLocal, vstrBusca)
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAjuste.ListaProductos", Err.Description
End Function

Public Function ActivaAjustes() As Double
On Error GoTo CtrlErr
      ActivaAjustes = gclsOracle.FN_Valor("BTLPROD.PKG_SOLICITUD_AJUSTE.FN_ACTIVA_AJUSTES")
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAjuste.ActivaAjustes", Err.Description
End Function

Public Function ListaEstados() As oraDynaset
    On Error GoTo CtrlErr
        Set ListaEstados = gclsOracle.FN_Cursor("BTLPROD.PKG_SOLICITUD_AJUSTE.FN_LISTA_ESTADOS", 0)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsSolicitudAjuste.ListaEstados", Err.Description
End Function

Public Function DevPermiteFraccion(ByVal TipoDevolucion As String, ByVal MotivoDevolucion As String) As String
    On Error GoTo CtrlErr
        DevPermiteFraccion = gclsOracle.FN_Valor("BTLPROD.PKG_AJUSTE.FN_DEV_FRACCION", TipoDevolucion, MotivoDevolucion)
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAjuste.DevNroSolicitud ", Err.Description
End Function
