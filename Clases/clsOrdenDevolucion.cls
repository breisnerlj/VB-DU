VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrdenDevolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ListaOD(ByVal vstrNumero As String, _
                        ByVal vstrFchIniEnv As String, _
                        ByVal vstrFchFinEnv As String, _
                        ByVal vstrFchIniVig As String, _
                        ByVal vstrFchFinVig As String, _
                        ByVal vstrEstado As String, _
                        ByVal vstrTipo As String, _
                        ByVal vstrMotivo As String, _
                        ByVal vstrCodigoLocal As String) As oraDynaset
    On Error GoTo Control

    Set ListaOD = gclsOracle.FN_Cursor("BTLPROD.PKG_ORDEN_DEVOLUCION.FN_LISTA_OD", 0, _
                                        vstrNumero, _
                                        vstrFchIniEnv, _
                                        vstrFchFinEnv, _
                                        vstrFchIniVig, _
                                        vstrFchFinVig, _
                                        vstrEstado, _
                                        vstrTipo, _
                                        vstrMotivo, _
                                        vstrCodigoLocal)

    Exit Function
Control:
    Err.Raise Err.Number, "clsOrdenDevolucion.ListaOD", Err.Description
End Function

Public Function ListaDetalleOD(ByVal vstrNumero As String, _
                               ByVal vstrCodLocal As String, _
                               Optional ByVal vstrCodLaboratorio As String = "", _
                               Optional ByVal vstrCodLinea As String = "") As oraDynaset
    On Error GoTo Control

    Set ListaDetalleOD = gclsOracle.FN_Cursor("BTLPROD.PKG_ORDEN_DEVOLUCION.FN_DETALLE_OD", 0, vstrNumero, _
                                                                                               vstrCodLocal, _
                                                                                               vstrCodLaboratorio, _
                                                                                               vstrCodLinea)
    
    Exit Function
Control:
    Err.Raise Err.Number, "clsOrdenDevolucion.ListaDetalleOD", Err.Description
End Function

Public Function ActEstado(vstrNumOrden As String, _
                          vstrCodEstado As String) As String

    Dim varValores As Variant
    Dim varIO As Variant
    
    On Error GoTo Control
    
    varValores = Array(vstrNumOrden, _
                       vstrCodEstado)
    
    varIO = Array(entrada, entrada)
    
    ActEstado = gclsOracle.SP("BTLPROD.PKG_ORDEN_DEVOLUCION.SP_ACT_ESTADO", varValores, varIO)
    
    Exit Function
Control:
    Err.Raise Err.Number, "clsOrdenDevolucion.ActEstado", Err.Description

End Function

Public Function AtenderOrden(vstrNumOrden As String, vstrCodLocal As String, _
                             vstrCodProducto As String, vstrCtdProducto As String, _
                             vstrCtdProductoFrac As String) As String

    Dim varValores As Variant
    Dim varIO As Variant
    
    On Error GoTo Control
    
    varValores = Array(vstrNumOrden, vstrCodLocal, vstrCodProducto, vstrCtdProducto, vstrCtdProductoFrac)
    
    varIO = Array(entrada, entrada, entrada, entrada, entrada)
    
    AtenderOrden = gclsOracle.SP("BTLPROD.PKG_ORDEN_DEVOLUCION.SP_ATENDER_ORDEN", varValores, varIO)
    
    Exit Function
Control:
    Err.Raise Err.Number, "clsOrdenDevolucion.AtenderOrden", Err.Description

End Function

Public Function ListaEstados(ByVal vstrCntCosto As String) As oraDynaset
    Set ListaEstados = gclsOracle.FN_Cursor("BTLPROD.PKG_ORDEN_DEVOLUCION.FN_LISTA_ESTADO", 0, vstrCntCosto)
End Function

Public Function FechaVencimiento(ByVal vstrCodProducto As String, _
                                 ByVal vstrCodAnio As String, _
                                 ByVal vstrCodMes As String) As String
    On Error GoTo Control
    
    FechaVencimiento = gclsOracle.FN_Valor("btlprod.pkg_orden_devolucion.fn_dev_fecha_venc", _
                                            vstrCodProducto, vstrCodAnio, vstrCodMes)
    
    Exit Function
Control:
    Err.Raise Err.Number, "clsOrdenDevolucion.AtenderOrden", Err.Description
End Function
