VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Function Solicita(ByVal IDFuncionalidad As String, _
                    ByRef OutNumeroSolicitud As String, _
                    ByRef OutMensaje As String, _
                    ByVal CodigoAprobacion As String, _
                    ByVal CodigoLocal As String, ByVal CodigoLiquidacion As String, ByVal CodigoCliente As String, _
                    ByVal CodigoDocumento As String, ByVal NumeroDocumento As String, ByVal CodigoReferencia As String, ByVal NumeroReferencia As String, _
                    ByVal MontoSolicitado As String, ByVal CodigoRemito As String, ByVal CodigoRemesa As String, ByVal CodigoSolicitante As String, _
                    ByVal DesObservacion As String, ByRef Respuesta As String, ByVal FechaDocRefer As String, ByVal MontoDocRefer As String, _
                    ByVal CodigoRefBeneficiario As String, ByVal ApellidoPaterno As String, ByVal ApellidoMaterno As String, _
                    ByVal Nombre As String, ByVal Cargo As String, ByVal NumDocumentoIdent As String, _
                    ByVal CodEstadoCivil As String, ByVal Email As String, _
                    ByVal CadenaCodigoProducto As String, _
                    ByVal CadenaCantidadProducto As String, _
                    ByVal CadenaCantidadFracProducto As String, _
                    ByVal CadenaPrecioUnitProducto As String, _
                    ByVal CadenaBaseImponible As String, _
                    ByVal CadenaImpuesto As String, _
                    ByVal CadenaExonerado As String, _
                    ByVal CadenaTotal As String, _
                    ByVal CadenaCodigoUsuario As String, ByVal CadenaTotalUsu As String, ByVal CodigoConvenio As String, ByVal CodigoCobro As String, ByVal CodigoProducto As String, _
                    ByVal CodigoAprobador As String, ByVal NumeroSolicitud1 As String, ByVal FlagCommit As String, ByVal CadenaAprobadores As String, ByVal NumeroSolicitud As String _
                    ) As String
                    
                    

On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant
    gvarValores = Array(IDFuncionalidad, OutNumeroSolicitud, OutMensaje, CodigoAprobacion, _
                    CodigoLocal, CodigoLiquidacion, CodigoCliente, CodigoDocumento, _
                    NumeroDocumento, CodigoReferencia, NumeroReferencia, MontoSolicitado, _
                    CodigoRemito, CodigoRemesa, CodigoSolicitante, DesObservacion, _
                    Respuesta, FechaDocRefer, MontoDocRefer, _
                    CodigoRefBeneficiario, ApellidoPaterno, ApellidoMaterno, _
                    Nombre, Cargo, NumDocumentoIdent, _
                    CodEstadoCivil, Email, _
                    CadenaCodigoProducto, _
                    CadenaCantidadProducto, _
                    CadenaCantidadFracProducto, _
                    CadenaPrecioUnitProducto, _
                    CadenaBaseImponible, _
                    CadenaImpuesto, _
                    CadenaExonerado, _
                    CadenaTotal, _
                    CadenaCodigoUsuario, CadenaTotalUsu, CodigoConvenio, CodigoCobro, CodigoProducto, CodigoAprobador, NumeroSolicitud1, FlagCommit, CadenaAprobadores, NumeroSolicitud)
    gvarIO = Array(entrada, Salida, Salida, entrada, _
                    entrada, entrada, entrada, entrada, _
                    entrada, entrada, entrada, entrada, _
                    entrada, entrada, entrada, entrada, _
                    Salida, entrada, entrada, entrada, _
                    entrada, entrada, entrada, entrada, _
                    entrada, entrada, entrada, entrada, _
                    entrada, entrada, entrada, entrada, _
                    entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, _
                    entrada, entrada, entrada, entrada, entrada _
                )
    Dim strme As String
    strme = gclsOracle.SP("BTLPROD.PKG_AUTORIZA.FN_GENERA_SOLICITA", gvarValores, gvarIO)
    
      OutNumeroSolicitud = gvarValores(1)
      OutMensaje = gvarValores(2)
      Respuesta = gvarValores(16)
      Solicita = strme
Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsAprobacion.Solicita", Err.Description
End Function


