VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocumentoPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Lista(Optional CodigoLocal As String = "") As oraDynaset
    On Error GoTo CtrlError
    
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO_DSCTO.FN_LISTA_VALE_DSCTO", 0, CodigoLocal)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsDocumentoPago", Err.Description
End Function

Public Function valida(ByVal FormaPago As String, ByVal FormaPagoHijo As String) As oraDynaset
    On Error GoTo CtrlError

    Set valida = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO_DSCTO.FN_VALIDA", 0, FormaPago, FormaPagoHijo)
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsDocumentoPago", Err.Description
End Function
Public Function Busca(ByVal vstrNumDoc As String) As oraDynaset
    On Error GoTo CtrlError
    
    Set Busca = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO_DSCTO.FN_DOC_DSCTO", 0, vstrNumDoc)

    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsDocumentoPago", Err.Description
End Function

Public Function validavale(ByVal vstrNumDoc As String) As String
    On Error GoTo CtrlError
    gclsOracle.Num_Intentos = 1
    validavale = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO_DSCTO.FN_VALIDA_VALE", vstrNumDoc)
    
Exit Function
CtrlError:
    Err.Raise Err.Number, "clsDocumentoPago", Err.Description
End Function
Public Function DetalleDocumento(ByVal NumeroDocumento As String, ByVal CodigoDocumento As String) As oraDynaset
    On Error GoTo CtrlError
    Set DetalleDocumento = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO_DSCTO.FN_LISTA_DETALLE", 0, NumeroDocumento, CodigoDocumento)
Exit Function
CtrlError:
    Err.Raise Err.Number, "clsDocumentoPago", Err.Description
End Function

Public Function buscaFun(ByVal codBarra As String) As String
    On Error GoTo CtrlError
    gclsOracle.Num_Intentos = 1
    buscaFun = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO_DSCTO.FN_BUSCA_FUNC", codBarra)
    
Exit Function
CtrlError:
    Err.Raise Err.Number, "clsDocumentoPago", Err.Description
End Function

