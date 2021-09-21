VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ListaMovimientos(ByVal vstrCodBtl As String) As oraDynaset
    On Error GoTo CtrlErr
       Set ListaMovimientos = gclsOracle.FN_Cursor("BTLPROD.PKG_KARDEX_LOCAL.FN_LISTA_MOVIMIENTOS", 0, vstrCodBtl)
       
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsKardex", Err.Description
End Function

Public Function ListaOrigenDestino(ByVal vstrCodBtl As String) As oraDynaset
    On Error GoTo CtrlErr
        Set ListaOrigenDestino = gclsOracle.FN_Cursor("BTLPROD.PKG_KARDEX_LOCAL.FN_LISTA_ORIGEN_DESTINO", 0, vstrCodBtl)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsKardex", Err.Description
End Function

Public Function ListaProducto(ByVal vstrCodProd As String) As oraDynaset
    On Error GoTo CtrlErr
        Set ListaProducto = gclsOracle.FN_Cursor("BTLPROD.PKG_KARDEX_LOCAL.FN_LISTA_PRODUCTO", 0, vstrCodProd)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsKardex", Err.Description
End Function

Public Function Lista(ByVal vstrCodProd As String, _
                      ByVal vstrCodBtl As String, _
                      ByVal vstrCodMov As String, _
                      ByVal vstrCodLocalRef As String, _
                      ByVal vstrFchMovIni As String, _
                      ByVal vstrFchMovFin As String, _
                      ByVal strOrden As String) As oraDynaset
    On Error GoTo CtrlErr
        Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_KARDEX_LOCAL.LISTA", 0, vstrCodProd, vstrCodBtl, vstrCodMov, vstrCodLocalRef, vstrFchMovIni, vstrFchMovFin, strOrden)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsKardex", Err.Description
End Function

Public Function BuscaStock(ByVal vstrCodProducto As String, _
                           ByVal vstrCodZona As String, _
                           ByVal vstrUF As String) As String
    On Error GoTo CtrlErr
    BuscaStock = gclsOracle.FN_Valor("NUEVO.FN_BTL_STOCK", vstrCodProducto, vstrCodZona, vstrUF)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsKardex", Err.Description
End Function
