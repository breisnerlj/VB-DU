VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServPrestados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ListaConcepto(ByVal vstrCodBtl As String) As OracleInProcServer.OraDynaset
    
    On Error GoTo PASE
    Set ListaConcepto = gclsOracle.FN_Cursor("BTLPROD.PKG_SERV_PRESTADO.FN_LISTA_CONCEPTO", 0, vstrCodBtl)
    
    Exit Function
PASE:
    Err.Raise Err.Number, "clsServPrestados", Err.Description
End Function

Public Function ListaSubConcepto(ByVal vstrCodBtl$, ByVal vstrCodConcept$)
    On Error GoTo PASE
    Set ListaSubConcepto = gclsOracle.FN_Cursor("BTLPROD.PKG_SERV_PRESTADO.FN_LISTA_SUBCONCEPTO", 0, vstrCodBtl, vstrCodConcept)
    
    Exit Function
PASE:
    Err.Raise Err.Number, "clsServPrestados", Err.Description
End Function

Public Function ListaProducto(ByVal vstrSConcep$)
    On Error GoTo PASE
    
    Set ListaProducto = gclsOracle.FN_Cursor("BTLPROD.PKG_SERV_PRESTADO.FN_LISTA_PRODUCTO", 0, vstrSConcep)
    
    Exit Function
PASE:
    Err.Raise Err.Number, "clsServPrestados", Err.Description
End Function

