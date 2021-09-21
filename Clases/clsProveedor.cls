VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ListaRegMagistral(ByVal vstrDato As String, _
                                  ByVal vstrFlg As String, _
                                  ByVal vstrCodLocal As String) As oraDynaset
    
    On Error GoTo CntrlError
    
    Set ListaRegMagistral = gclsOracle.FN_Cursor("CMR.PKG_PROVEEDOR.FN_PROV_REG_MAGISTRAL", 0, vstrDato, vstrFlg, vstrCodLocal)
    
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProveedor", Err.Description
    
End Function

Public Function Nombre(ByVal vstrProveedor As String) As String

On Error GoTo CntrlError
    Nombre = "" & gclsOracle.FN_Valor("CMR.PKG_PROVEEDOR.FN_NOM", vstrProveedor)
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsProveedor", Err.Description
    
End Function


