VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''Public Function Lista(ByVal Cia As String, vstrDato As String) As oraDynaset
'''On Error GoTo CnttlError
'''    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_LOCAL.FN_LISTA", 0, Cia, vstrDato)
'''Exit Function
'''CnttlError:
'''    Err.Raise Err.Number, "clsLocal.Lista", Err.Description
'''End Function

Public Function Lista() As oraDynaset
    On Error GoTo CnttlError
    
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_FIDELIZADO.FN_LISTA_TIPODOC", 0)
    
Exit Function
CnttlError:
    Err.Raise Err.Number, "clsLocal::Lista", Err.Description
End Function

