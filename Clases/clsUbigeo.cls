VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUbigeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ListaDepartamento() As oraDynaset
    On Error GoTo CtrlErr
    Set ListaDepartamento = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DEPARTAMENTO", 0)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsUbigeo", Err.Description
End Function

Public Function ListaProvincia(ByVal vstrCodDpto As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaProvincia = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_PROVINCIA", 0, vstrCodDpto)

    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsUbigeo", Err.Description
End Function

Public Function ListaDistrito(ByVal vstrCodDpto As String, _
                               ByVal vstrCodProv As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaDistrito = gclsOracle.FN_Cursor("BTLPROD.PKG_UBIGEO.FN_LISTA_DISTRITO", 0, vstrCodDpto, vstrCodProv)

    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsUbigeo", Err.Description
End Function

Public Function DevDistrito(ByVal CodigoPostal As String) As String
On Error GoTo handle
Dim strCodigoPostal As String
strCodigoPostal = gclsOracle.FN_Valor("BTLPROD.PKG_UBIGEO.FN_DEV_UBIGEO", CodigoPostal)
DevDistrito = gclsOracle.FN_Valor("BTLPROD.PKG_UBIGEO.fn_lista_dist_dato", Mid(strCodigoPostal, 1, 2), Mid(strCodigoPostal, 3, 2), Mid(strCodigoPostal, 5, 2))
Exit Function
handle:
    DevDistrito = ""
End Function

Public Function DevUbigeo(ByVal CodigoPostal As String) As String
On Error GoTo handle

    DevUbigeo = gclsOracle.FN_Valor("BTLPROD.PKG_UBIGEO.FN_DEV_UBIGEO", CodigoPostal)
Exit Function
handle:
    DevUbigeo = ""
End Function

