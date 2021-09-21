VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDireccDLV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ListaDirecc() As OraDynaset
    On Error GoTo CtrlErr
    Set ListaDirecc = gclsOracle.FN_Cursor("BTLPROD.PKG_TIPO_DIRECCION.FN_LISTA", 0)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDireccDLV", Err.Description
End Function

Public Function Graba(ByVal vstrCodTipoDirecc As String, _
                      ByVal vstrDesTipoDirecc As String, _
                      ByVal vstrFlgActivo As String, _
                      ByVal vstrCodUsuario As String) As String
    On Error GoTo CtrlErr
    
    Dim varValores As Variant
    Dim varIO As Variant
    varValores = Array(vstrCodTipoDirecc, _
                       vstrDesTipoDirecc, _
                       vstrFlgActivo, _
                       vstrCodUsuario)
    
    varIO = Array(entrada, _
                  entrada, _
                  entrada, _
                  entrada)
    
    Graba = gclsOracle.SP("BTLPROD.PKG_TIPO_DIRECCION.SP_GRABA", _
                                            varValores, varIO)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDireccDLV", Err.Description
End Function

