VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMoneda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Lista(Optional vstrCodMoneda As String) As OraDynaset
    On Error GoTo CtrlErr
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_MONEDA.LISTA", 0, vstrCodMoneda)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMoneda", Err.Description
End Function

Public Function Graba(ByVal vstrCodMoneda As String, _
                      ByVal vstrDesMoneda As String, _
                      ByVal vstrSmbMoneda As String, _
                      ByVal vstrDesLgMoneda As String, _
                      ByVal vstrFlgMoneda As String, _
                      ByVal vstrCodUsuario As String)

    On Error GoTo CtrlErr
    Dim gvarValores As Variant
    Dim gvarIO  As Variant
    gvarValores = Array(vstrCodMoneda, _
                        vstrDesMoneda, _
                        vstrSmbMoneda, _
                        vstrDesLgMoneda, _
                        vstrFlgMoneda, _
                        vstrCodUsuario)
    
    gvarIO = Array(entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada)
                   
    Graba = gclsOracle.SP("BTLPROD.PKG_MONEDA.SP_GRABA", _
                              gvarValores, _
                              gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsMoneda_Graba", Err.Description
    

End Function

Public Function TCambio(ByVal Tipo As String, ByVal Moneda As String) As Double

On Error GoTo CtrlErr
    TCambio = gclsOracle.FN_Valor("BTLPROD.PKG_MONEDA.FN_BTL_TCAMBIO", Tipo, Moneda)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsMoneda.Cambio", Err.Description
End Function

Public Function MaxCodMoneda()
    On Error GoTo CtrlErr
    MaxCodMoneda = gclsOracle.FN_Valor("BTLPROD.PKG_MONEDA.FN_DEV_COD_MONEDA")
    
    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsMoneda.MaxCodMoneda", Err.Description
End Function
