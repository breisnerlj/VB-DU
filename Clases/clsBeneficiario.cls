VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBeneficiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function CreditoReal(ByVal strCodConvenio As String, ByVal strCodCliente As String, ByVal strCia As String, ByVal strCodLocal As String) As Double

On Error GoTo CtrlErr

CreditoReal = gclsOracle.FN_Valor("BTLPROD.PKG_BENEFICIARIO.FN_RET_CREDITO_REAL", strCodConvenio, strCodCliente, strCia, strCodLocal)


Exit Function

CtrlErr:
    Err.Raise Err.Number, "clsBeneficiario.CreditoReal", Err.Description

End Function


Public Function Consumo(ByVal strCodConvenio As String, ByVal strCodCliente As String, ByVal strCia As String, ByVal strCodLocal As String) As Double

On Error GoTo CtrlErr

Consumo = gclsOracle.FN_Valor("BTLPROD.PKG_BENEFICIARIO.FN_RET_CONSUMO_BENEFICIARIO", strCodConvenio, strCodCliente, strCia, strCodLocal)


Exit Function

CtrlErr:
    Err.Raise Err.Number, "clsBeneficiario.Consumo", Err.Description

End Function

Public Function ConsumoUnif(ByVal strCodConvenio As String, ByVal strCodCliente As String, ByVal strCia As String, ByVal strCodLocal As String) As Double
On Error GoTo CtrlErr
    ConsumoUnif = gclsOracle.FN_Valor("BTLPROD.PKG_BENEFICIARIO.FN_RET_CONSUMO_BENF_MIFA_INK", strCodConvenio, strCodCliente, strCia, strCodLocal)
Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsBeneficiario.ConsumoUnif", Err.Description
End Function

Public Function Lista(ByVal strCodConvenio As String, ByVal strCodCliente As String) As oraDynaset



On Error GoTo CtrlErr

Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_BENEFICIARIO.FN_LISTA_TITULARES", "0", strCodConvenio, strCodCliente)


Exit Function

CtrlErr:
    Err.Raise Err.Number, "clsBeneficiario.Lista", Err.Description



End Function
