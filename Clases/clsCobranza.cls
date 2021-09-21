VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function Lista(CIA As String, CTA_CTE As String) As OraDynaset
    On Error GoTo PASE
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_COBRANZA.FN_LISTA", 0, CIA, CTA_CTE)
    Exit Function
PASE:
    Err.Raise Err.Number, "clsCobranza", Err.Description
End Function

Public Function Grabar(ByVal CIA As String, _
                       ByVal CodigoLocal As String, _
                       ByVal CTA_CTE As String, _
                       ByVal TipoDocumento As String, _
                       ByVal NumeroDocumento As String, _
                       ByVal MonedaVenta As String, _
                       ByVal ImporteVenta As String, _
                       ByVal ImporteAbono As String, _
                       ByVal MonedaAbono As String, _
                       ByVal ImporteAbonoMoneda As String, _
                       ByVal Usuario As String _
                       ) As String
    On Error GoTo PASE
    Dim arrValores As Variant
    Dim arrDireccion As Variant
    arrValores = Array(CIA, CodigoLocal, CTA_CTE, TipoDocumento, NumeroDocumento, MonedaVenta, ImporteVenta, ImporteAbono, MonedaAbono, ImporteAbonoMoneda, Usuario)
    arrDireccion = Array(entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada)
    Grabar = gclsOracle.SP("BTLPROD.PKG_COBRANZA.SP_GRABA", arrValores, arrDireccion)
    If Grabar = "" Then MsgBox "Se Grabo Satisfactorimente", vbExclamation, App.ProductName: Exit Function
    MsgBox Grabar, vbCritical, App.ProductName
    Exit Function
PASE:
    Err.Raise Err.Number, "clsCobranza", Err.Description
End Function

