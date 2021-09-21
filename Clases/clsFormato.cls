VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Lista() As oraDynaset

    Set Lista = gclsOracle.FN_Cursor("CMR.PKG_FORMATO.FN_LISTA", 0, "")
    
End Function



Public Function ListaxDoc(ByVal vstrTipDoc As String, _
                          ByVal vstrCabecera As String) As oraDynaset

    Set ListaxDoc = gclsOracle.FN_Cursor("CMR.PKG_FORMATO.FN_LISTA", 0, _
                            vstrTipDoc, _
                            vstrCabecera)
    
End Function



Public Function NumLineas(ByVal vstrCodFormato As String) As Integer

    NumLineas = gclsOracle.FN_Valor("CMR.PKG_FORMATO.FN_NUM_LINEAS", vstrCodFormato)
    
End Function

