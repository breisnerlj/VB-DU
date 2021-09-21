VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMotivoDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Lista(ByVal vstrCodTipoDev As String, _
                      ByVal vstrCntCosto As String, _
                      Optional ByVal vstrFlgActivo As String = "") As oraDynaset
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_ORDEN_DEVOLUCION.FN_MOTIVOS", 0, vstrCodTipoDev, vstrCntCosto, vstrFlgActivo)
End Function


