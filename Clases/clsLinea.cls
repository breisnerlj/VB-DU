VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Lista(ByVal vstrCodLab As String, Optional ByVal vstrFlgActivo As String, Optional ByVal vstrTextCab As String) As oraDynaset
    Set Lista = gclsOracle.FN_Cursor("CMR.PKG_LINEA.FN_LISTA", 0, vstrCodLab, "", vstrFlgActivo, vstrTextCab)
End Function

