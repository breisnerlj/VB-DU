VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function lista() As OraDynaset
    On Error GoTo CtrlErr
        Set lista = gclsOracle.FN_Cursor("BTLPROD.PKG_BANCO.FN_LISTA", 0)
        
      Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsBanco", Err.Description
End Function
