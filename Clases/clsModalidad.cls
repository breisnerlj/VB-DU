VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsModalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Lista() As oraDynaset
    On Error GoTo CntrlErr
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_MODALIDAD.FN_LISTA", 0)
    
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsModalidad", Err.Description
End Function

Private Sub Class_Initialize()

End Sub
