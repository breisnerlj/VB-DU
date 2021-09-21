VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ListaMotivoRechazo()
    On Error GoTo CtrlErr
    Set ListaMotivoRechazo = gclsOracle.FN_Cursor("BTLPROD.PKG_TARJETA.FN_LISTA_MOTIVO_RETRAZO", 0)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "cls_tarjeta.ListaMotivoRechazo", Err.Description
End Function
