VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLetras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function NumAletra(ByVal vdblNum As String) As String

        On Error GoTo CtrlErr
    
        'NumAletra = gclsOracle.FN_Valor("MEDCO.PACK_RUTINAS.LETRAS", vdblNum)
        NumAletra = gclsOracle.FN_Valor("BTLPROD.PKG_RUTINA.LETRAS", vdblNum)
        
        Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsLetras", Err.Description
    
End Function
