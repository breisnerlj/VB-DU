VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDenominacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function ListaDenominacion(Optional Codigo As String = "") As OraDynaset

    Set ListaDenominacion = gclsOracle.FN_Cursor("BTLPROD.PKG_DENOMINACION.FN_LISTA", 0, Trim(Codigo))
End Function
