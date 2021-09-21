VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaximoManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim strCodigo As String
Dim strDescripcion As String

Public Property Get Codigo() As String
    Codigo = strCodigo
End Property

Public Property Let Codigo(ByVal vNewValue As String)
    strCodigo = vNewValue
End Property

Public Property Get Descripcion() As String
    Descripcion = strDescripcion
End Property

Public Property Let Descripcion(ByVal vNewValue As String)
    strDescripcion = vNewValue
End Property

Function Lista() As oraDynaset
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_MOTIVO_SMM.FN_LISTA_MAXIMO_MANUAL", 0)
End Function


