VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Respuesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim pCodigoRespuesta  As Long
Dim PMensaje As String
Dim pCodigoAutorizacion As String
Public Property Get CodigoRespuesta() As Long
    'pCodigoRespuesta = "00"
    'CodigoRespuesta = 0
    CodigoRespuesta = pCodigoRespuesta
End Property

Public Property Let CodigoRespuesta(ByVal vNewValue As Long)
    pCodigoRespuesta = vNewValue
End Property


Public Property Get Mensaje() As String
    'PMensaje = "recargo"
    Mensaje = PMensaje
End Property

Public Property Let Mensaje(ByVal vNewValue As String)

    PMensaje = vNewValue
End Property


Public Function Atributo(ByVal tipo As String)
    Atributo = Me.CodigoAutorizacion
End Function

Public Property Get CodigoAutorizacion() As Long
    'pCodigoRespuesta = "00"
    'CodigoRespuesta = 0
    CodigoAutorizacion = pCodigoAutorizacion
End Property

Public Property Let CodigoAutorizacion(ByVal vNewValue As Long)
    pCodigoAutorizacion = vNewValue
End Property
