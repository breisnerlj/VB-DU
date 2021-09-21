VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sModalidad As String
Private m_sTienda As String

Public Property Get Modalidad() As String
    Modalidad = m_sModalidad
End Property

Public Property Let Modalidad(ByVal sModalidad As String)
    m_sModalidad = sModalidad
End Property

Public Property Get Tienda() As String
    Tienda = m_sTienda
End Property

Public Property Let Tienda(ByVal sTienda As String)
    m_sTienda = sTienda
End Property

Public Function Devuelve_Mensaje() As String
    Devuelve_Mensaje = gclsOracle.FN_Valor("BTLPROD.PKG_MENSAJES.DEVUELVE_MENSAJE_VTA", Modalidad, Tienda)
End Function
