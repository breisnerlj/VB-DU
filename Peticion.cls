VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Peticion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim pClaveCanal As String
Dim pProducto As String
Dim pterminal As String
Dim pEsPasswordTerminal As String
Dim pFecha As String
Dim pPassword As String
Dim pDestino As String
Public Monto As Double

Public Property Get ClaveCanal() As String
    ClaveCanal = pClaveCanal
End Property

Public Property Let ClaveCanal(ByVal vNewValue As String)
    pClaveCanal = vNewValue
End Property

''terminal

Public Property Get terminal() As String
    terminal = pterminal
End Property

Public Property Let terminal(ByVal vNewValue As String)
    pterminal = vNewValue
End Property


Public Property Get Producto() As String
    Producto = pProducto
End Property

Public Property Let Producto(ByVal vNewValue As String)
    pProducto = vNewValue
End Property



'req.Destino = NumeroTelefonico '"01991234567"
        'req.Password = NombreMaquina 'Me.ParametroValor("PASSWORDBS") '"btlvb6test"
        'req.SetMonto Importe '10
        'req.EsPasswordTerminal = Me.ParametroValor("FLGPASSBS") '0
        'req.Fecha

Public Property Get Password() As String
    Password = pPassword
End Property

Public Property Let Password(ByVal vNewValue As String)
    pPassword = vNewValue
End Property

Public Property Get EsPasswordTerminal() As String
    EsPasswordTerminal = pEsPasswordTerminal
End Property

Public Property Let EsPasswordTerminal(ByVal vNewValue As String)
    pEsPasswordTerminal = vNewValue
End Property

Public Property Get Fecha() As String
    Fecha = pFecha
End Property

Public Property Let Fecha(ByVal vNewValue As String)
    pFecha = vNewValue
End Property


Public Property Get Destino() As String
    Destino = pDestino
End Property

Public Property Let Destino(ByVal vNewValue As String)
    pDestino = vNewValue
End Property



Public Sub SetMonto(ByVal Importe As Double)
    Monto = Importe
End Sub
