VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTarjetaBean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarDNI As String
Private mvarTipoDNI As String
Private mvarIdTransaccion As String
Private mvarNumeroTarjeta As String
Private mvarCodigoAutororizacion As String
Private mvarNumeroAutororizacion As String
Private mvarMensaje As String
Private mvarEstadoTarjeta As String
Private mvarEstadoOperacion As String
Private mvarListaPedienteCanje As XArrayDB
Private mvarListaInscritos As XArrayDB
Private mvarListaProductosVendidos As XArrayDB
Private mvarListaBonificados As XArrayDB
Private mvarPuntosTotalAcumulados As Double
Private mvarPuntosAcumulados As Double
Private mvarPuntosRedimidos As Double
Private mvarNombreCompleto As String
Private mvarFecha As String
Private mvarDeslizaTarjeta As Boolean
Private mvarEscaneaDNI As Boolean
Private mvarAhorroTotal As Double
Private oFarmaPuntosConstante As clsFPConstante

Public Property Let DeslizaTarjeta(ByVal vData As Boolean)
    mvarDeslizaTarjeta = vData
End Property

Public Property Get DeslizaTarjeta() As Boolean
    DeslizaTarjeta = mvarDeslizaTarjeta
End Property

Public Property Let DNI(ByVal vData As String)
    ' Convetir en numerico
    If Not vData = "" Then
        If Len(vData) = 10 And InStr(1, vData, oFarmaPuntosConstante.TipoDocIdentidad.DNI) = 1 Then
            mvarDNI = Mid$(vData, 2, 10)
            mvarTipoDNI = "002"
        ElseIf Len(vData) = 10 And InStr(1, vData, oFarmaPuntosConstante.TipoDocIdentidad.CE) = 1 Then
            mvarDNI = Mid$(vData, 1, 10)
            mvarTipoDNI = "004"
        Else
            mvarDNI = vData
            mvarTipoDNI = IIf(Len(vData) = 8, "002", "004")
        End If
    Else
        mvarDNI = ""
        mvarTipoDNI = ""
    End If
End Property

Public Property Get DNI() As String
    DNI = mvarDNI
End Property

Public Property Get TipoDNI() As String
    TipoDNI = mvarTipoDNI
End Property

Public Property Let IdTransaccion(ByVal vData As String)
    mvarIdTransaccion = vData
End Property

Public Property Get IdTransaccion() As String
    IdTransaccion = mvarIdTransaccion
End Property

Public Property Let NumeroTarjeta(ByVal vData As String)
    'convetir en numerico
    If Not (vData = "") Then
        If (Len(vData) = 10 And InStr(1, vData, oFarmaPuntosConstante.TipoDocIdentidad.DNI)) Then
            mvarNumeroTarjeta = Mid$(vData, 2, 10)
            mvarTipoDNI = "002"
        ElseIf (Len(vData) = 10 And InStr(1, vData, oFarmaPuntosConstante.TipoDocIdentidad.CE)) Then
            mvarNumeroTarjeta = Mid$(vData, 1, 10)
            mvarTipoDNI = "003"
        Else
            mvarNumeroTarjeta = vData
            mvarTipoDNI = ""
        End If
    Else
        NumeroTarjeta = ""
        mvarTipoDNI = ""
    End If
End Property

Public Property Get NumeroTarjeta() As String
    NumeroTarjeta = mvarNumeroTarjeta
End Property

Public Property Let CodigoAutororizacion(ByVal vData As String)
    mvarCodigoAutororizacion = vData
End Property

Public Property Get CodigoAutororizacion() As String
    CodigoAutororizacion = mvarCodigoAutororizacion
End Property

Public Property Let NumeroAutororizacion(ByVal vData As String)
    mvarNumeroAutororizacion = vData
End Property

Public Property Get NumeroAutororizacion() As String
    NumeroAutororizacion = mvarNumeroAutororizacion
End Property

Public Property Let Mensaje(ByVal vData As String)
    mvarMensaje = vData
End Property

Public Property Get Mensaje() As String
    Mensaje = mvarMensaje
End Property

Public Property Let ListaPedienteCanje(ByVal vData As XArrayDB)
    Set mvarListaPedienteCanje = vData
End Property

Public Property Get ListaPedienteCanje() As XArrayDB
    Set ListaPedienteCanje = mvarListaPedienteCanje
End Property

Public Property Let ListaInscritos(ByVal vData As XArrayDB)
    Set mvarListaInscritos = vData
End Property

Public Property Get ListaInscritos() As XArrayDB
    Set ListaInscritos = mvarListaInscritos
End Property

Public Property Let ListaBonificados(ByVal vData As XArrayDB)
    Set mvarListaBonificados = vData
End Property

Public Property Get ListaBonificados() As XArrayDB
    Set ListaBonificados = mvarListaBonificados
End Property

Public Property Let PuntosTotalAcumulados(ByVal vData As Double)
    mvarPuntosTotalAcumulados = vData
End Property

Public Property Get PuntosTotalAcumulados() As Double
    PuntosTotalAcumulados = mvarPuntosTotalAcumulados
End Property

Public Property Let PuntosAcumulados(ByVal vData As Double)
    mvarPuntosAcumulados = vData
End Property

Public Property Get PuntosAcumulados() As Double
    PuntosAcumulados = mvarPuntosAcumulados
End Property

Public Property Let PuntosRedimiedos(ByVal vData As Double)
    mvarPuntosRedimidos = vData
End Property

Public Property Get PuntosRedimidos() As Double
    PuntosRedimidos = mvarPuntosRedimidos
End Property

Public Property Let NombreCompleto(ByVal vData As String)
    If Not "" & vData = "" Then
        mvarNombreCompleto = Replace$(vData, "@", " ")
    Else
        mvarNombreCompleto = vData
    End If
End Property

Public Property Get NombreCompleto() As String
    NombreCompleto = mvarNombreCompleto
End Property

Public Property Let Fecha(ByVal vData As String)
    mvarFecha = vData
End Property

Public Property Get Fecha() As String
    Fecha = mvarFecha
End Property

Public Property Let EstadoTarjeta(ByVal vData As String)
    mvarEstadoTarjeta = vData
End Property

Public Property Get EstadoTarjeta() As String
    EstadoTarjeta = mvarEstadoTarjeta
End Property

Public Property Let EstadoOperacion(ByVal vData As String)
    mvarEstadoOperacion = vData
End Property

Public Property Get EstadoOperacion() As String
    EstadoOperacion = mvarEstadoOperacion
End Property

Public Property Let ListaProductosVendidos(ByVal vData As XArrayDB)
    Set mvarListaProductosVendidos = vData
End Property

Public Property Get ListaProductosVendidos() As XArrayDB
    Set ListaProductosVendidos = mvarListaProductosVendidos
End Property

Public Property Let EscaneaDNI(ByVal vData As Boolean)
    mvarEscaneaDNI = vData
End Property

Public Property Get EscaneaDNI() As Boolean
    EscaneaDNI = mvarEscaneaDNI
End Property

' Total de redimidos en los ultimos 12 meses
Public Property Let AhorroTotal(ByVal vData As Double)
    mvarAhorroTotal = vData
End Property

Public Property Get AhorroTotal() As Double
    AhorroTotal = mvarAhorroTotal
End Property

Private Sub Class_Initialize()
    Set oFarmaPuntosConstante = New clsFPConstante
End Sub
