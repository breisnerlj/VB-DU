VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CanalVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim pHost As String
Dim pPORT As String

Function AnulacionVentaDePines()

End Function
Function VentaDePines( _
ByVal strCodigoComercio As String, _
ByVal NumeroTelefono As String, _
                                             ByVal CODIGO As String, _
                                             ByVal TIPOTARJETA As String, _
                                             ByVal MontoRecarga As String, _
                                             ByVal terminal As String, _
                                             ByVal Serie As String, _
                                             ByVal Trace As String, _
                                             ByVal strIpHost As String, _
                                             ByVal strPuertoHost As String, _
                                             ByVal strTimeOut As String, _
                                             ByVal A_ctipo_rcd_in As String _
)
Dim timeout As Double
timeout = 1000 * IIf(Val(strTimeOut) = 0, 60, strTimeOut)


VentaDePines = frmRecargasdirec.carga(Mid(terminal, 4, 3), Mid(Mid(strCodigoComercio, 6, 3) & NumeroTelefono, 1, 10), NumeroTelefono, MontoRecarga, Mid(terminal, 1, 6), terminal, A_ctipo_rcd_in)
'MsgBox VentaDePines
'VentaDePines = "QUE COSA"
End Function

Public Property Get host() As String
    host = pHost
End Property

Public Property Let host(ByVal vNewValue As String)
    pHost = vNewValue
End Property



Public Property Get PORT() As String
    PORT = pPORT
End Property

Public Property Let PORT(ByVal vNewValue As String)
    pPORT = vNewValue
End Property
Function venta(ByVal x As Peticion) As Respuesta
Dim petit As New Peticion
Set petit = x
Dim strRespuesta As String
    Dim obj  As New Respuesta
    Set venta = obj
    strRespuesta = VentaDePines(petit.terminal, petit.Destino, "", "", petit.Monto, petit.terminal, "", petit.terminal, "", "", "60", petit.Producto)
    obj.CodigoRespuesta = IIf(fncPalote(fncPalote(strRespuesta, 1, "|")) = "", 99, fncPalote(fncPalote(strRespuesta, 1, "|")))
    obj.Mensaje = IIf(fncPalote(fncPalote(strRespuesta, 1, "|"), 1, "|") = "", "No se encontro respuesta del proveedor", fncPalote(fncPalote(strRespuesta, 1, "|"), 1, "|"))
    obj.CodigoAutorizacion = IIf(fncPalote(strRespuesta, 0, "|") = "", 0, fncPalote(strRespuesta, 0, "|"))
    
    
End Function




