VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''Public Function ListaPedidoDLV(ByVal vstrCia As String, _
''                               ByVal vstrCodLocal As String) As oraDynaset
''
''    On Error GoTo CtrlErr
''    Set ListaPedidoDLV = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA.LISTA_PEDIDO_DLV", 0, vstrCia, vstrCodLocal)
''
''    Exit Function
''CtrlErr:
''    Err.Raise Err.Number, "clsDelivery", Err.Description
''End Function

Public Function ValidaLlevando(ByVal CodLocal As String, _
                    ByVal NumProforma As String, _
                    ByVal CodEstado As String, _
                    ByVal ApproxCount As Long) As Boolean
Dim bolPasaValidacion As Boolean
Dim objPedido As clsProforma
Dim vstrMensaje As String

On Error GoTo CtrlErr
bolPasaValidacion = True
Set objPedido = New clsProforma
objUsuario.CodigoMotorizado = objUsuario.DevMotorizado(objUsuario.CodigoEmpresa, CodLocal, NumProforma)
If CodEstado = objPedido.PedidoAvisado And objUsuario.CodigoMotorizado = "" Then
   MsgBox "Se tiene que asignar un Motorizado", vbOKOnly + vbExclamation, App.ProductName
   bolPasaValidacion = False
Else
    If objUsuario.CodigoMotorizado <> "" Then
       vstrMensaje = objUsuario.GrabaEstadoMotorizado(objUsuario.CodigoMotorizado, _
                                                      objPedido.PedidoLlevando, _
                                                      objUsuario.Codigo)
       If vstrMensaje <> "" Then MsgBox vstrMensaje, vbCritical, App.ProductName
    End If
    If CodEstado = objPedido.PedidoAvisado Then
        If MsgBox("Desea cambiar a Llevando sin necesidad que el pedido este Proformado ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirme") = vbNo Then
            If (ApproxCount = 0) Or (CodEstado <> objPedido.PedidoProforma) Then bolPasaValidacion = False
        End If
    ElseIf CodEstado <> objPedido.PedidoProforma Then
        MsgBox "No se puede cambiar el estado al pedido, debe estar en proforma (PRO) o avisado (AVI)", vbOKOnly + vbExclamation, "Error"
        bolPasaValidacion = False
    End If
End If
Set objPedido = Nothing
ValidaLlevando = bolPasaValidacion
Exit Function

CtrlErr:
    Err.Raise Err.Number, "clsDelivery.ValidaLlevando", Err.Description

End Function


Public Function ValidaLlegadaDestino(ByVal CodEstado As String, _
                    ByVal ApproxCount As Long) As Boolean
Dim bolPasaValidacion As Boolean
Dim objPedido As clsProforma

On Error GoTo CtrlErr
bolPasaValidacion = True
Set objPedido = New clsProforma

If ApproxCount = 0 Or CodEstado <> objPedido.PedidoLlevando Then bolPasaValidacion = False

Set objPedido = Nothing
ValidaLlegadaDestino = bolPasaValidacion
Exit Function

CtrlErr:
    Err.Raise Err.Number, "clsDelivery.ValidaLlegadaDestino", Err.Description

End Function

Public Function ValidaEntregado(ByVal CodLocal As String, _
                    ByVal NumProforma As String, _
                    ByVal CodEstado As String, _
                    ByVal ApproxCount As Long) As Boolean
Dim bolPasaValidacion As Boolean
Dim objPedido As clsProforma
Dim vstrMensaje As String


On Error GoTo CtrlErr
bolPasaValidacion = True
Set objPedido = New clsProforma

     objUsuario.CodigoMotorizado = objUsuario.DevMotorizado(objUsuario.CodigoEmpresa, CodLocal, NumProforma)
     
     If objUsuario.CodigoMotorizado <> "" Then
        vstrMensaje = objUsuario.GrabaEstadoMotorizado(objUsuario.CodigoMotorizado, _
                                                       objPedido.PedidoEntregado, _
                                                       objUsuario.Codigo)
        If vstrMensaje <> "" Then MsgBox vstrMensaje, vbCritical, App.ProductName: bolPasaValidacion = False
     End If
    If ApproxCount = 0 Or CodEstado <> objPedido.PedidoLlegada Then bolPasaValidacion = False

Set objPedido = Nothing
ValidaEntregado = bolPasaValidacion
Exit Function

CtrlErr:
    Err.Raise Err.Number, "clsDelivery.ValidaEntregado", Err.Description

End Function

Public Function ValidaLlegadaLocal(ByVal CodLocal As String, _
                    ByVal NumProforma As String, _
                    ByVal CodEstado As String, _
                    ByVal ApproxCount As Long) As Boolean
Dim bolPasaValidacion As Boolean
Dim objPedido As clsProforma
Dim vstrMensaje As String


On Error GoTo CtrlErr
bolPasaValidacion = True
Set objPedido = New clsProforma


     objUsuario.CodigoMotorizado = objUsuario.DevMotorizado(objUsuario.CodigoEmpresa, CodLocal, NumProforma)
     
     If objUsuario.CodigoMotorizado <> "" Then vstrMensaje = objUsuario.GrabaEstadoMotorizado(objUsuario.CodigoMotorizado, _
                                                       objPedido.PedidoLlegadaLocal, _
                                                       objUsuario.Codigo)
    If vstrMensaje <> "" Then MsgBox vstrMensaje, vbCritical, App.ProductName: bolPasaValidacion = False
    If ApproxCount = 0 Or CodEstado <> objPedido.PedidoEntregado Then bolPasaValidacion = False

Set objPedido = Nothing
ValidaLlegadaLocal = bolPasaValidacion
Exit Function

CtrlErr:
    Err.Raise Err.Number, "clsDelivery.ValidaLlegadaLocal", Err.Description

End Function

Public Function ValidaLiberar(ByVal CodLocal As String, _
                    ByVal NumProforma As String, _
                    ByVal CodEstado As String, _
                    ByVal ApproxCount As Long) As Boolean
Dim bolPasaValidacion As Boolean
Dim objPedido As clsProforma
Dim vstrMensaje As String


On Error GoTo CtrlErr
bolPasaValidacion = True
Set objPedido = New clsProforma


    objUsuario.CodigoMotorizado = objUsuario.DevMotorizado(objUsuario.CodigoEmpresa, CodLocal, NumProforma)
     
    If objUsuario.CodigoMotorizado <> "" Then vstrMensaje = objUsuario.GrabaEstadoMotorizado(objUsuario.CodigoMotorizado, _
                                                       objPedido.PedidoLiberado, _
                                                       objUsuario.Codigo)
    If vstrMensaje <> "" Then MsgBox vstrMensaje, vbCritical, App.ProductName: bolPasaValidacion = False
    
    
    If CodEstado = objPedido.PedidoEntregado Then bolPasaValidacion = False


Set objPedido = Nothing
ValidaLiberar = bolPasaValidacion
Exit Function

CtrlErr:
    Err.Raise Err.Number, "clsDelivery.ValidaLiberar", Err.Description

End Function

Public Function ListaPedido_x_Tarjetas(ByVal vstrFchIni As String, _
                                       ByVal vstrFchFin As String, _
                                       ByVal vstrCodTarjeta As String, _
                                       ByVal vstrNroTarjeta As String)
On Error GoTo CtrlErr

    Set ListaPedido_x_Tarjetas = gclsOracle.FN_Cursor("BTLPROD.PKG_REPO_DLV.FN_LISTA_TARJETAS_PEDIDO", 0, vstrFchIni, vstrFchFin, vstrCodTarjeta, vstrNroTarjeta)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDelivery.ListaPedido_x_Tarjetas", Err.Description
End Function
