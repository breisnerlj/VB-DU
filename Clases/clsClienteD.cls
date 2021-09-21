VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClienteD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function ListaCliente(ByVal TipoDocumento As String, _
                          ByVal NroDocumento As String, _
                          ByVal CodTarjeta As String) As OracleInProcServer.oraDynaset
On Error GoTo Control

    Set ListaCliente = gclsOracle.FN_Cursor("BTLPROD.PKG_FIDELIZADO.FN_LISTA_CLIENTE", 0, TipoDocumento, NroDocumento, CodTarjeta)
    Exit Function
Control:
    Err.Raise Err.Number, "clsClienteD.ListaCliente", Err.Description
End Function

Public Function ListaClienteBarra(ByVal strCodBarra As String) As oraDynaset
    On Error GoTo CntrlError
    Set ListaClienteBarra = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_CLIENTE_BARRA", 0, strCodBarra)
    
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsCliente.ListaClienteBarra", Err.Description
End Function

Public Function ListaClienteCMR(ByVal strCodTarjeta As String) As oraDynaset
    On Error GoTo CntrlError
    Set ListaClienteCMR = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_CLIENTE_CMR", 0, strCodTarjeta)
    
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsCliente.ListaClienteCMR", Err.Description
End Function

Public Function GrabarTarjetaCMR(ByVal strCodClienteCMR As String, _
                          ByVal StrNumTarjeta As String, _
                          ByVal CodigoUsuario As String) As String

    Dim varValores As Variant
    Dim varIO As Variant
    
    varValores = Array(strCodClienteCMR, _
                       StrNumTarjeta, _
                       CodigoUsuario)
    
    varIO = Array(entrada, _
                  entrada, _
                  entrada)
    
    GrabarTarjetaCMR = gclsOracle.SP("BTLPROD.SP_GRABA_TARJETA_CMR", varValores, varIO)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsClienteD.GrabarTarjetaCMR", Err.Description
End Function

