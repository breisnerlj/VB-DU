VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCElectronica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function ListaCierreDiario(ByVal strCod_Local As String) As OracleInProcServer.oraDynaset
On Error GoTo Control

    Set ListaCierreDiario = gclsOracle.FN_Cursor("BTLPROD.PKG_CAJA_ELECTRONICA.FN_LISTA_CIERRE_DIARIO", 0, strCod_Local)
    Exit Function
Control:
    Err.Raise Err.Number, "clsCElectronica.ListaCierreDiario", Err.Description
End Function


Public Function ListaCierres(ByVal strCod_Local As String, _
                             ByVal strEstado As String, _
                             ByVal strFchInicio As String, _
                             ByVal strFchFin As String) As OracleInProcServer.oraDynaset
On Error GoTo Control

    Set ListaCierres = gclsOracle.FN_Cursor("INTERFASE_SAP.PKG_CAJA_ELECTRONICA.FN_LISTA_CIERRES", 0, strCod_Local, strEstado, strFchInicio, strFchFin)
    Exit Function
Control:
    Err.Raise Err.Number, "clsCElectronica.ListaCierres", Err.Description
End Function

Public Function ListaEstado(ByVal Texto As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaEstado = gclsOracle.FN_Cursor("INTERFASE_SAP.PKG_CAJA_ELECTRONICA.FN_LISTA_ESTADOS", 0, Texto)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsCElectronica.ListaEstado", Err.Description
End Function

Public Function ListaCierreDiarioDetalle(ByVal strFchCierre As String, ByVal strCod_Local As String) As OracleInProcServer.oraDynaset
On Error GoTo Control

    Set ListaCierreDiarioDetalle = gclsOracle.FN_Cursor("INTERFASE_SAP.PKG_CAJA_ELECTRONICA.FN_LISTA_CIERRE_DIARIO_DET", 0, strFchCierre, strCod_Local)
    Exit Function
Control:
    Err.Raise Err.Number, "clsCElectronica.ListaCierreDiarioDetalle", Err.Description
End Function

Public Function ListaNuevoCierreDiarioDetalle(ByVal strFchCierre As String, ByVal strCod_Local As String) As OracleInProcServer.oraDynaset
On Error GoTo Control

    Set ListaNuevoCierreDiarioDetalle = gclsOracle.FN_Cursor("INTERFASE_SAP.PKG_CAJA_ELECTRONICA.FN_LISTA_CIERRE_DIARIO_DET_NEW", 0, strFchCierre, strCod_Local)
    Exit Function
Control:
    Err.Raise Err.Number, "clsCElectronica.ListaNuevoCierreDiarioDetalle", Err.Description
End Function


Public Function GrabaCierreDiario(ByVal vstrFchCierre As String, _
                             ByVal vstrCodLocal As String, _
                             ByVal vstrCodUsuario As String, _
                             ByVal vdblMtoTarjetas As Double, _
                             ByVal vdblMtoPos As Double, _
                             ByVal vdblCtdCajero As Integer, _
                             ByVal vbolRemesa As String _
                             ) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant
    gvarValores = Array(vstrFchCierre, _
                        vstrCodLocal, _
                        vstrCodUsuario, _
                        vdblMtoTarjetas, _
                        vdblMtoPos, _
                        vdblCtdCajero, _
                        vbolRemesa)
    
    gvarIO = Array(entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada)
                   
    GrabaCorXDoc = gclsOracle.SP("INTERFASE_SAP.PKG_CAJA_ELECTRONICA.SP_GRABA_CIERRE_DIARIO", _
                              gvarValores, _
                              gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsCElectronica.GrabaCierreDiario", Err.Description
End Function

