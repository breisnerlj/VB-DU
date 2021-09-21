VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim gvarValores As Variant
Dim gvarIO As Variant
Dim gvarError As Variant


Public Function oraClaseConvenio(ByVal oradb As OraDatabase) As oraDynaset

Dim OraSqlStmt As OraSqlStmt

    oradb.Parameters.Remove "P_RET_CURSOR"
    
    
    oradb.Parameters.Add "P_RET_CURSOR", 0, ORAPARM_OUTPUT
    oradb.Parameters("P_RET_CURSOR").serverType = ORATYPE_CURSOR
    
    
                   
    Set OraSqlStmt = oradb.CreateSql("BEGIN :P_RET_CURSOR := BTLCADENA.PKG_PAGO_TARJETA.FN_RET_TARJETA;END;", ORASQL_FAILEXEC)
    
    'Set oraTarjeta = oradb.Parameters("P_RET_CURSOR").Value
    
    Exit Function
Error:
    Err.Raise Err.Number, "clsTarjeta.oraClaseConvenio", Err.Description


End Function


Public Function Lista(Optional strCodConvenio As String) As oraDynaset
    On Error GoTo CtrlErr
    If strCodConvenio = "" Then
        Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_LISTA", 0, strCodConvenio, 1)
    Else
        Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_LISTA", 0, strCodConvenio)
    End If

    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsConvenio_Lista", Err.Description
End Function



Public Function ListaDocVerif(ByVal strCodConvenio As String) As oraDynaset

On Error GoTo ctrlErrCnv
Set ListaDocVerif = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_LISTA_DOC_VERIF", 0, strCodConvenio)
Exit Function

ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.ListaDocVerif", Err.Description

End Function

Public Function ListaDocVerificacion(Optional CodDocumento As String = "")
On Error GoTo ctrlErrCnv
Set ListaDocVerificacion = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_LISTA_DOC_VER", 0, CodDocumento)
Exit Function

ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.ListaDocVerificacion", Err.Description
End Function

Public Function ListaClase(Optional CodClase As String = "") As oraDynaset
On Error GoTo ctrlErrCnv
Set ListaClase = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_LISTA_CLASE", 0, CodClase)
Exit Function

ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.ListaClase", Err.Description
End Function
Public Function ListaBeneficiario(ByVal CIA As String, strCodConvenio As String, ByVal strCriterio As String, ByVal Estado As String) As oraDynaset


On Error GoTo ctrlErrCnv
Set ListaBeneficiario = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_LISTA_BENEFICIARIO", 0, CIA, strCodConvenio, strCriterio, IIf(IsNumeric(strCriterio) = True, 1, 0), Estado)
Exit Function

ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.ListaBeneficiario", Err.Description

End Function


Public Function ListaRepartidor(ByVal strCodConvenio As String) As oraDynaset

On Error GoTo ctrlErrCnv
Set ListaRepartidor = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_LISTA_REPARTIDOR", 0, strCodConvenio)
Exit Function

ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.ListaRepartidor", Err.Description


End Function


Public Function ListaMedico(ByVal strCodConvenio As String) As oraDynaset

On Error GoTo ctrlErrCnv
Set ListaMedico = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_LISTA_MEDICO", 0, strCodConvenio)
Exit Function

ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.ListaMedico", Err.Description


End Function



Public Function ListaTipoCampo(ByVal strCodConvenio As String) As oraDynaset

On Error GoTo ctrlErrCnv
Set ListaTipoCampo = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_LISTA_TIPO_CAMPO", 0, strCodConvenio)
Exit Function

ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.ListaTipoCampo", Err.Description


End Function

Public Function ListaXLocal(ByVal strCia As String, ByVal strCodLocal As String, Optional ByVal Nombre As String = "") As oraDynaset

On Error GoTo ctrlErrCnv
Set ListaXLocal = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_LISTA_X_LOCAL", 0, strCia, strCodLocal, Nombre)
Exit Function

ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.ListaXLocal", Err.Description


End Function

Public Function ListaDignostico(ByVal strCriterio As String) As oraDynaset
On Error GoTo ctrlErrCnv
Set ListaDignostico = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_LISTA_DIAGNOSTICO", 0, strCriterio)
Exit Function
ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.ListaDignostico", Err.Description
End Function


Public Function ListaPacienteRimac(ByVal Criterio As String, ByVal CodigoLocal, ByVal CodigoConvenio As String) As oraDynaset
On Error GoTo ctrlErrCnv
Criterio = Replace(Criterio, " ", "%")
Set ListaPacienteRimac = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.SP_LISTA_PACIENTE_UM", 0, Criterio, CodigoLocal, CodigoConvenio)
Exit Function
ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.ListaPacienteRimac", Err.Description
End Function

Public Function ListaCabeceraReceta(ByVal TipoDocumento As String, ByVal NumeroDocumento As String) As oraDynaset
On Error GoTo ctrlErrCnv
Set ListaCabeceraReceta = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_LISTA_CAB_RECETA", 0, TipoDocumento, NumeroDocumento)
Exit Function
ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.ListaXLocal", Err.Description
End Function
Public Function ListaDetalleReceta(ByVal TipoDocumento As String, ByVal NumeroDocumento As String) As oraDynaset
On Error GoTo ctrlErrCnv
Set ListaDetalleReceta = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_LISTA_DET_RECETA", 0, TipoDocumento, NumeroDocumento)
Exit Function
ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.ListaCabeceraReceta", Err.Description
End Function


Public Function VtaCnv_x_Dlv_Empresa_Asume(ByVal vstrCia As String, _
                                           ByVal vstrCodConvenio As String, _
                                           ByVal vstrFlgActivo As String, _
                                           ByVal vstrFlgTipoCnv As String) As Double
On Error GoTo ctrlErrCnv
    VtaCnv_x_Dlv_Empresa_Asume = gclsOracle.FN_Valor("BTLPROD.PKG_CONVENIO.FN_VTA_CNV_X_DLV_EMPRESA_ASUME", _
                                                      vstrCia, _
                                                      vstrCodConvenio, _
                                                      vstrFlgActivo, _
                                                      vstrFlgTipoCnv)
                                                      
    Exit Function
ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.VtaCnv_x_Dlv_Empresa_Asume", Err.Description
End Function

Public Function ValorPctEmpresa() As Integer
On Error GoTo ctrlErrCnv
    ValorPctEmpresa = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_VTA_CNV_X_DLV_100_EMPRESA")
    Exit Function
ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.ValorPctEmpresa", Err.Description
End Function



Public Function Lote(ByVal strCodModalidadVenta As String, _
                    Optional ByVal strCodConvenio As String = "") As String
On Error GoTo ctrlErrCnv
    Lote = gclsOracle.FN_Valor("BTLPROD.PKG_CONVENIO.FN_LOTE", strCodModalidadVenta, strCodConvenio)
    Exit Function
ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.Lote", Err.Description
End Function

Public Function ListaconPetitorio() As oraDynaset
On Error GoTo ctrlErrCnv
    Set ListaconPetitorio = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_LISTA_C_PETITORIO", 0)
Exit Function
ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.ListaconPetitorio", Err.Description
End Function

Public Function FnDev_ImprimeImp(ByVal vstrCodConvenio As String)
On Error GoTo CtrlErr
    FnDev_ImprimeImp = gclsOracle.FN_Valor("BTLPROD.PKG_CONVENIO.FN_FLG_IMPRIME_IMP", vstrCodConvenio)

Exit Function
CtrlErr:
  Err.Raise Err.Number, "clsConvenio", Err.Description
End Function

Public Function FnDev_Deducible(ByVal vstrCodConvenio As String)
On Error GoTo CtrlErr
    FnDev_Deducible = gclsOracle.FN_Valor("BTLPROD.PKG_CONVENIO.FN_FLG_PRECIO_DEDUCIBLE", vstrCodConvenio)

Exit Function
CtrlErr:
  Err.Raise Err.Number, "clsConvenio", Err.Description
End Function
Public Function EsRimac(ByVal CodigoConvenio As String) As Boolean
Dim intFlag As Integer
On Error GoTo CtrlErr
    
    EsRimac = IIf(gclsOracle.FN_Valor("BTLPROD.PKG_CONVENIO.FN_FLG_RIMAC", CodigoConvenio) = 1, True, False)

Exit Function
CtrlErr:
  Err.Raise Err.Number, "clsConvenio", Err.Description
End Function



Public Function TransfierePaciente(ByVal CodigoConvenio As String, _
                                    ByVal Dato As String, _
                                  ByVal Nombre As String) As oraDynaset
On Error GoTo ctrlErrCnv
    Set TransfierePaciente = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_TRANS_PAC_RIMAC", 0, objUsuario.CodigoEmpresa, objUsuario.Codigo, CodigoConvenio, Dato, Nombre)
    
Exit Function
ctrlErrCnv:
    Err.Raise Err.Number, "clsConvenio.TransfierePaciente", Err.Description
End Function

Public Function EsDataRimac(ByVal CodigoConvenio As String) As Boolean
Dim intFlag As Integer
On Error GoTo CtrlErr
    
    EsDataRimac = IIf(gclsOracle.FN_Valor("BTLPROD.PKG_CONVENIO.FN_FLG_DATA_RIMAC", CodigoConvenio) = 1, True, False)

Exit Function
CtrlErr:
  Err.Raise Err.Number, "clsConvenio", Err.Description
End Function

Public Function ValidaConvenioBTL(ByVal CodigoConvenio As String, ByVal RucEmpresa As String, ByVal CodigoCliente As String, ByVal CodigoUsuario As String) As Integer
Dim intFlag As Integer
On Error GoTo CtrlErr
    ValidaConvenioBTL = "" & gclsOracle.FN_Valor("BTLPROD.PKG_CONVENIO.FN_VALIDA_CONVENIO_BTL", CodigoConvenio, RucEmpresa, CodigoCliente, CodigoUsuario)
Exit Function
CtrlErr:
  Err.Raise Err.Number, "clsConvenio.ValidaConvenioBTL", Err.Description
End Function

Public Function ListaZonal(ByVal CodigoZonal As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaZonal = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_LISTA_ZONAL", 0, CodigoZonal)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsConvenio.ListEditionDatosAdic", Err.Description
End Function

Public Function GrabaBenefiarioCnv(ByVal vstrCia As String, _
                                   ByVal vstrCodConvenio As String, _
                                   ByVal vstrCodCliente As String, _
                                   ByVal vstrDni As String, _
                                   ByVal vstrApePaterno As String, _
                                   ByVal vstrApeMaterno As String, _
                                   ByVal vstrNombres As String, _
                                   ByVal vstrCargo As String, _
                                   ByVal vdblImpLinCred As Double, _
                                   ByVal vstrFlgCamTemLinCred As String, _
                                   ByVal vdblImpLinTmp As Double, _
                                   ByVal vstrFchIniNvaLinCred As String, _
                                   ByVal vstrFchFinNvaLinCred As String, _
                                   ByVal vstrDesObs As String, _
                                   ByVal vstrCodEmpleado As String, _
                                   ByVal vstrFlgActivo As String, _
                                   ByVal vstrCodUsuario As String, _
                                   ByVal vstrCodEstadoCivil As String, _
                                   ByVal vstrFchNacimiento As String, _
                                   ByVal vstrDesEmail As String, _
                                   ByVal vstrFlgCommit As String, _
                                   ByVal vstrTipoTransacc As String, Optional CodigoZonal As String = "")
                                   
    On Error GoTo CtrlErr
    gvarValores = Array(vstrCia, vstrCodConvenio, _
                        vstrCodCliente, vstrDni, _
                        vstrApePaterno, vstrApeMaterno, _
                        vstrNombres, vstrCargo, _
                        vdblImpLinCred, vstrFlgCamTemLinCred, _
                        vdblImpLinTmp, vstrFchIniNvaLinCred, _
                        vstrFchFinNvaLinCred, vstrDesObs, _
                        vstrCodEmpleado, vstrFlgActivo, _
                        vstrCodUsuario, vstrCodEstadoCivil, _
                        vstrFchNacimiento, vstrDesEmail, _
                        vstrFlgCommit, vstrTipoTransacc, CodigoZonal)
                        
    gvarIO = Array(entrada, entrada, _
                   Entrada_Salida, entrada, _
                   entrada, entrada, _
                   entrada, entrada, _
                   entrada, entrada, _
                   entrada, entrada, _
                   entrada, entrada, _
                   entrada, entrada, _
                   entrada, entrada, _
                   entrada, entrada, _
                   entrada, entrada, entrada)

    GrabaBenefiarioCnv = gclsOracle.SP("BTLPROD.PKG_BENEFICIARIO.SP_GRABA", gvarValores, gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsConvenio.GrabaBenefiarioCnv", Err.Description
End Function


Public Function LineaCreditoBase(ByVal CodigoConvenio As String) As String
    On Error GoTo CtrlErr
        LineaCreditoBase = gclsOracle.FN_Valor("btlprod.pkg_convenio.fn_dev_linea_base", CodigoConvenio)
    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsConvenio", Err.Description
End Function



Public Function ListaAddBeneficiario() As oraDynaset
    On Error GoTo CtrlErr
    
        Set ListaAddBeneficiario = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_LISTA_ADD_BENEF", 0)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsConvenio.ListaAddBeneficiario", Err.Description
End Function


Public Function TieneListaDatoAdicional(ByVal CodigoDatoAdicional As String) As String
    On Error GoTo CtrlErr
        TieneListaDatoAdicional = gclsOracle.FN_Valor("btlprod.pkg_convenio.FN_DEV_lISTA_DA", CodigoDatoAdicional)
    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsConvenio.TieneListaDatoAdicional", Err.Description
End Function
Public Function ListaValoresDatosAdicionales(ByVal CodigoDatoAdicional As String) As oraDynaset
    On Error GoTo CtrlErr
    
        Set ListaValoresDatosAdicionales = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_lISTA_DATOS_DA", 0, CodigoDatoAdicional)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsConvenio.ListaValoresDatosAdicionales", Err.Description
End Function


Public Function ListaBarraConvenio(ByVal strBusqueda As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaBarraConvenio = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO_DSCTO.fn_dev_barraconvenio", 0, strBusqueda)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsConvenio.ListaBarraConvenio", Err.Description
End Function

Public Function ListaConvenioCompetencia(ByVal CodigoConvenio As String) As Integer
    On Error GoTo CtrlErr
        ListaConvenioCompetencia = gclsOracle.FN_Valor("BTLPROD.PKG_CONVENIO.FN_LISTA_CONVENIO_COMPETENCIA", CodigoConvenio)
    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsConvenio", Err.Description
End Function
