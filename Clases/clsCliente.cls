VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ListaTipoContacto() As oraDynaset
      Set ListaTipoContacto = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_TIPO_CONT", 0)
End Function
Public Function ListaDireccion(Codigo As String, Optional CodDireccionCli As String = "") As oraDynaset
      Set ListaDireccion = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_DIRECCION", 0, Codigo, CodDireccionCli)
End Function

Public Function DesSubFijo(ByVal vstrCodSuFijo As String) As oraDynaset
      Set DesSubFijo = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_DES_SUBFIJO", 0, vstrCodSuFijo)
End Function

Public Function ListaxRuc(ByVal vstrRuc_Cliente As String) As oraDynaset
      Set ListaxRuc = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_CLIENTE_CAMPO", 0, "CMR.MAE_CLIENTE", "NUM_DOCUMENTO_ID", vstrRuc_Cliente)
End Function

Public Function ListaxDNI(ByVal vstrDNI_Cliente As String) As oraDynaset
      Set ListaxDNI = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_CLIENTE_CAMPO", 0, "CMR.MAE_CLIENTE", "COD_DOCUMENTO_IDENTIDAD", vstrDNI_Cliente)
End Function

Public Function ListaCampo(ByVal Campo As String, ByVal Dato As String) As oraDynaset
    Set ListaCampo = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_CLIENTE_CAMPO", 0, "CMR.MAE_CLIENTE", Campo, Dato)
End Function
Public Function Lista(Optional ByVal CodigoCliente As String) As oraDynaset
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA", 0, CodigoCliente)
End Function
'I.ECASTILLO 17.12.2020
Public Function Lista_v2(Optional ByVal CodigoCliente As String) As oraDynaset
    Set Lista_v2 = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_V2", 0, CodigoCliente)
End Function
'F.ECASTILLO 17.12.2020
Public Function ListaTipoDoc_Indet() As oraDynaset
    On Error GoTo CntrlError
    Set ListaTipoDoc_Indet = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_TIPO_DOC_IDENT", 0)
    
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsCliente", Err.Description
End Function

Public Function ListaTipoDocIndentidad(Optional strCodigo As String = "") As oraDynaset
    On Error GoTo handle
    Set ListaTipoDocIndentidad = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_TIPO_DOCUMENTO", 0, strCodigo)
    
    Exit Function
handle:
    Err.Raise Err.Number, "clsCliente", Err.Description
End Function

Public Function ListaClientesGen(ByVal vstrCliente As String, ByVal vstrFlg As String) As oraDynaset
    On Error GoTo CntrlError
    Set ListaClientesGen = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_GEN", 0, vstrCliente, vstrFlg)
    
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsCliente", Err.Description
End Function

Public Function ListaTipo() As oraDynaset
    On Error GoTo CntrlError
        Set ListaTipo = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_TIPO", 0)
        
    Exit Function
CntrlError:
    Err.Raise Err.Number, "clsCliente", Err.Description
End Function

'-- Verificacion de cliente nuevo --'
Public Function ListaClixVerif(ByVal vstrCia As String, _
                               ByVal vstrCodLocal As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaClixVerif = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_CLI_X_VERIF", 0, vstrCia, vstrCodLocal)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsCliente", Err.Description
End Function

Public Function GrabarSoloLineas(ByVal CodCli As String, _
                                 ByVal Cod_TipoLInea As String, _
                                 ByVal Cod_Secuencia As String, _
                                 ByVal Cod_DesValor As String) As String
On Error GoTo CtrlErr
    Dim varValores As Variant
    Dim varIO As Variant
    
    varValores = Array(CodCli, _
                       Cod_TipoLInea, Cod_Secuencia, Cod_DesValor)

    varIO = Array(entrada_salida, entrada, entrada, entrada)

    GrabarSoloLineas = gclsOracle.SP("BTLPROD.PKG_CLIENTE.SP_GRABA_TELEF", _
                                                           varValores, varIO)
Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsCliente.GrabarSoloLineas", Err.Description
End Function

Public Function GrabaVerificacion(ByVal vstrCodCli As String, _
                                  ByVal vstrVerif As String, _
                                  ByVal Cia As String, _
                                  ByVal CodigoLocal As String, _
                                  ByVal CodUsuario As String, _
                                  ByVal NumProforma As String _
                                  ) As String
    On Error GoTo CtrlErr
    Dim varValores As Variant
    Dim varIO As Variant
    
    varValores = Array(vstrCodCli, _
                       vstrVerif, Cia, CodigoLocal, CodUsuario, NumProforma)
    
    varIO = Array(entrada, _
                  entrada, _
                  entrada, _
                  entrada, _
                  entrada, _
                  entrada)
    
    GrabaVerificacion = gclsOracle.SP("BTLPROD.PKG_CLIENTE.SP_GRABA_VERIF_CLIENTE", _
                                                           varValores, varIO)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsCliente", Err.Description
End Function

Public Function ListaTelefono(ByVal Campo As String, Optional Cia As String) As oraDynaset
    Set ListaTelefono = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_TELEFONO", 0, Campo, Cia)
End Function

Public Function TelefonoCliente(ByVal CodigoCliente As String, ByVal Campo As String) As oraDynaset
    Set TelefonoCliente = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_TELEFONO_CLIENTE", 0, CodigoCliente, Campo)
End Function

Public Function Ultima_Sec_Telefono(ByVal CodigoCliente As String) As oraDynaset
    Set Ultima_Sec_Telefono = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_ULTIMA_SEC_TELEFONO", 0, CodigoCliente)
End Function

Public Function ListaTarjetas(ByVal CodigoCliente As String) As oraDynaset
      Set ListaTarjetas = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_TARJETAS", 0, CodigoCliente)
End Function
Public Function ListaRazonSocial(ByVal CodigoCliente As String) As oraDynaset
      Set ListaRazonSocial = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_RAZONSO", 0, CodigoCliente)
End Function

Public Function ListaCliente(ByVal Nombres As String, _
                             ByVal TipoCliente As String, _
                             ByVal TipoDoc As String, _
                             ByVal NumDoc As String, _
                             ByVal Activos As String, _
                             Optional Fono As String) As oraDynaset
                             
      Set ListaCliente = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_CLIENTE_NEW", 0, _
                                                                                    Nombres, _
                                                                                    TipoCliente, _
                                                                                    TipoDoc, _
                                                                                    NumDoc, _
                                                                                    Activos, _
                                                                                    Fono)
End Function


Public Property Get TipoClienteNatural() As Variant
    TipoClienteNatural = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIPO_CLIENTE_NATURAL")
End Property


Public Property Get TipoClienteJuridico() As Variant
    TipoClienteJuridico = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIPO_CLIENTE_JURIDICO")
End Property

Public Function SuFijoDirecc() As oraDynaset
    On Error GoTo CtrlErr
    Set SuFijoDirecc = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LST_SUFIJOS_DIREC", 0)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsCliente.SuFijoDirecc", Err.Description
End Function

Public Function ListaUrbanizacion(ByVal pCodUbigeo As String, Optional ByVal pCodUrbanizacion As String = "") As oraDynaset
    On Error GoTo CtrlErr
    Set ListaUrbanizacion = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_URBANIZACION", 0, pCodUbigeo, pCodUrbanizacion)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsCliente.ListaUrbanizacion", Err.Description
End Function


Public Function GrabaAuxDireccionCliente(ByVal pCodCliente As String, _
                                  ByVal pCodTipDireccion As String, _
                                  ByVal pDesDireccion As String, _
                                  ByVal pDesReferencia As String, _
                                  ByVal pCodUbigeo As String, _
                                  ByVal pCodUrbanizacion As String, _
                                  ByVal pCodSufijoDireccion As String, _
                                  ByVal pFlgPrincipal As String) As String
    On Error GoTo CtrlErr
    Dim varValores As Variant
    Dim varIO As Variant
    
    
    
    varValores = Array(pCodCliente, _
                                  pCodTipDireccion, _
                                  pDesDireccion, _
                                  pDesReferencia, _
                                  pCodUbigeo, _
                                  pCodUrbanizacion, _
                                  pCodSufijoDireccion, _
                                  pFlgPrincipal)
    
    varIO = Array(entrada, _
                entrada, _
                entrada, _
                entrada, _
                  entrada, _
                  entrada, _
                  entrada, _
                  entrada)
    
    
    
    
    GrabaAuxDireccionCliente = gclsOracle.SP("BTLPROD.PKG_CLIENTE.SP_GRABA_DIRECC_AUX", _
                                                           varValores, varIO)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsCliente.GrabaAuxDireccionCliente", Err.Description
End Function


Public Function ListaTipoDireccionCEN(Optional ByVal pCodTipoDireccion As String = "") As oraDynaset
    On Error GoTo CtrlErr
    Set ListaTipoDireccionCEN = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_TIPO_DIRECCION_CEN", 0, pCodTipoDireccion)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsCliente.ListaTipoDireccionCEN", Err.Description
End Function



Public Function fnValidaRUC(ByVal strNumRuc As String) As Integer
    On Error GoTo CntrlErr
    fnValidaRUC = gclsOracle.FN_Valor("BTLCERO.FN_BTL_VALIDA_RUC", strNumRuc)
    
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsCliente.fnValidaRUC", Err.Description

End Function



Public Function ListaBuscador(ByVal strDato As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaBuscador = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_BUSCADOR", 0, strDato)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsCliente.ListaBuscador", Err.Description


End Function



Public Function AnularDireccion(ByVal CodigoCliente As String, _
                                ByVal CodigoDireccion As String, _
                                ByVal TipoDireccion As String) As String
On Error GoTo CtrlErr
    Dim varValores As Variant
    Dim varIO As Variant
    
    varValores = Array(CodigoCliente, _
                       CodigoDireccion, TipoDireccion)

    varIO = Array(entrada_salida, entrada, entrada)

    AnularDireccion = gclsOracle.SP("BTLPROD.PKG_CLIENTE.SP_ANULA_DIRECCION", _
                                                           varValores, varIO)
Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsCliente.AnularDireccion", Err.Description
End Function


Public Function ListaHistoricoRazonSocial(CodigoCliente As String) As oraDynaset
      Set ListaHistoricoRazonSocial = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_LISTA_HIST_RAZON", 0, CodigoCliente)
End Function


Public Function Nombre(ByVal strCodCliente As String) As String
    On Error GoTo CntrlErr
    
    Nombre = gclsOracle.FN_Valor("BTLPROD.PKG_CLIENTE.FN_NOMBRE", strCodCliente)
    
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsCliente.Nombre", Err.Description

End Function
'I.ECASTILLO 20.07.2021
Public Function getUbigeoDesc(ByVal strDato As String) As oraDynaset
    On Error GoTo CtrlErr
    Set getUbigeoDesc = gclsOracle.FN_Cursor("BTLPROD.PKG_CLIENTE.FN_GET_DES_UBIGEO", 0, strDato)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsCliente.getUbigeoDesc", Err.Description
End Function
'F.ECASTILLO 20.07.2021
