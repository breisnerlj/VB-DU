VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private lsCodClaseCOM As Variant
Private lsCodSubClaseCom As Variant
Private lsCodFamiliaCom As Variant
Private lsCodCategoriaCom As Variant
Private lsCodProdTrans As Variant
Private imagen As Image


Dim objProforma As New clsProforma
Dim pTipoDoc$
Dim pTipoPN$, pTipoPJ$
Dim pDni$, pNombres$, pApellidos$, pDireccionPN$
Dim pRuc$, pRazonSocial$, pDireccionPJ$

Dim objMensaje As New clsMensaje
Dim objCliente As New clsCliente





'''''Public Function CargaDatosPN(ByVal vstrTipoDoc As String, '''''                               ByVal vstrTipo As String, '''''                               ByVal vstrDni As String, '''''                               ByVal vstrNombres As String, '''''                               ByVal vstrApellidos As String, '''''                               ByVal vstrDireccion As String) As String
'''''
'''''    On Error GoTo CnrlErr
'''''
'''''            pTipoDoc = vstrTipoDoc
'''''            pTipoPN = vstrTipo
'''''            pDni = vstrDni
'''''            pNombres = vstrNombres
'''''            pApellidos = vstrApellidos
'''''            pDireccionPN = vstrDireccion
'''''
'''''    Exit Function
'''''CnrlErr:
'''''    Err.Raise Err.Number, "clsDocumento", Err.Description
'''''
'''''End Function
'''''
'''''Public Function CargaDatosPJ(ByVal vstrTipoDoc As String, '''''                               ByVal vstrTipo As String, '''''                               ByVal vstrruc As String, '''''                               ByVal vstrRazon As String, '''''                               ByVal vstrDireccion As String) As String
'''''
'''''    On Error GoTo CnrlErr
'''''
'''''            pTipoDoc = vstrTipoDoc
'''''            pTipoPN = vstrTipo
'''''            pRuc = vstrruc
'''''            pRazonSocial = vstrRazon
'''''            pDireccionPJ = vstrDireccion
'''''
'''''    Exit Function
'''''CnrlErr:
'''''    Err.Raise Err.Number, "clsDocumento", Err.Description
'''''
'''''End Function

''xxxxxxxxxxxxxxxxxxxxxxxx'
''---- Trae Valores -----'
''xxxxxxxxxxxxxxxxxxxxxxxx'
'Public Property Let TipoDoc(pTipoDoc)
'    TipoDoc = pTipoDoc
'End Property
'
'Public Property Let Tipo(vstrTipo)
'    Tipo = vstrTipo
'End Property
'
'Public Property Let Ruc(vstrruc)
'    Ruc = vstrruc
'End Property
'
'Public Property Let RazonSocial(vstrRazon)
'    RazonSocial = vstrRazon
'End Property
'
'Public Property Let DireccionPJ(vstrDireccion)
'    DireccionPJ = vstrDireccion
'End Property
'
'Public Property Let Dni(vstrDni)
'    Dni = vstrDni
'End Property
'
'Public Property Let Nombres(vstrNombres)
'    Nombres = vstrNombres
'End Property
'
'Public Property Let Apellidos(vstrApellidos)
'    Apellidos = vstrApellidos
'End Property
'
'Public Property Let DireccionPN(vstrDireccion)
'    DireccionPN = vstrApellidos
'End Property

''xxxxxxxxxxxxxxxxxxxxxxxx'
''---- Lleva Valores -----'
''xxxxxxxxxxxxxxxxxxxxxxxx'
Public Property Get TipoDoc()
    TipoDoc = pTipoDoc
End Property

Public Property Get Tipo()
    Tipo = pTipoPN
End Property

Public Property Get Ruc()
    Ruc = pRuc
End Property

Public Property Get RazonSocial()
    RazonSocial = pRazonSocial
End Property

Public Property Get DireccionPJ()
    DireccionPJ = pDireccionPJ
End Property

Public Property Get DNI()
    DNI = pDni
End Property

Public Property Get Nombres()
    Nombres = pNombres
End Property

Public Property Get Apellidos()
    Apellidos = pApellidos
End Property

Public Property Get DireccionPN()
    DireccionPN = pDireccionPN
End Property
'''''''''''''''''12/11/2009  --Arturo Escate: Comento esta funcion porque deberia de llamarse a clsVenta.GrabaDoc
'''''''''''''''''**** Graba Documento Creado Por -- Crueda ****'
''''''''''''''''Public Function Graba(ByVal vstrEmpresa As String, ByVal vstrTipoDoc As String, _
''''''''''''''''                      ByVal vstrDocumento As String, ByVal vstrLocal As String, _
''''''''''''''''                      ByVal vstrModalidad As String, ByVal vstrConvenio As String, _
''''''''''''''''                      ByVal vstrCliente As String, ByVal vstrTipoVta As String, _
''''''''''''''''                      ByVal vstrObs As String, ByVal vstrImpTotal As String, _
''''''''''''''''                      ByVal vstrTipoCambio As String, ByVal vstrFlgExo As String, _
''''''''''''''''                      ByVal vstrImpValor As String, ByVal vstrImpIgv As String, _
''''''''''''''''                      ByVal vstrNumReceta As String, ByVal vstrPctBenef As String, _
''''''''''''''''                      ByVal vstrMaquina As String, ByVal vstrlocalRef As String, _
''''''''''''''''                      ByVal vstrTipoDocRef As String, ByVal vstrNumDocRef As String, _
''''''''''''''''                      ByVal vstrConcepto As String, ByVal vstrSubConcepto As String, _
''''''''''''''''                      ByVal vstrUSUARIO As String, ByVal vstrCadProducto As String, _
''''''''''''''''                      ByVal vstrCadPromocion As String, ByVal vstrCadCtdUnd As String, _
''''''''''''''''                      ByVal vstrCadCtdFrac As String, ByVal vstrCadTotFrac As String, _
''''''''''''''''                      ByVal vstrImpPreLocal As String, ByVal vstrCadImpPreVta As String, _
''''''''''''''''                      ByVal vstrImpPreKairos As String, ByVal vstrCadImpIgv As String, _
''''''''''''''''                      ByVal vstrCadImpSTotal As String, ByVal vstrCadImpSubTot As String, _
''''''''''''''''                      ByVal vstrCadImpComision As String, ByVal vstrCadPctDscto As String, _
''''''''''''''''                      ByVal vstrCadFlgPromo As String) As String
''''''''''''''''Dim gvarValores As Variant
''''''''''''''''Dim gvarIO  As Variant
''''''''''''''''                  On Error GoTo CnrlErr
''''''''''''''''                      gvarValores = Array(vstrEmpresa, vstrTipoDoc, _
''''''''''''''''                                         "", vstrLocal, _
''''''''''''''''                                         vstrModalidad, vstrConvenio, _
''''''''''''''''                                         vstrCliente, vstrTipoVta, _
''''''''''''''''                                         vstrObs, vstrImpTotal, _
''''''''''''''''                                         vstrTipoCambio, vstrFlgExo, _
''''''''''''''''                                         vstrImpValor, vstrImpIgv, _
''''''''''''''''                                         vstrNumReceta, vstrPctBenef, _
''''''''''''''''                                         vstrMaquina, vstrlocalRef, _
''''''''''''''''                                         vstrTipoDocRef, vstrNumDocRef, _
''''''''''''''''                                         vstrConcepto, vstrSubConcepto, _
''''''''''''''''                                         vstrUSUARIO, vstrCadProducto, _
''''''''''''''''                                         vstrCadPromocion, vstrCadCtdUnd, _
''''''''''''''''                                         vstrCadCtdFrac, vstrCadTotFrac, _
''''''''''''''''                                         vstrImpPreLocal, vstrCadImpPreVta, _
''''''''''''''''                                         vstrImpPreKairos, vstrCadImpIgv, _
''''''''''''''''                                         vstrCadImpSTotal, vstrCadImpSubTot, _
''''''''''''''''                                         vstrCadImpComision, vstrCadPctDscto, _
''''''''''''''''                                         vstrCadFlgPromo)
''''''''''''''''
''''''''''''''''                        gvarIO = Array(entrada, entrada, _
''''''''''''''''                                       Salida, entrada, _
''''''''''''''''                                       entrada, entrada, _
''''''''''''''''                                       entrada, entrada, _
''''''''''''''''                                       entrada, entrada, _
''''''''''''''''                                       entrada, entrada, _
''''''''''''''''                                       entrada, entrada, _
''''''''''''''''                                       entrada, entrada, _
''''''''''''''''                                       entrada, entrada, _
''''''''''''''''                                       entrada, entrada, _
''''''''''''''''                                       entrada, entrada, _
''''''''''''''''                                       entrada, entrada, _
''''''''''''''''                                       entrada, entrada, _
''''''''''''''''                                       entrada, entrada, _
''''''''''''''''                                       entrada, entrada, _
''''''''''''''''                                       entrada, entrada, _
''''''''''''''''                                       entrada, entrada, _
''''''''''''''''                                       entrada, entrada, _
''''''''''''''''                                       entrada)
''''''''''''''''
''''''''''''''''                        Graba = gclsOracle.SP("BTLPROD.PKG_DOCUMENTO.SP_GRABA", gvarValores, gvarIO)
''''''''''''''''
''''''''''''''''        Exit Function
''''''''''''''''CnrlErr:
''''''''''''''''        Err.Raise Err.Number, "clsDocumento", Err.Description
''''''''''''''''End Function

Public Function ListaTipo(ByVal strCodCia As String, ByVal strCodLocal As String) As oraDynaset
On Error GoTo CtrlErr
    Set ListaTipo = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_LISTA_TIPO", 0, strCodCia, strCodLocal)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaTipo", Err.Description
End Function

Public Function ListaTipoDocVta(ByVal strCodCia As String, ByVal strCodLocal As String, ByVal strCodMaquina As String) As oraDynaset
On Error GoTo CtrlErr
    Set ListaTipoDocVta = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_LISTA_TIPODOC_VTA ", 0, strCodCia, strCodLocal, strCodMaquina)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaTipoDocVta", Err.Description
End Function

'---- Nota de credito ---'
Public Function ListaNotaCredito(ByVal vstrNumNC As String) As oraDynaset

    On Error GoTo CntrlErr
        
        Set ListaNotaCredito = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_LISTA_NOTA_CREDITO", 0, vstrNumNC)
        
        Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsNotaCredito", Err.Description
End Function

Public Function Lista(ByVal strCodCia As String, _
                ByVal strCodLocal As String, _
                ByVal strTipoDoc As String, _
                ByVal strNumDocIni As String, _
                ByVal strNumDocFin As String, _
                ByVal strFchRegIni As Date, _
                ByVal strFchRegFin As Date) As oraDynaset
                
On Error GoTo CnrlErr
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_LISTA", 0, strCodCia, strCodLocal, strTipoDoc, strNumDocIni, strNumDocFin, strFchRegIni, strFchRegFin)
    Exit Function
CnrlErr:
    Err.Raise Err.Number, "clsDocumento", Err.Description
    
End Function

Public Function Reactiva(ByVal pCodCIA As String, _
                            ByVal pCodLocal As String, _
                            ByVal pCodTipoDoc As String, _
                            ByVal pNumDocumento As String, _
                            ByVal pCodEstado As String, _
                            ByVal pCodUsuario As String, _
                            ByRef pRetDocAnulado As String) As String

On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO As Variant
                        gvarValores = Array(pCodCIA, pCodLocal, pCodTipoDoc, pNumDocumento, pCodEstado, pCodUsuario, pRetDocAnulado)
                                            
                        gvarIO = Array(entrada, entrada, entrada, entrada, entrada, entrada, Salida)
                                       
                        Reactiva = gclsOracle.SP("BTLPROD.PKG_DOCUMENTO.SP_REACTIVA", gvarValores, gvarIO)
                        
                        pRetDocAnulado = gvarValores(6)
    
Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.Reactiva", Err.Description
End Function

Public Function ImpresionGuia(ByVal strCia As String, ByVal strNumGuia As String, Optional strCodLocal As String = "") As oraDynaset

    On Error GoTo CntrlErr
        
        Set ImpresionGuia = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_IMPRESION_GUIA", 0, strCia, strNumGuia, strCodLocal)
        
        Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsDocumento.ImpresionGuia", Err.Description
End Function

Public Function DocumentosReferencia(ByVal strCia As String, ByVal strCodLocal As String, _
                                     ByVal strTipoDocumento As String, ByVal strDocumento As String) As oraDynaset

    On Error GoTo CntrlErr
        
        Set DocumentosReferencia = gclsOracle.FN_Cursor("BTLPROD.PKG_CONVENIO.FN_DEV_REFERENCIAS", 0, strCia, strCodLocal, _
                                                        strTipoDocumento, strDocumento)

        Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsDocumento.DocumentosReferencia", Err.Description
End Function

Public Function Anula(ByVal pCodCIA As String, _
                            ByVal pCodLocal As String, _
                            ByVal pCodTipoDoc As String, _
                            ByVal pNumDocumento As String, _
                            ByVal pCodEstado As String, _
                            ByVal pCodUsuario As String, _
                            ByRef pRetDocAnulado As String) As String

On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO As Variant
Dim pRetNavsat As String

'objVenta.AnulaTransaccionNavsat("","","","","")

                        gvarValores = Array(pCodCIA, pCodLocal, pCodTipoDoc, pNumDocumento, pCodEstado, pCodUsuario, pRetDocAnulado, "", pRetNavsat)
                                            
                        gvarIO = Array(entrada, entrada, entrada, entrada, entrada, entrada, Salida, entrada, Salida)
                                       
                        Anula = gclsOracle.SP("BTLPROD.PKG_DOCUMENTO.SP_ANULA", gvarValores, gvarIO)
                        pRetDocAnulado = gvarValores(6)
    
'''''''                        If gvarValores(8) = "1" Then
'''''''                            'objVenta.AnulaTransaccionNavsat(
'''''''                        End If
Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.Anula", Err.Description
End Function

Public Function Correlativo(ByVal pCodCIA As String, _
                            ByVal pCodLocal As String, _
                            ByVal pCodTipoDoc As String, _
                            ByVal pNumDocDice As String, _
                            ByVal pNumDocDebe As String, _
                            ByVal pCodUsuario As String) As String

On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO As Variant
                        gvarValores = Array(pCodCIA, pCodLocal, pCodTipoDoc, pNumDocDice, pNumDocDebe, pCodUsuario)
                                            
                        gvarIO = Array(entrada, entrada, entrada, entrada, entrada, entrada)
                                       
                        Correlativo = gclsOracle.SP("BTLPROD.PKG_DOCUMENTO.SP_CORRELATIVO", gvarValores, gvarIO)
    
Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.Correlativo", Err.Description
End Function

Public Function AsignaUsuario(ByVal pCodCIA As String, _
                            ByVal pCodLocal As String, _
                            ByVal pCodTipoDoc As String, _
                            ByVal pNumDocumento As String, _
                            ByVal pCodUsuDep As String, _
                            ByVal pCodUsuario As String) As String

On Error GoTo CtrlErr

Dim gvarValores As Variant
Dim gvarIO  As Variant
                        gvarValores = Array(pCodCIA, pCodLocal, pCodTipoDoc, pNumDocumento, pCodUsuDep, pCodUsuario)
                                            
                        gvarIO = Array(entrada, entrada, entrada, entrada, entrada, entrada)
                                       
                        AsignaUsuario = gclsOracle.SP("BTLPROD.PKG_DOCUMENTO.SP_ASIGNA_USUARIO", gvarValores, gvarIO)

Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.Correlativo", Err.Description
End Function

'---- Para la nota de credito ----'
Public Function ListaCabecera(ByVal vstrCia As String, _
                              ByVal vstrNumDoc As String, _
                              ByVal vstrTipoDoc As String, _
                              ByVal vstrCodBtl As String, _
                              Optional ByVal flgOtroLocal As String = "1") As oraDynaset
    On Error GoTo CntrlErr
    
    Set ListaCabecera = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_CAB_DOCUMENTO", 0, vstrCia, vstrNumDoc, vstrTipoDoc, vstrCodBtl, flgOtroLocal)
    
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsDocumento", Err.Description
End Function

Public Function ListaDetalle(ByVal vstrCia As String, _
                              ByVal vstrNumDoc As String, _
                              ByVal vstrTipoDoc As String, _
                              Optional vstrCodLocal As String) As oraDynaset

    On Error GoTo CntrlErr
    
    Set ListaDetalle = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_DET_DOCUMENTO", 0, vstrCia, vstrNumDoc, vstrTipoDoc)
    
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsDocumento", Err.Description
End Function

'------- @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ ---------'
'--===== Funciones que hacen lectura a documentos de BTLCADENA =======--'
'------- @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ ---------'

Public Function Lista_CDocumentos(ByVal vstrNumDoc As String, _
                                  ByVal vstrcstrTipDoc As String, _
                                  ByVal vstrCodBtl As String) As oraDynaset
    On Error GoTo CtrlErr
        Set Lista_CDocumentos = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_C_DOCUMENTOS", 0, vstrNumDoc, vstrcstrTipDoc, vstrCodBtl)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.Lista_CDocumentos", Err.Description
End Function

Public Function Lista_DDocumentos(ByVal vstrNumDoc As String, _
                                  ByVal vstrcstrTipDoc As String) As oraDynaset
    On Error GoTo CtrlErr
        Set Lista_DDocumentos = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_D_DOCUMENTOS", 0, vstrNumDoc, vstrcstrTipDoc)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.Lista_DDocumentos", Err.Description
End Function

Public Function ListaMotivoNC() As oraDynaset
    On Error GoTo CntrlErr
    
    Set ListaMotivoNC = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_LISTA_MOTIVO_NC", 0)
    
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsDocumento", Err.Description
End Function

Public Function GrabaGuia(ByVal Cia As String, _
                    ByVal CodigoLocal As String, _
                    ByVal TipoDocumento As String, _
                    ByVal Maquina As String, _
                    ByVal CodigoConvenio As String, _
                    ByVal CodigoBeneficiario As String, _
                    ByVal flgKardex As String, _
                    ByVal LocalDestino As String, _
                    ByVal usuario As String, _
                    ByVal Motivo As String, _
                    ByVal PctCopago As String, _
                    ByVal MotivoMovimiento As String, _
                    ByVal CadCodProducto As String, _
                    ByVal CadCtdProductoFrac As String, _
                    ByVal CadCtdProducto As String, _
                    ByVal CadImpProducto As String, _
                    Optional ByRef numDocumento As String = "", _
                    Optional ByRef ImpTotal As String = "", _
                    Optional strCommit As String, _
                    Optional ByRef arrValor As Variant) As String
                          
    Dim arrValores As Variant
    Dim arrDireccion As Variant
    On Error GoTo CntrlErr
    arrValores = Array(numDocumento, _
                      ImpTotal, _
                      Cia, _
                      CodigoLocal, _
                      TipoDocumento, _
                      Maquina, _
                      CodigoConvenio, _
                      CodigoBeneficiario, _
                      flgKardex, _
                      LocalDestino, _
                      usuario, _
                      Motivo, _
                      PctCopago, _
                      MotivoMovimiento, _
                      CadCodProducto, _
                      CadCtdProductoFrac, _
                      CadCtdProducto, _
                      CadImpProducto, _
                      "0", _
                      strCommit)
                      
    arrDireccion = Array(Salida, Salida, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada)
    GrabaGuia = gclsOracle.SP("BTLCERO.PKG_GUIA_CLIENTE.SP_GRABA_GUIA", arrValores, arrDireccion)
    arrValor = arrValores
'    Dim valor As String
'    valor = arrValores(0)
'     If GrabaGuia = "" Then
'          MsgBox "Se grabo la guia Nº " & valor & " ", vbInformation, ""
'          frmPedido.psub_BeginArry
'          mdiPrincipal.subNuevo
'          frm_VTA_Busqueda.grdProductos.Limpiar
'          frm_VTA_Busqueda.grdAlternativos.Limpiar
'          frm_VTA_Busqueda.grdComplementarios.Limpiar
'          frm_VTA_Busqueda.txtBuscar.selection
'       Else
'          MsgBox GrabaGuia, vbCritical
'     End If
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsDocumento", Err.Description
End Function

Public Function Impresion(ByVal Cia As String, _
                        ByVal TipoDocumento As String, _
                        ByVal numDocumento As String, _
                        Optional ByVal Donacion As String = "0") As oraDynaset
                        
On Error GoTo CtrlErr

    Set Impresion = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_IMPRESION", 0, Cia, TipoDocumento, numDocumento, Donacion)

    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.Impresion", Err.Description
End Function

Public Function Impresion_Ticket(ByVal Cia As String, _
                                 ByVal TipoDocumento As String, _
                                 ByVal numDocumento As String, _
                                 Optional ByVal Donacion As String = "0") As oraDynaset
                        
On Error GoTo CtrlErr

    Set Impresion_Ticket = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_IMPRESION_TICKET", 0, Cia, TipoDocumento, numDocumento, Donacion)

    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.Impresion_Ticket", Err.Description

End Function

Public Function ListaDocumentoMaquina(ByVal Cia As String, _
                        ByVal TipoDocumento As String, _
                        ByVal codMaquina As String) As oraDynaset
                        
On Error GoTo CtrlErr

    Set ListaDocumentoMaquina = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_LISTA_DOCUMENTO_MAQUINA", 0, Cia, TipoDocumento, codMaquina)

    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.ListaDocumentoMaquina", Err.Description
End Function

Public Function ListaTipoCampo(ByVal Cia As String, _
                        ByVal TipoDocumento As String, _
                        ByVal numDocumento As String, _
                        Optional CodLocal As String = "") As oraDynaset
                        
On Error GoTo CtrlErr

    Set ListaTipoCampo = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_LISTA_TIPO_CAMPO", 0, Cia, TipoDocumento, numDocumento, CodLocal)

    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.ListaTipoCampo", Err.Description
End Function

Public Function PctImpuesto(ByVal Cia As String, _
                        ByVal TipoDocumento As String, _
                        ByVal numDocumento As String) As Double

On Error GoTo CtrlErr

    PctImpuesto = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO.FN_PCT_IMPUESTO", Cia, TipoDocumento, numDocumento)

    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.PctImpuesto", Err.Description
End Function

Public Function UltDocEmitido(ByVal Cia As String, ByVal codMaquina As String) As oraDynaset

On Error GoTo CtrlErr

    Set UltDocEmitido = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_ULT_DOC_EMITIDO", 0, Cia, codMaquina)

    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.UltDocEmitido", Err.Description
End Function

Public Function GrabaUltDocEmitido(ByVal Cia As String, ByVal codMaquina As String, ByVal CodTipoDocumento As String) As String

On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO As Variant
                        gvarValores = Array(Cia, codMaquina, CodTipoDocumento)
                                            
                        gvarIO = Array(entrada, entrada, entrada)
                                       
                        GrabaUltDocEmitido = gclsOracle.SP("BTLPROD.PKG_DOCUMENTO.SP_GRABA_ULT_DOC_EMITIDO", gvarValores, gvarIO)

Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.GrabaUltDocEmitido", Err.Description
End Function

'-------------------------------------------------------------------------------------------------------------------------------'
'--- Impresion de la Nota de Credito ---'

Public Function fnCabNotaCred(ByVal vstrCia As String, _
                              ByVal vstrCodLocal As String, _
                              ByVal vstrNumDoc As String) As oraDynaset
    On Error GoTo CtrlErr
    Set fnCabNotaCred = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_CAB_NOTA_CRED", 0, vstrCia, vstrCodLocal, vstrNumDoc)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento_fnCabNotaCred", Err.Description
End Function

Public Function fnDetNotaCred(ByVal vstrNumDoc As String) As oraDynaset
    On Error GoTo CtrlErr
    Set fnDetNotaCred = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_DET_NOTA_CRED", 0, vstrNumDoc)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento_fnDetNotaCred", Err.Description
End Function

Public Function ConceptoCredito() As String
    On Error GoTo CtrlErr
    ConceptoCredito = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_CONCEPTO_NOTA_CRED")
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento_fnConceptoCredito", Err.Description
End Function


Private Sub Factura(ByVal TipoDocumento As String, ByVal numDocumento As String, Optional FlgAnulado As Boolean = False, Optional ByVal FlgDolares As Boolean = False)
    Dim rstTipoCampo As oraDynaset
    Dim rstImpresion As oraDynaset
    Dim rstDocumentoMaquina As oraDynaset
    Dim rstDonacion As oraDynaset
    Dim rstConvenio As oraDynaset
    Dim objDocumento As New clsDocumento
    Dim dblIGV As Double
    Dim strEntidadDonacion As String
    Dim i As Integer
    Dim j As Integer
    Dim arrMeses(12) As String
    Dim intNumeroLineas As Integer
    Dim dblDonacion As Double
    Dim strLetras As String
    Dim dblMtoBaseImp As Double
    Dim dblMtoExonerado As Double
    Dim dblMtoImpuesto As Double
    Dim dblMtoTotal As Double
    Dim strMensaje As String
    Dim intTamañoEspacio  As Integer
    Dim intFormato  As Integer
    Dim objImpresiones As New clsProforma
    Dim p As Printer
    Dim strGlosaConv As String
    Dim dblMtoTotalRef As Double
    '**Modified Martinnets
    Dim rsFormato As oraDynaset
    Dim objFormato As New clsImpresion
    Dim lsFormato As String
    '**************
    Dim strDesAuxCliNombre As String
    Dim strCodMaquina As String
    Dim strCodUsuarioDependiente As String
    Dim strCodConvenio As String
    Dim strlinea1 As String
    Dim strlinea2 As String
    Dim strlinea3 As String
    Dim strlinea4 As String
    Dim strlinea5 As String
    Dim strlinea6 As String
    Dim strlinea7 As String

    arrMeses(1) = "Enero"
    arrMeses(2) = "Febrero"
    arrMeses(3) = "Marzo"
    arrMeses(4) = "Abril"
    arrMeses(5) = "Mayo"
    arrMeses(6) = "Junio"
    arrMeses(7) = "Julio"
    arrMeses(8) = "Agosto"
    arrMeses(9) = "Setiembre"
    arrMeses(10) = "Octubre"
    arrMeses(11) = "Noviembre"
    arrMeses(12) = "Diciembre"
     
    On Error GoTo ErrorImpresora
   
    For Each p In Printers
        If UCase(p.PORT) = "LPT1:" Then
            Set Printer = p
            Exit For
        End If
    Next p

    'COMENTADO POR JLOPEZ PARA VERIFICAR LA IMPRESION EN XP

'    Set rsFormato = ListaFormato(objUsuario.CodigoEmpresa, TipoDocumento, NumDocumento)
'    lsFormato = rsFormato.Fields("COD_FORMATO")
'    objFormato.cargaParametros lsFormato

    Set rstConvenio = ListaDoc(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
    If Not rstConvenio.EOF Then strCodConvenio = "" & rstConvenio("COD_CONVENIO").Value
   
    Set rstImpresion = objDocumento.Impresion(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
    If rstImpresion("COD_MODALIDAD_VENTA") = "002" Then
        FacturaNew TipoDocumento, numDocumento
        Exit Sub
    End If
    If rstImpresion("COD_FORMATO").Value = "002" Then 'factura corta
   
        If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
    '        intFormato = objImpresiones.CrearFormato("Formato BTL - 3", 225, 152)
    '        Set objImpresiones = Nothing
    '        Printer.PaperSize = intFormato
            If Not objFormato.cargaParametros("002") Then
                Exit Sub
            End If
             ': Printer.Print
        Else
            Printer.Width = 225 * 56.7
            Printer.Height = 153 * 56.7 '152*56.7, ' 93.3 * 56.7 ' 280
        End If
       
        Printer.Print: Printer.Print
        Printer.FontName = "Draft 17cpi"
        Printer.Print "."
        Printer.FontName = "Draft 10cpi"
        Printer.Print Space(55) & "NIT " & objUsuario.Ruc
        Printer.FontName = "Draft 12cpi"
    '   Printer.Print
    '   Printer.Print
    
        dblIGV = objDocumento.PctImpuesto(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
        Set rstTipoCampo = objDocumento.ListaTipoCampo(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento, objUsuario.CodigoLocal)
    
        With rstTipoCampo
            i = 0
            If Not .EOF Then
                While Not .EOF
                    Printer.Print Space(28) & rstTipoCampo("DES_TIPO_CAMPO").Value & " " & rstTipoCampo("DES_VALOR").Value
                    i = i + 1
                    .MoveNext
                Wend
            End If
    
    '      i = i + 1
    
        End With
    '      If i < 4 Then
    '        For j = i To 4
    '           Printer.Print
    '        Next j
    '      End If
    
        For j = 1 To 4 - i
            Printer.Print
        Next j
    
        intNumeroLineas = 0
        dblDonacion = 0
        'Set rstDocumentoMaquina = objDocumento.ListaDocumentoMaquina(objUsuario.CodigoEmpresa, TipoDocumento, objUsuario.NombrePC)
        'If Not rstDocumentoMaquina.EOF Then intNumeroLineas = rstDocumentoMaquina("NUM_LINEAS_DOC").Value
        
        Dim objFormDoc As New clsFormato
        intNumeroLineas = objFormDoc.NumLineas("002")
        Set objFormDoc = Nothing
        
        dblMtoBaseImp = 0
        dblMtoExonerado = 0
        dblMtoImpuesto = 0
        dblMtoTotal = 0
         
        strCodMaquina = rstImpresion("COD_MAQUINA").Value
        strCodUsuarioDependiente = rstImpresion("COD_USUARIO_DEPENDIENTE").Value
        
        '*****************************************************************************'
        '***** Datos de proforma para imprimir cuando el la venta
        '***** se haya originado por delivery aqui saldra impreso la BOL/FAC con PRO
        '***** Hecho el 19/07/2007
    
        objVenta.NumPedidoPadre = "" & rstImpresion("NUM_PEDIDO_PADRE").Value
         
        '******************************************************************************'
    
        If Not rstImpresion.EOF Then
            'Printer.FontName = "Draft 12cpi"
            'Printer.Print
            'Printer.Print Space(18) & rstImpresion("COD_MAQUINA").Value & "   " & rstImpresion("COD_USUARIO_DEPENDIENTE").Value & "   " & rstImpresion("COD_USUARIO").Value
            ' AGREGADO POR DJARA PARA IMPRIMIR EN CASO DE CAMBIO DE DIRECCION
            If objUsuario.ImprimirDireccion = True Then
                Printer.Font.name = "Draft 20cpi"
                'Printer.Print Space(16) & "NUEVA DIR: " + left(objUsuario.direccion, AnchoFactura - 11)
                Printer.Print "" & Mensaje1(objUsuario.CodigoLocal, "FAC")
                Printer.Print "" & Mensaje2(objUsuario.CodigoLocal, "FAC")
                Printer.Font.name = "Draft 17cpi"
            Else
                Printer.Print
            End If
            ' DJARA - 03/09/2009
            
            Printer.Print Space(8) & left(CStr(Day(rstImpresion("FCH_REGISTRA").Value)) & " de " & arrMeses(Month(rstImpresion("FCH_REGISTRA").Value)) & " de " & CStr(Year(rstImpresion("FCH_REGISTRA").Value)) & Space(23), 23) & Mensaje3(objUsuario.CodigoLocal, "FAC")
            Printer.Print Space(8) & left(rstImpresion("DES_RAZON_SOCIAL").Value & Space(60), 60) & " " & rstImpresion("NUM_DOCUMENTO").Value
            Printer.Print Space(8) & rstImpresion("NUM_RUC_EMPRESA").Value & Space(10) & IIf(strCodConvenio = objUsuario.CodConvenioScotiaBank, "Crédito 30 días", "")
            Printer.Print Space(8) & rstImpresion("DES_AUX_CLI_DIRECC").Value
    
            'Printer.Print
            Printer.Print
            Printer.Print
            'Printer.FontName = "Draft 10cpi"
            
            Dim dblTipoCambio As Double
            strLetras = rstImpresion("LETRAS").Value
            dblMtoBaseImp = rstImpresion("MTO_BASE_IMP").Value
            dblMtoExonerado = rstImpresion("MTO_EXONERADO").Value
            dblMtoImpuesto = rstImpresion("MTO_IMPUESTO").Value
            dblMtoTotal = rstImpresion("MTO_TOTAL").Value
            ''esto lo hice para cambiar a dolares 07/06/2010, Arturo Escate
            If FlgDolares = True Then
                dblTipoCambio = fnDevTipoCambioDoc(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
                dblMtoTotal = rstImpresion("MTO_TOTAL").Value / dblTipoCambio
                strLetras = gclsOracle.FN_Valor("btlprod.pkg_rutina.letras", dblMtoTotal)
                dblMtoBaseImp = Round(rstImpresion("MTO_BASE_IMP").Value / dblTipoCambio, 2)
                dblMtoExonerado = Round(rstImpresion("MTO_EXONERADO").Value / dblTipoCambio, 2)
                dblMtoImpuesto = Round(rstImpresion("MTO_IMPUESTO").Value / dblTipoCambio, 2)
            End If
            
            strDesAuxCliNombre = "" & rstImpresion("DES_AUX_CLI_NOMBRE").Value
            Printer.Print
            Printer.FontName = "Draft 17cpi"
            While Not rstImpresion.EOF
                'Printer.Print right(Space(7) & Trim(rstImpresion("CANTIDAD").Value), 7) & Space(5) & left(rstImpresion("COD_PRODUCTO").Value & " " & rstImpresion("DES_PRODUCTO").Value & Space(40), 40) & _
                            Space(2) & right(Space(8) & Format(rstImpresion("PRC_UNIT_VTA").Value, "#####0.#0"), 8) & _
                           Space(6) & right(Space(8) & Format(rstImpresion("SUB_TOTAL").Value, "#####0.#0"), 8) & " " & right(Space(6) & Format(rstImpresion("PCT_DSCT_UNIT").Value, "#####0.#0"), 6) & _
                            " " & right(Space(8) & Format(rstImpresion("TOTAL").Value, "#####0.#0"), 8)
                If FlgDolares = True Then
                    Printer.Print right(Space(7) & Trim(rstImpresion("CANTIDAD").Value), 7) & Space(2) & left(rstImpresion("COD_PRODUCTO").Value & " " & rstImpresion("DES_PRODUCTO").Value & Space(40), 40) & _
                                  Space(2) & right(Space(20) & IIf(IsNull(rstImpresion("NUM_LOTE").Value), "", rstImpresion("NUM_LOTE").Value), 20) & _
                                  Space(2) & right(Space(10) & IIf(IsNull(rstImpresion("FCH_VENCIMIENTO").Value), "", rstImpresion("FCH_VENCIMIENTO").Value), 10) & Space(2) & right(Space(8) & Format((rstImpresion("PRC_UNIT_VTA").Value / dblTipoCambio), "#####0.#0"), 8) & _
                                  Space(2) & right(Space(8) & Format((rstImpresion("SUB_TOTAL").Value / dblTipoCambio), "#####0.#0"), 8) & Space(2) & right(Space(6) & Format((rstImpresion("PCT_DSCT_UNIT").Value / dblTipoCambio), "#####0.#0"), 6) & _
                                  Space(2) & right(Space(8) & Format((rstImpresion("TOTAL").Value / dblTipoCambio), "#####0.#0"), 8)
                Else
                    Printer.Print right(Space(7) & Trim(rstImpresion("CANTIDAD").Value), 7) & Space(2) & left(rstImpresion("COD_PRODUCTO").Value & " " & rstImpresion("DES_PRODUCTO").Value & Space(40), 40) & _
                                  Space(2) & right(Space(20) & IIf(IsNull(rstImpresion("NUM_LOTE").Value), "", rstImpresion("NUM_LOTE").Value), 20) & _
                                  Space(2) & right(Space(10) & IIf(IsNull(rstImpresion("FCH_VENCIMIENTO").Value), "", rstImpresion("FCH_VENCIMIENTO").Value), 10) & Space(2) & right(Space(8) & Format(rstImpresion("PRC_UNIT_VTA").Value, "#####0.#0"), 8) & _
                                  Space(2) & right(Space(8) & Format(rstImpresion("SUB_TOTAL").Value, "#####0.#0"), 8) & Space(2) & right(Space(6) & Format(rstImpresion("PCT_DSCT_UNIT").Value, "#####0.#0"), 6) & _
                                  Space(2) & right(Space(8) & Format(rstImpresion("TOTAL").Value, "#####0.#0"), 8)
                
                End If
                If dblIGV = 0 Then
                    If Not IsNull(rstImpresion("NUM_PART_ARANCEL").Value) Then
                        Printer.Print "        P.A. : " & rstImpresion("NUM_PART_ARANCEL").Value
                        i = i + 1
                    End If
                End If
    
                rstImpresion.MoveNext
                i = i + 1

            Wend
            
            Printer.FontName = "Draft 12cpi"
            
                        
            'Autor : Juan Arturo Eacate Espichan
            'Fecha : 11/10/2010
            Dim rsServicio As oraDynaset
            Dim objServicio As New clsServicio
            Dim u As Integer
            Set rsServicio = objServicio.ListaServiciosAsoc(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
            While Not rsServicio.EOF
                If objVenta.EsNavsat("" & rsServicio("COD_TIPO_SERVICIO")) = "1" And FlgAnulado = False Then
                    Printer.Print Space(10) & "PIN o Telefono : " & rsServicio("NUM_SUMINISTRO")
                    Printer.Print Space(10) & "Nº Aprobacion:" & rsServicio("NUM_RECIBO")
                    'Printer.Print " " & rsServicio("DES_LEYENDA")
                    i = i + 2
                    If Not IsNull(rsServicio("DES_LEYENDA")) Then
                
                        If InStr(rsServicio("DES_LEYENDA"), "1[") > 0 Then
                            strlinea1 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "1[") + 2), InStr(rsServicio("DES_LEYENDA"), "]1") - ((InStr(rsServicio("DES_LEYENDA"), "1[") + 2)))
                            If strlinea1 <> "" Then Printer.Print Space(10) & strlinea1: i = i + 1
                        End If
                    
                        If InStr(rsServicio("DES_LEYENDA"), "2[") > 0 Then
                            strlinea2 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "2[") + 2), InStr(rsServicio("DES_LEYENDA"), "]2") - ((InStr(rsServicio("DES_LEYENDA"), "2[") + 2)))
                            If strlinea2 <> "" Then Printer.Print Space(10) & strlinea2: i = i + 1
                        End If
                    
                        If InStr(rsServicio("DES_LEYENDA"), "3[") > 0 Then
                            strlinea3 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "3[") + 2), InStr(rsServicio("DES_LEYENDA"), "]3") - ((InStr(rsServicio("DES_LEYENDA"), "3[") + 2)))
                            If strlinea3 <> "" Then Printer.Print Space(10) & strlinea3: i = i + 1
                        End If
                    
                        If InStr(rsServicio("DES_LEYENDA"), "4[") > 0 Then
                            strlinea4 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "4[") + 2), InStr(rsServicio("DES_LEYENDA"), "]4") - ((InStr(rsServicio("DES_LEYENDA"), "4[") + 2)))
                            If strlinea4 <> "" Then Printer.Print Space(10) & strlinea4: i = i + 1
                        End If
                    
                        If InStr(rsServicio("DES_LEYENDA"), "5[") > 0 Then
                            strlinea5 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "5[") + 2), InStr(rsServicio("DES_LEYENDA"), "]5") - ((InStr(rsServicio("DES_LEYENDA"), "5[") + 2)))
                            If strlinea5 <> "" Then Printer.Print Space(10) & strlinea5: i = i + 1
                        End If
                    
                        If InStr(rsServicio("DES_LEYENDA"), "6[") > 0 Then
                            strlinea6 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "6[") + 2), InStr(rsServicio("DES_LEYENDA"), "]6") - ((InStr(rsServicio("DES_LEYENDA"), "6[") + 2)))
                            If strlinea6 <> "" Then Printer.Print Space(10) & strlinea6: i = i + 1
                        End If
                    
                        If InStr(rsServicio("DES_LEYENDA"), "7[") > 0 Then
                            strlinea7 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "7[") + 2), InStr(rsServicio("DES_LEYENDA"), "]7") - ((InStr(rsServicio("DES_LEYENDA"), "7[") + 2)))
                            If strlinea7 <> "" Then Printer.Print Space(10) & strlinea7: i = i + 1
                        End If

                    End If
            
                End If
                rsServicio.MoveNext
            Wend
            i = i + 1
            Set objServicio = Nothing
                        
            For j = 1 To intNumeroLineas - i
                Printer.Print
            Next j
    
            dblMtoTotalRef = 0
    
            dblMtoTotalRef = DevImpDocCnv(objUsuario.CodigoEmpresa, numDocumento, TipoDocumento, objUsuario.CodigoLocal)
    
            If Not rstConvenio.EOF Then
                strGlosaConv = DevTextoDocCnv(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento, , objUsuario.CodigoLocal)
                If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio Then
                    If rstConvenio("FLG_DOC_COPAGO").Value = "1" Then
                        Printer.Print "B E N E F I C I A R I O  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%)"
                    Else
                        Printer.Print "I N S T I T U C I O N  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%)"
                    End If
    
                    Printer.Print "#REF " & rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value & " " & rstConvenio("NUM_DOCUMENTO_REFER").Value & " (" & 100 - rstConvenio("PCT_BENEFICIARIO").Value & "%) - S/.: " & Format(IIf(dblMtoTotalRef < 0, 0, dblMtoTotalRef), "###,##0.00") & IIf(strDesAuxCliNombre = "", "", "   BENEFICIARIO: " & strDesAuxCliNombre)
    
                    Printer.Print DevNomRepartidor(objUsuario.CodigoEmpresa, _
                                          objUsuario.CodigoLocal, _
                                          TipoDocumento, _
                                          numDocumento)
                End If
            Else
                '' AQUI VA EL TEXTO DE LA PROMOCION
                Printer.Print
                Printer.Print
            End If
    
            Set rstDonacion = objDocumento.Impresion(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento, "1")
    
            If rstDonacion.EOF Then
                Printer.Print
            Else
                dblDonacion = 0
                While Not rstDonacion.EOF
                   dblDonacion = dblDonacion + rstDonacion("IMP_CON_REDONDEO").Value
                   strEntidadDonacion = strEntidadDonacion & left(rstDonacion("DES_HIJO").Value, 54) & "   " & Trim(objUsuario.SmbMoneda) & right("        " & Format(rstDonacion("IMP_CON_REDONDEO").Value, "####0.#0"), 8) & " "
                   rstDonacion.MoveNext
                Wend
    
                Printer.Print "     Usted está donando " & Trim(objUsuario.SmbMoneda) & Trim(Format(dblDonacion, "#####0.#0")); " a " & strEntidadDonacion
            End If
    
            '** Datos del Repartidor de tener **'
        If FlgAnulado = True Then
            Dim s As Variant
            s = Printer.FontSize
            Printer.FontSize = 17
            Printer.Print "DOCUMENTO ANULADO"
            Printer.FontSize = s
        End If
        '**
        Printer.Print "     SON : " & strLetras & " " & IIf(FlgDolares = True, "DOLARES AMERICANOS", (objUsuario.DesLgMoneda))
        Printer.Print
    '    Printer.Print
    '    Printer.Print
    '    Printer.Print
    
        Printer.Print Space(40) & "%  " & right(Space(11) & Format(dblIGV * 100, "##0.00"), 11)
        Printer.Print Space(8) & IIf(FlgDolares = True, "$", Trim(objUsuario.SmbMoneda)) & right(Space(11) & _
                      Format(IIf(dblMtoBaseImp > 0, dblMtoBaseImp, dblMtoExonerado), "#### ##0.#0"), 11) & Space(18) & _
                      IIf(FlgDolares = True, "$", Trim(objUsuario.SmbMoneda)) & right(Space(11) & Format(dblMtoImpuesto, "#### ##0.#0"), 11) & Space(25) & _
                      IIf(FlgDolares = True, "$", Trim(objUsuario.SmbMoneda)) & right(Space(11) & Format(dblMtoTotal, "#### ##0.#0"), 11)
        Printer.Print
        'Printer.Print
        strMensaje = rstImpresion("COD_MAQUINA").Value & "   " & rstImpresion("COD_USUARIO_DEPENDIENTE").Value & "   " & rstImpresion("COD_USUARIO").Value
        Printer.Print Space(AnchoFactura - (Len(strMensaje) + 1)) & strMensaje
        strMensaje = strCodMaquina + "   " + strCodUsuarioDependiente
        Printer.Print Space(AnchoFactura - (Len(strMensaje) + 1)) & strMensaje
        
        ''AQUI VA EL MENSAJE
        strMensaje = DevMensaje(objUsuario.CodigoEmpresa, numDocumento, TipoDocumento, objUsuario.CodigoLocal)
        Dim rsNewMess As oraDynaset
        Set rsNewMess = ListaMensajes
        While Not rsNewMess.EOF
            Printer.FontSize = 7
            Printer.Print "" & rsNewMess("DES_PARAMETRO").Value
            rsNewMess.MoveNext
        Wend

        'intTamañoEspacio = Int((AnchoFactura - Len(IIf(IsNull(strMensaje), "", strMensaje))) / 2)
        'Printer.Print Space(intTamañoEspacio) + IIf(IsNull(strMensaje), "", strMensaje)
        Printer.Print Space(AnchoFactura - (Len(strMensaje) + 1)) + IIf(IsNull(strMensaje), "", strMensaje)
        Printer.EndDoc
       
        Else
            MsgBox "No se encontraron datos en el documento " + Chr(13) + "F/ " & rstImpresion("NUM_DOCUMENTO").Value, vbYesNo + vbDefaultButton1
        End If
    Else
        ImprimeFacLarga objUsuario.CodigoEmpresa, numDocumento ' impresion de factura larga
    End If
    Set rstImpresion = Nothing
    
    ImprimeCuponMonedero TipoDocumento, numDocumento, FlgAnulado
    
    Exit Sub
ErrorImpresora:
    Err.Raise Err.Number, "clsImpresiones.Boleta", Err.Description
End Sub

Private Sub ImprimeFacLarga(ByVal vstrCia As String, _
                            ByVal vstrNum_Documento As String)

Dim odynCab As oraDynaset, odynDet As oraDynaset
Dim strImporteLetras As String

    Dim i%, intNumLineas%, intLinea%
    Dim strIgv As String

    Set odynCab = gclsOracle.FN_Cursor("BTLCERO.PKG_DOC.FN_CABECERA", 0, vstrCia, "FAC", vstrNum_Documento)
    If odynCab.RecordCount = 0 Then MsgBox "No se encontró Documento en la BD", vbCritical, "ERROR": Exit Sub
        
    Dim objFormato As New clsFormato
    intNumLineas = objFormato.NumLineas("001")
    Set objFormato = Nothing
        
    Set odynDet = gclsOracle.FN_Cursor("BTLCERO.PKG_DOC.FN_DETALLE", 0, vstrCia, "FAC", vstrNum_Documento, "1")
    
    If odynDet.RecordCount = 0 Then MsgBox "Documento sin detalle", vbCritical, "ERROR": Exit Sub
    
    strIgv = odynDet("PCT_IGV").Value
    
    Dim strSec_Caja$, strFch_Emision$, strNomCliente$, strDirEntrega$, strRuc_Cliente$, _
        strUsuPed$, strUsuDoc$, strTipoVenta$, strGuia$, strOC$
        
    Dim dblTotal As Double, dblIGV As Double, dblSubTotal As Double
    
    strFch_Emision = Format(odynCab("FCH_EMISION").Value, "DD/MM/YYYY")
    strNomCliente = "" & odynCab("NOMBRE").Value
    strDirEntrega = "" & odynCab("DIR_ENTREGA_PROD").Value
    strRuc_Cliente = "" & odynCab("RUC_CLIENTE").Value
    
    If Len("" & odynCab("COD_MAQUINA").Value) > 4 Then
        strSec_Caja = Mid("" & odynCab("COD_MAQUINA").Value, 4, Len("" & odynCab("RUC_CLIENTE").Value) - 3)
    End If
    strUsuPed = "" & odynCab("USU_PEDIDO").Value
    strUsuDoc = "" & odynCab("USU_EMISION").Value

    dblTotal = Val(odynCab("MTO_TOTAL").Value)
    dblIGV = Val(odynCab("MTO_IMPUESTO").Value)
    dblSubTotal = Val(odynCab("MTO_BASE_IMP").Value) + Val(odynCab("MTO_EXONERADO").Value)
    strTipoVenta = IIf(odynCab("FLG_TIPO_VENTA").Value = "0", "CONTADO", "CREDITO")
    
    If odynCab("COD_CLIENTE_IMP").Value = "0000000408" And Not IsNull(odynCab("COD_CONVENIO").Value) Then
        strTipoVenta = "CREDITO 45 DIAS"
    Else
        strTipoVenta = ""
    End If
        
    strGuia = "" & odynCab("NUM_GUIA").Value
    strOC = "" & odynCab("OBS_ORDEN_COMPRA").Value
    
    On Error GoTo ErrorImpresora
    
    Dim objImpresion As New clsImpresion
            
    If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
        
         If Not objImpresion.cargaParametros("001") Then
             Exit Sub
         End If
     Else
         Printer.Width = 225 * 56.7
         Printer.Height = 334 * 56.7
    End If
    
    Printer.Print
    Printer.Print
    
    Dim rstTipoCampo As oraDynaset
    Dim j As Integer
    Set rstTipoCampo = ListaTipoCampo(objUsuario.CodigoEmpresa, "FAC", vstrNum_Documento, objUsuario.CodigoLocal)
    
    With rstTipoCampo
      i = 0
      If Not .EOF Then
          While Not .EOF
              Printer.Print Space(28) & rstTipoCampo("DES_TIPO_CAMPO").Value & ": " & rstTipoCampo("DES_VALOR").Value
              i = i + 1
              .MoveNext
          Wend
      End If
    End With
    For j = 1 To 4 - i
      Printer.Print
    Next j
    
'    Printer.Print
'    Printer.Print
'    Printer.Print
'    Printer.Print
    
    Printer.Print Space(18) & strSec_Caja & "   " & strUsuPed & "   " & strUsuDoc
    Printer.Print '"DROGUERIA"
    Printer.Print
    Printer.Print
    Printer.Print Space(8) & Day(strFch_Emision) & " de " & Format(strFch_Emision, "mmmm") & " del " & Year(strFch_Emision)
    Printer.Print Space(48) & Format(strFch_Emision, "dd/mm/yyyy") & " " & left(vstrNum_Documento, 3) & "-" & right(vstrNum_Documento, 7)
    Printer.Print Space(54) & strTipoVenta
    Printer.Print
    Printer.FontName = "Draft 12cpi"
    Printer.Print Space(8) & left(strNomCliente & Space(50), 50) & " OC: " & Trim(strOC)
    Printer.FontName = "Draft 10cpi"
    Printer.Print
    Printer.Print Space(8) & strRuc_Cliente & Space(10) & Trim(strGuia)
    Printer.Print
    Printer.FontName = "Draft 17cpi"
    Printer.Print Space(8) & left(strDirEntrega & Space(80), 80)
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print

        intLinea = 0
        Dim strCodProdAnt As String
        strCodProdAnt = "*"
        While Not odynDet.EOF

            Printer.FontName = "Draft 17cpi"
    
            If strCodProdAnt <> odynDet("COD_PRODUCTO").Value Then
                Printer.Print left(Trim(odynDet("CANT").Value) & Space(6), 6) & Space(1) & _
                        left(odynDet("COD_PRODUCTO").Value & " " & odynDet("OBS_DES_PRODUCTO").Value & Space(55), 55) & Space(1) & _
                        left("" & odynDet("NUM_LOTE").Value & Space(10), 10) & Space(1) & _
                        left("" & odynDet("FCH_VEN").Value & Space(7), 7) & Space(1) & _
                        left("" & odynDet("CTD_LOTE").Value & Space(5), 5) & Space(1) & _
                        right(Space(6 + gintDec) & Format(odynDet("UNITARIO").Value, "# ##0." & String(gintDec, "0")), 6 + gintDec) & Space(18 - gintDec) & _
                        right(Space(10 + gintDec) & Format(odynDet("SUBTOTAL").Value, "# ### ##0." & String(gintDec, "0")), 10 + gintDec)
    
            Else
                Printer.Print left("" & Space(6), 6) & Space(1) & _
                        left("" & Space(55), 55) & Space(1) & _
                        left("" & odynDet("NUM_LOTE").Value & Space(10), 10) & Space(1) & _
                        left("" & odynDet("FCH_VEN").Value & Space(7), 7) & Space(1) & _
                        left("" & odynDet("CTD_LOTE").Value & Space(5), 5) & Space(1) & _
                        right(Space(6 + gintDec) & "", 6 + gintDec) & Space(18 - gintDec) & _
                        right(Space(10 + gintDec) & "", 10 + gintDec)
    
            End If
            strCodProdAnt = odynDet("COD_PRODUCTO").Value
        
            odynDet.MoveNext
            intLinea = intLinea + 1
        Wend
        For i = intLinea To intNumLineas
            Printer.Print
        Next i
        Printer.Print
        
        strImporteLetras = "" & odynCab("LETRAS").Value
        
        Printer.FontName = "Draft 17cpi"
        Printer.Print "" & gclsOracle.FN_Valor("BTLCERO.PKG_PEDIDO.FN_MSJ_PD", vstrCia, "FAC", vstrNum_Documento)
        Printer.Print "" & gclsOracle.FN_Valor("BTLCERO.PKG_PEDIDO.FN_MSJ_BOTIQUIN", vstrCia, "FAC", vstrNum_Documento)
        Printer.FontName = "Draft 12cpi"
        Printer.Print "     SON : " & strImporteLetras & " " & Trim(objUsuario.DesLgMoneda)
        
            Printer.Print
        
        Printer.Print Space$(78 - gintDec) & "S/." & right(Space(10 + gintDec) & Format(dblSubTotal, "# ### ##0." & String(gintDec, "0")), 10 + gintDec)
        Printer.Print
        Printer.Print Space$(62 - gintDec) & "Igv" & Space(2) & Val(strIgv * 100) & "%" & Space(8) & "S/." & right(Space(10 + gintDec) & Format(dblIGV, "# ### ##0." & String(gintDec, "0")), 10 + gintDec)
        Printer.Print
        Printer.Print Space$(76 - gintDecTot) & "S/." & right(Space(10 + gintDecTot) & Format(dblTotal, "# ### ##0." & String(gintDecTot, "0")), 10 + gintDecTot)
        Printer.Print
        
        Printer.EndDoc
    
    Exit Sub
ErrorImpresora:
    If MsgBox("Existe un problema con la Impresora" & Chr(13) & _
             Err.Description & Chr(13) & _
             "Desea esperar a ser resuelto?", vbQuestion + vbYesNo) = vbYes Then
        Resume
    Else
        Exit Sub
    End If
End Sub

Sub ImprimirDocumento(ByVal TipoDocumento As String, ByVal numDocumento As String, Optional ByVal Formato As String = "", Optional Modalidad As String, Optional FlgAnulado As Boolean = False, Optional flgProformaDLV As Boolean = False, Optional FlgDolares As Boolean = False)
    Dim NomImpresora As String
   
    On Error GoTo CtrlErr

    'Actualizacion 03/06/2015 para IMPRESORA LOCAL O RED DJARA
    NomImpresora = "" & gclsOracle.FN_Valor("BTLPROD.FN_NOM_IMPRESORA_TKB", objUsuario.NombrePC)
    If Not NomImpresora = "" Then
        Call Me.LocalORed(objUsuario.CodigoEmpresa, objUsuario.NombrePC, "TKB")
    End If
    'Fin de la actualizacion

    Select Case TipoDocumento
            Case "BOL"
                    Boleta TipoDocumento, numDocumento, FlgAnulado
            Case "FAC"
                    Factura TipoDocumento, numDocumento, FlgAnulado, FlgDolares
                    
'''Arturo escate 09/02/2010
'''                    If Modalidad = objUsuario.ModalidadVentaConvenio Then  'objUsuario.ModalidadVentaConvenio Then
'''                        FacturaNew TipoDocumento, NumDocumento
'''                    Else
'''                        Factura TipoDocumento, NumDocumento, FlgAnulado
'''                    End If
            Case "NCR"
                    NotaCred objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, TipoDocumento, numDocumento
            Case "GRL"
                If Formato = "001" Then
                    ImprimirGuiaTransferencia numDocumento
                Else
                    'Guia objUsuario.CodigoEmpresa, TipoDocumento, NumDocumento
                    GuiaNew objUsuario.CodigoEmpresa, TipoDocumento, numDocumento
                End If
            Case "REC"
                    ImprimirRecibo objUsuario.CodigoEmpresa, numDocumento, TipoDocumento, objUsuario.CodigoLocal, FlgAnulado
            Case "AJU"
                    Imprime_Ajuste_Cje numDocumento
            Case "PRO"
'                If objUsuario.TipDocDefault = "TKB" Then
'                    objProforma.ImprimirTicket_Proforma objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, NumDocumento
'                Else
                    objProforma.Imprimir objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, numDocumento
'                End If
                    'Proforma objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, NumDocumento, "0"
                    'objProforma.Imprimir objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, NumDocumento
            Case "TKF", "TKB"
'                   Ticket "BOL", "0010000007"
                    Ticket TipoDocumento, numDocumento, mdiPrincipal.pdblPagTkt, mdiPrincipal.pdblTPagTkt, FlgAnulado, flgProformaDLV
                    'TicketNew TipoDocumento, numDocumento, mdiPrincipal.pdblPagTkt, mdiPrincipal.pdblTPagTkt, FlgAnulado, flgProformaDLV
    End Select
        
   Exit Sub

CtrlErr:
    Err.Raise Err.Number, "clsImpresiones.ImprimirDocumento", Err.Description

End Sub

Private Sub Boleta(ByVal TipoDocumento As String, ByVal numDocumento As String, Optional FlgAnulado As Boolean = False)
Dim rstImpresion As oraDynaset
Dim rstDocumentoMaquina As oraDynaset
Dim rstDonacion As oraDynaset
Dim rstConvenio As oraDynaset
Dim objDocumento As New clsDocumento
Dim objConvenio As New clsConvenio
Dim strCodigo As String
Dim strProducto As String
Dim strCantidad As String
Dim dblSubTotal As Double
Dim dblImporte As Double
Dim dblTotal As Double
Dim dblDonacion As Double
Dim strLetras As String
Dim strCodMaquina As String
Dim strCodUsuarioDependiente As String
Dim strCodUsuario As String
Dim strGlosaConv As String
Dim strMensaje As String

'Dim xStrBenef As String

Dim i As Integer
Dim j As Integer
'''''''''''''''''''''''''''''''''''''''''''''''
Dim p As Printer
Dim intRespuesta As Integer
Dim intNumeroLineas As Integer
Dim intFormato  As Integer
Dim intTamañoEspacio  As Integer
Dim objImpresiones As New clsImpresiones
Dim dblMtoTotalRef As Double
'**Modified Martinnets
Dim rsFormato As oraDynaset
Dim objFormato As New clsImpresion
Dim lsFormato As String
Dim objCliente As clsCliente

Dim strlinea1 As String
Dim strlinea2 As String
Dim strlinea3 As String
Dim strlinea4 As String
Dim strlinea5 As String
Dim strlinea6 As String
Dim strlinea7 As String

'**************
'On Error GoTo ErrorImpresora

    For Each p In Printers
        If UCase(p.PORT) = "LPT1:" Then
            Set Printer = p
            Exit For
        End If
    Next p
    
    'COMENTADO POR JLOPEZ PARA PROBAR LA IMPRESION EN XP
    'Set rsFormato = ListaFormato(objUsuario.CodigoEmpresa, TipoDocumento, NumDocumento)
    'lsFormato = rsFormato.Fields("COD_FORMATO")
    'objFormato.cargaParametros lsFormato
    
    Set rstConvenio = ListaDoc(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
         
    If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
'        intFormato = objImpresiones.CrearFormato("Formato BTL - 2", 115, 94)
'        Set objImpresiones = Nothing
'        Printer.PaperSize = intFormato
        Dim objImpresion As New clsImpresion
        If Not objImpresion.cargaParametros("004") Then
            Exit Sub
        End If
        'Printer.Print: Printer.Print ': Printer.Print
    Else
        Printer.Width = 225 * 56.7
        Printer.Height = (280 / 3) * 56.7
        Printer.Print "."
    End If
    intNumeroLineas = 0
    dblDonacion = 0
    Set rstDocumentoMaquina = objDocumento.ListaDocumentoMaquina(objUsuario.CodigoEmpresa, TipoDocumento, objUsuario.NombrePC)
    If Not rstDocumentoMaquina.EOF Then intNumeroLineas = rstDocumentoMaquina("NUM_LINEAS_DOC").Value
    
    Set rstImpresion = objDocumento.Impresion(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
    dblTotal = 0
    '*****************************************************************************'
    '***** Datos de proforma para imprimir cuando el la venta
    '***** se haya originado por delivery aqui saldra impreso la BOL/FAC con PRO
    '***** Hecho el 19/07/2007
    
    objVenta.NumPedidoPadre = "" & rstImpresion("NUM_PEDIDO_PADRE").Value
    '******************************************************************************'
    If Not rstImpresion.EOF Then
        'Printer.Font.name = "Draft 15cpi"
        'Printer.Print Space(6) & "" & Mensaje2(objUsuario.CodigoLocal, "BOL")
        'Printer.Font.name = "Draft 17cpi"
        '''' COMENTADO POR JAHZEEL LOPEZ EL 19/11 PARA CAMBIAR DES_AUX_CLI_NOMBRE POR DES_RAZON_SOCIAL
        ''''Printer.Print Space(6) & "" & left(IIf(IsNull(rstImpresion("DES_AUX_CLI_NOMBRE").Value), "", rstImpresion("DES_AUX_CLI_NOMBRE").Value) & Space(25), 25) & _
                      Space(1) & left(rstImpresion("NUM_RUC_EMPRESA").Value & Space(15), 15) & rstImpresion("FCH_REGISTRA").Value & " " & rstImpresion("NUM_DOCUMENTO").Value
                      
        
        'Printer.Font.name = "Draft 20cpi"
        'Printer.Print "N. Dir.:" & gclsOracle.FN_Valor("CMR.PKG_LOCAL.FN_DIR", objUsuario.CodigoLocal)
        'Printer.Font.name = "Draft 17cpi"
        '"NUEVA DIRECCCION: " &
        
        
        Set objCliente = New clsCliente
        Printer.Print Space(6) & "" & left(IIf(IsNull(rstImpresion("DES_RAZON_SOCIAL").Value), IIf(IsNull(rstImpresion("DES_AUX_CLI_NOMBRE").Value), IIf("" & rstImpresion("COD_CLIENTE").Value = "", "", objCliente.Nombre("" & rstImpresion("COD_CLIENTE").Value)), "" & rstImpresion("DES_AUX_CLI_NOMBRE").Value), rstImpresion("DES_RAZON_SOCIAL").Value) & Space(25), 25) & _
                      Space(1) & left(rstImpresion("NUM_RUC_EMPRESA").Value & Space(15), 15) & rstImpresion("FCH_REGISTRA").Value & " " & rstImpresion("NUM_DOCUMENTO").Value
        Set objCliente = Nothing
        Printer.Print
        Printer.Print
        dblTotal = rstImpresion("MTO_TOTAL").Value
        strLetras = rstImpresion("LETRAS").Value
        strCodMaquina = rstImpresion("COD_MAQUINA").Value
        strCodUsuarioDependiente = rstImpresion("COD_USUARIO_DEPENDIENTE").Value
        strCodUsuario = rstImpresion("COD_USUARIO").Value

        Dim rstStrCodCom As String, strCodCom As String
        Do While Not rstImpresion.EOF
            'COMENTADO POR JLOPEZ PARA VERIFICAR LA IMPRESION DE LAS TRANSFERENCIAS GRATUITAS
            'rstStrCodCom = rstImpresion("COD_CLASE_COM").Value & rstImpresion("COD_SUBCLASE_COM").Value & rstImpresion("COD_FAM_COM").Value
            'strCodCom = lsCodClaseCOM & lsCodSubClaseCom & lsCodFamiliaCom
            strCodigo = rstImpresion("COD_PRODUCTO").Value
            'If lsCodProdTrans = strCodigo Or strCodCom = rstStrCodCom And InStr(1, lsCodCategoriaCom, rstImpresion("COD_CATEGORIA_COM").Value) > 0 Then
                '    strCodigo = ""
                '    strCantidad = ""
                '    dblSubTotal = 0
                '    dblImporte = 0
                '    strProducto = rstImpresion("DES_PRODUCTO").Value '& Chr(10) & "Transferencia Gratuita"
                '   Printer.Print Left(Trim(strCodigo) & "     ", 5) & " " & _
                                Left(strProducto & Space(40), 40) & _
                                Right("      " & "", 5) & " " & _
                                Right("        " & "", 8) & " " & _
                                Right("        " & "", 8)
                
            'Else
                strCantidad = "" & rstImpresion("CANTIDAD").Value
                dblSubTotal = "" & rstImpresion("SUB_TOTAL").Value
                dblImporte = "" & rstImpresion("TOTAL").Value
                strProducto = "" & rstImpresion("DES_PRODUCTO").Value
                Printer.Print left(Trim(strCodigo) & "     ", 5) & " " & _
                              left(strProducto & Space(40), 40) & _
                              right("      " & strCantidad, 5) & " " & _
                              right("        " & Format(dblSubTotal, "####0.#0"), 8) & " " & _
                              right("        " & Format(dblImporte, "####0.#0"), 8)

            'End If
            rstImpresion.MoveNext
            i = i + 1
        Loop
        'Autor : Juan Arturo Eacate Espichan
        'Fecah 16/04/2008
       
        Dim rsServicio As oraDynaset
        Dim objServicio As New clsServicio
        Dim u As Integer
        Set rsServicio = objServicio.ListaServiciosAsoc(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
        While Not rsServicio.EOF
            If objVenta.EsNavsat("" & rsServicio("COD_TIPO_SERVICIO")) = "1" And FlgAnulado = False Then
                Printer.Print "PIN o Telefono : " & rsServicio("NUM_SUMINISTRO")
                Printer.Print "Nº Aprobacion:" & rsServicio("NUM_RECIBO")
                'Printer.Print " " & rsServicio("DES_LEYENDA")
                i = i + 2
                If Not IsNull(rsServicio("DES_LEYENDA")) Then
                
                    If InStr(rsServicio("DES_LEYENDA"), "1[") > 0 Then
                        strlinea1 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "1[") + 2), InStr(rsServicio("DES_LEYENDA"), "]1") - ((InStr(rsServicio("DES_LEYENDA"), "1[") + 2)))
                        If strlinea1 <> "" Then Printer.Print strlinea1: i = i + 1
                    End If
                    
                    If InStr(rsServicio("DES_LEYENDA"), "2[") > 0 Then
                        strlinea2 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "2[") + 2), InStr(rsServicio("DES_LEYENDA"), "]2") - ((InStr(rsServicio("DES_LEYENDA"), "2[") + 2)))
                        If strlinea2 <> "" Then Printer.Print strlinea2: i = i + 1
                    End If
                    
                    If InStr(rsServicio("DES_LEYENDA"), "3[") > 0 Then
                        strlinea3 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "3[") + 2), InStr(rsServicio("DES_LEYENDA"), "]3") - ((InStr(rsServicio("DES_LEYENDA"), "3[") + 2)))
                        If strlinea3 <> "" Then Printer.Print strlinea3: i = i + 1
                    End If
                    
                    If InStr(rsServicio("DES_LEYENDA"), "4[") > 0 Then
                        strlinea4 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "4[") + 2), InStr(rsServicio("DES_LEYENDA"), "]4") - ((InStr(rsServicio("DES_LEYENDA"), "4[") + 2)))
                        If strlinea4 <> "" Then Printer.Print strlinea4: i = i + 1
                    End If
                    
                    If InStr(rsServicio("DES_LEYENDA"), "5[") > 0 Then
                        strlinea5 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "5[") + 2), InStr(rsServicio("DES_LEYENDA"), "]5") - ((InStr(rsServicio("DES_LEYENDA"), "5[") + 2)))
                        If strlinea5 <> "" Then Printer.Print strlinea5: i = i + 1
                    End If
                    
                    If InStr(rsServicio("DES_LEYENDA"), "6[") > 0 Then
                        strlinea6 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "6[") + 2), InStr(rsServicio("DES_LEYENDA"), "]6") - ((InStr(rsServicio("DES_LEYENDA"), "6[") + 2)))
                        If strlinea6 <> "" Then Printer.Print strlinea6: i = i + 1
                    End If
                    
                    If InStr(rsServicio("DES_LEYENDA"), "7[") > 0 Then
                        strlinea7 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "7[") + 2), InStr(rsServicio("DES_LEYENDA"), "]7") - ((InStr(rsServicio("DES_LEYENDA"), "7[") + 2)))
                        If strlinea7 <> "" Then Printer.Print strlinea7: i = i + 1
                    End If

                End If
            
            End If
            rsServicio.MoveNext
        Wend
        Set objServicio = Nothing
        If FlgAnulado = True Then
            Dim s As Variant
            s = Printer.FontSize
            Printer.FontSize = 17
            Printer.Print "DOCUMENTO ANULADO"
            Printer.FontSize = s
        
        End If
       '---------------------------------------------------------------------
       For j = i To intNumeroLineas
            If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
                If j = intNumeroLineas Then Exit For Else Printer.Print
            Else
                Printer.Print
            End If
       Next j

       Printer.Print Trim(objUsuario.Empresa) & Space(11) & "Total Venta " & Trim(objUsuario.SmbMoneda) & right("        " & Format(dblTotal, "####0.#0"), 8)

       Printer.Print "SON : " & strLetras & " " & objUsuario.DesLgMoneda
       'Printer.Print objVenta.Out_Direccion
        
       'Set rstConvenio = ListaDoc(objUsuario.CodigoEmpresa, TipoDocumento, NumDocumento)
       
'''''''''''''''''       If Not rstConvenio.EOF Then
'''''''''''''''''            strGlosaConv = DevTextoDocCnv(objUsuario.CodigoEmpresa, TipoDocumento, NumDocumento)
'''''''''''''''''            If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio Then
'''''''''''''''''                If rstConvenio("FLG_DOC_COPAGO").Value = "1" Then
'''''''''''''''''                    Printer.Print "B E N E F I C I A R I O  - " & strGlosaConv
'''''''''''''''''                Else
'''''''''''''''''                    Printer.Print "I N S T I T U C I O N  - " & strGlosaConv
'''''''''''''''''                End If
'''''''''''''''''
'''''''''''''''''                Printer.Print "(" & rstConvenio("PCT_BENEFICIARIO").Value & "%) - #REF " & rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value & " " & rstConvenio("NUM_DOCUMENTO_REFER").Value
'''''''''''''''''
'''''''''''''''''            End If
'''''''''''''''''       End If
        
       dblMtoTotalRef = 0
       
       dblMtoTotalRef = DevImpDocCnv(objUsuario.CodigoEmpresa, numDocumento, TipoDocumento, objUsuario.CodigoLocal)
       
       If Not rstConvenio.EOF Then
            strGlosaConv = DevTextoDocCnv(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento, , objUsuario.CodigoLocal)
            If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio Then
                If rstConvenio("FLG_DOC_COPAGO").Value = "1" Then
                    'Printer.Print "B E N E F I C I A R I O  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%)"
                    Printer.Print "B E N E F I C I A R I O  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO_ORIG").Value & "%)"
                    'xStrBenef = strGlosaConv
                    objVenta.Beneficiario = strGlosaConv
                    'Printer.Print "#REF " & rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value & " " & rstConvenio("NUM_DOCUMENTO_REFER").Value & " (" & 100 - rstConvenio("PCT_BENEFICIARIO").Value & "%) - S/.: " & Format(dblMtoTotalRef, "###,##0.00") & "  "
                    If objConvenio.FnDev_Deducible(rstConvenio("COD_CONVENIO").Value) = "0" Then
                       Printer.Print "#REF " & rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value & " " & rstConvenio("NUM_DOCUMENTO_REFER").Value & " (" & 100 - rstConvenio("PCT_BENEFICIARIO_ORIG").Value & "%) - S/.: " & Format(dblMtoTotalRef, "###,##0.00") & "  "
                    End If
                Else
                    Set objCliente = New clsCliente
                    objVenta.Beneficiario = objCliente.Nombre(IIf(IsNull(rstConvenio("COD_CLIENTE").Value), "", rstConvenio("COD_CLIENTE").Value))
                    Set objCliente = Nothing
                    Printer.Print "I N S T I T U C I O N  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%)"
                    If dblMtoTotalRef > -1 Then
                        Printer.Print "#REF " & rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value & " " & rstConvenio("NUM_DOCUMENTO_REFER").Value & " (" & 100 - rstConvenio("PCT_BENEFICIARIO").Value & "%) - S/.: " & Format(dblMtoTotalRef, "###,##0.00") & Space(2) & "BENEF:" & Space(2) & objVenta.Beneficiario
                    Else
                        Printer.Print "BENEF:" & Space(2) & objVenta.Beneficiario
                    End If
                End If
                
                'Printer.Print "#REF " & rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value & " " & rstConvenio("NUM_DOCUMENTO_REFER").Value & " (" & 100 - rstConvenio("PCT_BENEFICIARIO").Value & "%) - S/.: " & Format(dblMtoTotalRef, "###,##0.00") & "  "
                
            End If
        End If
        
        Set rstDonacion = objDocumento.Impresion(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento, "1")

        If rstDonacion.EOF Then
'            If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
'                If strGlosaConv <> "" Then Else Printer.Print ': Printer.Print
'            Else
'                'Printer.Print
'                Printer.Print
'            End If
        Else
            dblDonacion = 0
            While Not rstDonacion.EOF
                dblDonacion = dblDonacion + rstDonacion("IMP_CON_REDONDEO").Value
                Printer.Print left(rstDonacion("DES_HIJO").Value, 54) & "   " & Trim(objUsuario.SmbMoneda) & right("        " & Format(rstDonacion("IMP_CON_REDONDEO").Value, "####0.#0"), 8)
                rstDonacion.MoveNext
            Wend
            Printer.Print Space$(30) & "TOTAL VENTA Y DONACION " & Trim(objUsuario.SmbMoneda) & "    " & right("        " & Format(dblTotal + dblDonacion, "####0.#0"), 8)
        End If
        
        ' AGREGADO POR DJARA PARA IMPRIMIR EN CASO DE CAMBIO DE DIRECCION
        If objUsuario.ImprimirDireccion = True Then
            Printer.Font.name = "Draft 20cpi"
            Printer.Print "" & Mensaje1(objUsuario.CodigoLocal, "BOL")
            Printer.Print Space(15) & "" & Mensaje2(objUsuario.CodigoLocal, "BOL")
            '"NUEVA DIR: " + left(objUsuario.direccion, AnchoBoleta - 11)
            Printer.Font.name = "Draft 17cpi"
        Else
            Printer.Print
        End If
        ' DJARA - 03/09/2009
        
        'COMENTADO POR PHERRERA 17/09/08 -- PARECE QUE NUNCA SE USO Y SE LLENO LA VARIABLE
        Printer.Print "" 'objVenta.Out_Direccion

        strMensaje = DevMensaje(objUsuario.CodigoEmpresa, numDocumento, TipoDocumento, objUsuario.CodigoLocal)

        Dim rsNewMess As oraDynaset
        Set rsNewMess = ListaMensajes
        While Not rsNewMess.EOF
            Printer.FontSize = 7
            'Printer.FontName = "Draft 10cpi"
            Printer.Print "" & left(rsNewMess("DES_PARAMETRO").Value, 70) & "-"
            Printer.Print "" & Mid(rsNewMess("DES_PARAMETRO").Value, 71, 70)
            Printer.Print "" & Mid(rsNewMess("DES_PARAMETRO").Value, 141, 100)
            rsNewMess.MoveNext
        Wend

        Printer.Print Space(20) + strCodMaquina + "   " + strCodUsuarioDependiente + "   " + IIf(IsNull(strMensaje), "", strMensaje)
        'intTamañoEspacio = Int((AnchoBoleta - Len(IIf(IsNull(strMensaje), "", strMensaje))) / 2)
        'Printer.Print Space(intTamañoEspacio) + IIf(IsNull(strMensaje), "", strMensaje)
       
        Printer.EndDoc
        
        ImprimeCuponMonedero TipoDocumento, numDocumento, FlgAnulado
        
   Else
        MsgBox "No se encontraron datos en el documento " + Chr(13) + "B/ " & rstImpresion("NUM_DOCUMENTO").Value, vbCritical, App.ProductName
        
   End If

    Exit Sub
ErrorImpresora:
    Err.Raise Err.Number, "clsImpresiones.Boleta", Err.Description
End Sub

Public Sub Guia(ByVal Cia As String, ByVal TipDoc As String, numGuia As String)
Dim recImpGuia As oraDynaset
Dim rstConvenio As oraDynaset
Dim strGlosaConv  As String
Dim i, j As Integer
Dim dblImpTotal As Double

Dim strNumDocumentoRef As String
Dim strCodTipodocRef As String
Dim dblPctCopago As Double
               
Dim dblIGV As Double
Dim p As Printer
Dim dblMtoTotalRef As Double

 On Error GoTo ErrorImpresora

    For Each p In Printers
       If UCase(p.PORT) = "LPT1:" Then
          Set Printer = p
          Exit For
       End If
    Next p
    
    Dim odynCab As oraDynaset
    Set odynCab = gclsOracle.FN_Cursor("BTLCERO.PKG_GUIA_CLIENTE.FN_CABECERA", 0, numGuia, objUsuario.CodigoLocal)
    If odynCab.RecordCount = 0 Then MsgBox "No se encontró Guía en la BD", vbCritical, "ERROR": Exit Sub
    
    If odynCab("COD_FORMATO").Value = "005" Then 'formato corto
    
        'jrazuri 18/07/2007
        '------------------------------------------
        
        If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
            Dim objImpresion As New clsImpresion
         
            If Not objImpresion.cargaParametros("005") Then
                Exit Sub
            End If
        Else
            
            Printer.Width = 225 * 56.7
            Printer.Height = 152 * 56.7
            
        End If
        
        dblIGV = PctImpuesto(Cia, TipDoc, numGuia)
        
        Set recImpGuia = ImpresionGuia(Cia, numGuia, objUsuario.CodigoLocal)
        
        dblImpTotal = 0
        If Not recImpGuia.EOF Then
                    strNumDocumentoRef = IIf(IsNull(recImpGuia("NUM_DOCUMENTO_REF").Value), "", recImpGuia("NUM_DOCUMENTO_REF").Value)
                    strCodTipodocRef = IIf(IsNull(recImpGuia("COD_TIPODOC_REF").Value), "", recImpGuia("COD_TIPODOC_REF").Value)
                    strGlosaConv = DevTextoDocCnv(objUsuario.CodigoEmpresa, "" & recImpGuia("COD_TIPODOC_REF").Value, "" & recImpGuia("NUM_DOCUMENTO_REF").Value, , objUsuario.CodigoLocal)
                    dblPctCopago = IIf(IsNull(recImpGuia("PCT_COPAGO").Value), 0, recImpGuia("PCT_COPAGO").Value)
                    dblImpTotal = recImpGuia("MTO_TOTAL").Value
                    Printer.FontName = "Draft 17cpi"
                    Printer.Print "."
                    Printer.Print
                    Printer.Print
                    Printer.Print
                    Printer.Print
                    Printer.Print
                    
                    ' AGREGADO POR DJARA PARA IMPRIMIR EN CASO DE CAMBIO DE DIRECCION
                    If objUsuario.ImprimirDireccion = True Then
                        Printer.Font.name = "Draft 20cpi"
                        Printer.Print "NUEVA DIR: " + objUsuario.Direccion
                        Printer.Font.name = "Draft 17cpi"
                    Else
                        Printer.Print
                    End If
                    ' DJARA - 03/09/2009

                    'Printer.Print
                    Printer.Print Space$(13) & recImpGuia("FCH_EMISION").Value
                    'Printer.Print
                    Printer.Print Space$(60) & recImpGuia("NUM_GUIA").Value
                    Printer.Print Space$(22) & recImpGuia("COD_BENEFICIARIO").Value & "-" & strGlosaConv
                    Printer.Print Space$(22) & recImpGuia("DES_RAZON_SOCIAL").Value
                    Printer.Print Space$(28) & recImpGuia("DIR_ENTREGA").Value & Space(35) & objUsuario.LabelImpConvGuia02 & " " & Trim(recImpGuia("PCT_COPAGO").Value) & "%"
                    Printer.Print Space$(15) & recImpGuia("NUM_DOCUMENTO_ID").Value
                    Printer.Print
                    Printer.Print
                    Printer.Print
                    'Printer.FontName = "Draft 12cpi"
                    i = 0
                    Printer.Print Space(63) & CME("Und", 3, "D") & CME("  ", 3, "C") & CME("Fra", 3, "D") & CME("  ", 3, "C") & CME("Lote", 25, "I") & CME("F.Vmto", 10, "C") & CME("Importe", 15, "D")
                    While Not recImpGuia.EOF
                                Printer.Print Space$(5) & CME(recImpGuia("COD_PRODUCTO").Value, 8, "I") & _
                                    CME(recImpGuia("DES_PRODUCTO").Value, 50, "I") & _
                                    CME(recImpGuia("CTD_PRODUCTO").Value, 3, "D") & _
                                    CME("  ", 3, "C") & _
                                    CME(recImpGuia("CTD_PRODUCTO_FRAC").Value, 3, "D") & _
                                    CME("  ", 3, "C") & _
                                    CME(IIf(IsNull(recImpGuia("NUM_LOTE").Value), "", recImpGuia("NUM_LOTE").Value), 25, "I") & _
                                    CME(IIf(IsNull(recImpGuia("FCH_VENCIMIENTO").Value), "", recImpGuia("FCH_VENCIMIENTO").Value), 10, "C") & _
                                    CME(Format(recImpGuia("IMP_VENTA").Value, "###,##0.00"), 15, "D")
                        'Impresión de la partida arancelaria
                        'jlopez
                        '06/05/2008
                        If dblIGV = 0 Then
                           If Not IsNull(recImpGuia("NUM_PART_ARANCEL").Value) Then
                              Printer.Print "        P.A. : " & recImpGuia("NUM_PART_ARANCEL").Value
                              i = i + 1
                           End If
                        End If
                        recImpGuia.MoveNext
                        i = i + 1
                    Wend
                    'Printer.Print Space$(5) & CME("--------------", 82, "D")
                    'Printer.Print Space$(5) & CME(Format(dblImpTotal, "###,##0.00"), 82, "D")
                    
                    Printer.Print Space$(5) & CME("--------------", 120, "D")
                    Printer.Print Space$(5) & CME(Format(dblImpTotal, "###,##0.00"), 120, "D")
                    
                    For j = 1 To 15 - i
                        Printer.Print
                    Next
                    
                    'Printer.FontName = "Arial"
    
                    'If objGuia.intImpresionZonal(OraDatabase, strCodConvenio) = 1 Then
                    '    Printer.Print Space$(5) & CME("Cod.Zonal:", 25, "I") & CME(strCodZonal, 50, "I")
                    '    Printer.Print Space$(5) & CME("Crédito Senati: Institución 80% : S/. ", 37, "I") & CME(Round(dblImpSinCopago * 0.8, 2), 50, "I")
                    '    Printer.Print Space$(5) & CME("Crédito Senati: Beneficiario 20%: S/. ", 37, "I") & CME(Round(dblImpSinCopago * 0.2, 2), 50, "I")
                    'End If
    
                    If strNumDocumentoRef <> "" And strCodTipodocRef <> "" Then
                        Set rstConvenio = ListaDoc(objUsuario.CodigoEmpresa, strCodTipodocRef, strNumDocumentoRef)
                        dblMtoTotalRef = 0
                        dblMtoTotalRef = DevImpDocCnv(objUsuario.CodigoEmpresa, numGuia, TipDoc, objUsuario.CodigoLocal)
    '''''''''''                    If Not rstConvenio.EOF Then
    '''''''''''                         strGlosaConv = DevTextoDocCnv(objUsuario.CodigoEmpresa, TipDoc, NumGuia)
    '''''''''''                         If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio Then
    '''''''''''                             If recImpGuia("FLG_DOC_COPAGO").Value = "1" Then
    '''''''''''                                 Printer.Print "B E N E F I C I A R I O  - " & strGlosaConv
    '''''''''''                             Else
    '''''''''''                                 Printer.Print "I N S T I T U C I O N  - " & strGlosaConv
    '''''''''''                             End If
    '''''''''''                             Printer.Print "(" & dblPctCopago & "%) - #REF " & strCodTipodocRef & " " & strNumDocumentoRef
    '''''''''''                         End If
    '''''''''''                    End If
                        If Not rstConvenio.EOF Then
                             strGlosaConv = "" & DevTextoDocCnv(objUsuario.CodigoEmpresa, TipDoc, numGuia, , objUsuario.CodigoLocal)
                             If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio Then
                                 If rstConvenio("FLG_DOC_COPAGO").Value = "1" Then
                                     Printer.Print "I N S T I T U C I O N  - " & strGlosaConv & " (" & 100 - rstConvenio("PCT_BENEFICIARIO").Value & "%)"
                                 Else
                                     Printer.Print "B E N E F I C I A R I O  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%)"
                                 End If
                                 Printer.Print "#REF " & strCodTipodocRef & " " & strNumDocumentoRef & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%) - S/.: " & Format(dblMtoTotalRef, "###,##0.00") & "- BENEF " & DevTextoDocCnv(objUsuario.CodigoEmpresa, TipDoc, numGuia, 1, objUsuario.CodigoLocal)
    ''''''                             Printer.Print DevNomRepartidor(objUsuario.CodigoEmpresa, _
    ''''''                                      objUsuario.CodigoLocal, _
    ''''''                                      TipDoc, _
    ''''''                                      NumGuia)
                                 If rstConvenio("COD_CONVENIO") = objUsuario.CodConvenioSenati Then
                                    Printer.Print Space$(5) & CME("Crédito Senati: Institución 80% : S/. ", 37, "I") & CME(Round(dblMtoTotalRef * 0.8, 2), 50, "I")
                                    Printer.Print Space$(5) & CME("Crédito Senati: Beneficiario 20%: S/. ", 37, "I") & CME(Round(dblMtoTotalRef * 0.2, 2), 50, "I")
                                 End If
                             End If

                        End If
                    End If
    
                    '**********************************************
                    '**********************************************
                    '**********************************************
                        Printer.FontSize = 10
                        Dim objDocumento As New clsDocumento
                        Dim rstTipoCampo As oraDynaset
                        
                        Set rstTipoCampo = objDocumento.ListaTipoCampo(Cia, TipDoc, numGuia, objUsuario.CodigoLocal)
                        
                        With rstTipoCampo
                          i = 0
                          If Not .EOF Then
                              While Not .EOF
                                  Printer.Print Space(10) & rstTipoCampo("DES_TIPO_CAMPO").Value & " " & rstTipoCampo("DES_VALOR").Value
                                  i = i + 1
                                  .MoveNext
                              Wend
                          End If
                        End With
                        Set objDocumento = Nothing
                        
                        If Not objVenta.CodigoConvenio = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_CNV_RIMAC") Then ''CAMBIAR POR CONSTANTE EN RIMAC
                       Printer.Print DevNomRepartidor(objUsuario.CodigoEmpresa, _
                                          objUsuario.CodigoLocal, _
                                          TipDoc, _
                                          numGuia)
                        End If
                        Dim objConvenio As New clsConvenio
                        Dim rsCabeceraReceta As oraDynaset
                        Set rsCabeceraReceta = objConvenio.ListaCabeceraReceta(TipDoc, numGuia)
                        
                        Dim o As Integer
                        'este se agrego
                        Do Until rsCabeceraReceta.EOF
                            While o < rsCabeceraReceta.Fields.Count
                                Printer.Print Space(10) & rsCabeceraReceta.Fields(o).name & ": " & rsCabeceraReceta(o).Value
                                o = o + 1
                            Wend
                            rsCabeceraReceta.MoveNext
                        Loop
                        '09/11/07 comentado por pherrera no imprimia nada
    '                    While o < rsCabeceraReceta.Fields.Count
    '                        While Not rsCabeceraReceta.EOF
    '                        Printer.Print Space(10) & rsCabeceraReceta.Fields(o).Name & ": " & rsCabeceraReceta(o).Value
    '                        rsCabeceraReceta.MoveNext
    '                        Wend
    '                        o = o + 1
    '                    Wend
                        Printer.FontName = "Draft 12cpi"
                        Set objConvenio = Nothing
                    '**********************************************
                    '**********************************************
                    '**********************************************
    
                    For j = 1 To 4
                        Printer.Print
                    Next
    
                    Printer.FontName = "Draft 17cpi"
    
                    Printer.Print CME("X", 135, "D")
                    Printer.Print CME(objUsuario.LabelImpConvGuia01, 135, "D")
                    Dim strMensaje  As String
                    strMensaje = DevMensaje(objUsuario.CodigoEmpresa, numGuia, "GRL", objUsuario.CodigoLocal)
                    Printer.Print Space(50) + IIf(IsNull(strMensaje), "", strMensaje)

                    Printer.EndDoc
    
            End If
            
        Set rstConvenio = Nothing
        Set recImpGuia = Nothing
    Else
        'IMPRESION LARGA
        ImprimeGuiaLarga odynCab, numGuia
    End If
        Exit Sub
    
ErrorImpresora:
    Err.Raise Err.Number, "clsDocumento.Guia", Err.Description
End Sub

Private Sub ImprimeGuiaLarga(ByRef odynCab As oraDynaset, ByVal vstrNumguia As String)
'Dim odynCab As oraDynaset
Dim odynDet As oraDynaset
 
    'Cabecera'
    'Set odynCab = gclsOracle.FN_Cursor("BTLCERO.PKG_GUIA_CLIENTE.FN_CABECERA", 0, vstrNumguia)
    'If odynCab.RecordCount = 0 Then MsgBox "No se encontró Guía en la BD", vbCritical, "ERROR": Exit Sub
        
    'Detalle'
    Set odynDet = gclsOracle.FN_Cursor("BTLCERO.PKG_GUIA_CLIENTE.FN_DETALLE", 0, vstrNumguia, odynCab("NUM_PEDIDO").Value, "1", objUsuario.CodigoLocal)
    If odynDet.RecordCount = 0 Then MsgBox "Documento sin detalle", vbCritical, "ERROR": Exit Sub
'--------------------------------------------------------------------------------------------------------------'
    Dim intLinea%
               
    On Error GoTo ErrorImpresora
    
    Dim objImpresion As New clsImpresion
            
    If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
         If Not objImpresion.cargaParametros("001") Then
             Exit Sub
         End If
     Else
         Printer.Width = 1440 * 8.5
         Printer.Height = 1440 * 12
    End If
    
    '-----------
    Printer.Print
    Printer.Print
            Dim rstTipoCampo As oraDynaset
            Dim i As Integer, j As Integer
            Set rstTipoCampo = ListaTipoCampo(objUsuario.CodigoEmpresa, "GRL", vstrNumguia, objUsuario.CodigoLocal)
            
            With rstTipoCampo
              i = 0
              If Not .EOF Then
                  While Not .EOF
                      Printer.Print Space(28) & rstTipoCampo("DES_TIPO_CAMPO").Value & " " & rstTipoCampo("DES_VALOR").Value
                      i = i + 1
                      .MoveNext
                  Wend
              End If
            End With
            For j = 1 To 4 - i
              Printer.Print
            Next j
    
    'Printer.Print
    'Printer.Print
    'Printer.Print
    'Printer.Print
    Printer.Print
    Printer.Print
    Printer.FontName = "Draft 10cpi"
    Printer.Print Space(55) & Format(odynCab("NUM_GUIA").Value, "000-0000000")
    Printer.FontName = "Draft 12cpi"
    Printer.Print
    Printer.Print
    Printer.Print
                
    ''            Dim strObs$
    ''            If odynCab("OBS_GUIA_CLIENTE").Value <> "*" Then
    ''                strObs = "Obs :" & odynCab("OBS_GUIA_CLIENTE").Value
    ''            End If
    ''            Printer.FontName = "Draft 17cpi"
    Printer.Print "DROGUERIA" 'strObs
    ''            Printer.FontName = "Draft 15cpi"
    Printer.Print
    'Printer.Print Space(18) & Format(rstDate(0).Value, "DD") & " de " & Left(Format(rstDate(0).Value, "MMMM") & Space(10), 10) & " del " & Format(rstDate(0).Value, "YYYY") & Space(69) & "X"
    Printer.Print Space(18) & Format(odynCab("FCH_EMISION").Value, "DD") & " de " & left(Format(odynCab("FCH_EMISION").Value, "MMMM") & Space(10), 10) & " del " & Format(odynCab("FCH_EMISION").Value, "YYYY") & Space(46) & "X"
    Printer.Print
    Printer.Print
    
    Dim strDestinatario$, strOC$
                  
    If odynCab("BROKER").Value <> "*" Then
        strDestinatario = odynCab("BROKER").Value
    End If
    
    If odynCab("OBS_ORDEN_COMPRA").Value <> "*" Then
        strOC = "O/C Cliente : " & odynCab("OBS_ORDEN_COMPRA").Value
    End If
    Printer.Print Space(23) & strDestinatario
    Printer.Print Space(23) & strOC
    Printer.Print Space(23) & odynCab("CLIENTE").Value
    Printer.FontName = "Draft 17cpi"
    Printer.Print Space(32) & odynCab("DIR_ENTREGA_PROD").Value
    Printer.FontName = "Draft 10cpi"
    Printer.Print Space(10) & odynCab("PERSONA").Value
    Printer.FontName = "Draft 17cpi"
    Printer.Print String(130, "-")
    Printer.Print right(Space(2) & "#", 2) & Space(1) & _
                   right(Space(6) & "Cant.", 6) & Space(1) & _
                   left("Código" & Space(6), 6) & Space(2) & _
                   left("Descripción" & Space(60), 60) & Space(2) & _
                   left("Laboratorio" & Space(15), 15) & Space(2) & _
                   left("Lote" & Space(15), 15) & Space(1) & _
                   left("F.Ven" & Space(7), 7) & Space(2) & _
                   right(Space(7) & "Ctd.Lte", 7)
    Printer.Print String(130, "-")
    '---
    intLinea = 0
    Dim strCodProdAnte As String
    strCodProdAnte = "*"
    While Not odynDet.EOF
          'Detalle
          If strCodProdAnte <> odynDet("COD_PRODUCTO").Value Then
            Printer.Print right(Space(2) & odynDet("NUM_ITEM").Value, 2) & Space(1) & _
                           right(Space(6) & odynDet("CANTIDADES").Value, 6) & Space(1) & _
                           left(odynDet("COD_PRODUCTO").Value & Space(6), 6) & Space(2) & _
                           left(odynDet("DES_PRODUCTO").Value & Space(60), 60) & Space(2) & _
                           left(odynDet("LAB").Value & Space(15), 15) & Space(2) & _
                           left(odynDet("NUM_LOTE").Value & Space(15), 15) & Space(1) & _
                           left("" & odynDet("FCH_VEN").Value & Space(7), 7) & Space(2) & _
                           right(Space(7) & odynDet("CANTIDADES_LOTE").Value, 7)
          
          Else
            Printer.Print right(Space(2) & "", 2) & Space(1) & _
                           right(Space(6) & "", 6) & Space(1) & _
                           left("" & Space(6), 6) & Space(2) & _
                           left("" & Space(60), 60) & Space(2) & _
                           left("" & Space(15), 15) & Space(2) & _
                           left(odynDet("NUM_LOTE").Value & Space(15), 15) & Space(1) & _
                           left("" & odynDet("FCH_VEN").Value & Space(7), 7) & Space(2) & _
                           right(Space(7) & odynDet("CANTIDADES_LOTE").Value, 7)
          End If
          strCodProdAnte = odynDet("COD_PRODUCTO").Value
          'dblTotal = dblTotal + rstDetPed("MTO_PRECIO_VTA_BTL").Value
          
          intLinea = intLinea + 1
          odynDet.MoveNext
    Wend
    Printer.Print
    
    Dim intMaxLineas As Integer
    
    Dim objFormato As New clsFormato
    intMaxLineas = objFormato.NumLineas("001")
    Set objFormato = Nothing
    
    For i = 1 To intMaxLineas - intLinea
        Printer.Print
    Next i
    Printer.Print
    Dim strObs$
    If odynCab("OBS_GUIA_CLIENTE").Value <> "*" Then
        strObs = "Obs :" & odynCab("OBS_GUIA_CLIENTE").Value
    End If
    Printer.FontName = "Draft 17cpi"
    Printer.Print strObs
    Printer.Print Space(10) & gclsOracle.FN_Valor("BTLCERO.PKG_PEDIDO.FN_MSJ_BOTIQUIN", odynCab("NUM_GUIA").Value)
    Printer.FontName = "Draft 15cpi"
    
    Printer.EndDoc
    
    Exit Sub
ErrorImpresora:
    If MsgBox("Existe un Problema con la Impresora" & Chr(13) & _
             "Error: " & CStr(Err.Number) & " - " & Err.Description & Chr(13) & _
             "¿Desea Esperar a ser Resuelto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
          Resume
       Else
          Exit Sub
     End If
End Sub

Function ImprimirRecibo(ByVal Cia As String, ByVal NumeroDocumento As String, ByVal TipoDocumento, ByVal CodigoLocal As String, Optional FlgAnulado As Boolean = False)
Dim rsCabecera As oraDynaset
Set rsCabecera = Me.ListaCabecera(Cia, NumeroDocumento, TipoDocumento, CodigoLocal)
Dim rsDetalle As oraDynaset
Set rsDetalle = Me.ListaDetalle(Cia, NumeroDocumento, TipoDocumento, CodigoLocal)

'======================================================================================================================================

Printer.FontName = "Draft 17cpi"
Printer.Width = 225 * 56.7
Printer.Height = 93 * 56.7

While Not rsDetalle.EOF

         Printer.Print "Nro. RECIBO: " & Format(rsCabecera("NUM_DOCUMENTO").Value, "000000") & Space(10) & "Fecha: " & Format(rsCabecera("FCH_REGISTRA"), "dd/mm/yyyy") & Space(15) & "Nro. RECIBO: " & Format(rsCabecera("NUM_DOCUMENTO").Value, "000000") & Space(20) & "Fecha: " & Format(rsCabecera("FCH_REGISTRA"), "dd/mm/yyyy")
         Printer.Print Space(29) & "Hora : " & Format(rsCabecera("FCH_REGISTRA"), "HH:MM:SS AM/PM ") & Space(52) & "Hora : " & Format(rsCabecera("FCH_REGISTRA"), "HH:MM:SS AM/PM ")
    '--     Printer.FontBold = True
         Printer.Print
         Printer.Print Space(1) & " BOTICA TORRES DE LIMATAMBO S.A.C. " & Space(37) & " BOTICA TORRES DE LIMATAMBO S.A.C. "
         Printer.Print Space(1) & " --------------------------------- " & Space(37) & " --------------------------------- "
         Printer.Print
         Printer.Print "Entidad            : " & "   " & rsDetalle("DES_ENTIDAD").Value & Space(30) & "Entidad            :" & "   " & rsDetalle("DES_ENTIDAD").Value
         Printer.Print
         Printer.Print "Suministro Nª      : " & "   " & rsDetalle("NUM_SERVICIO").Value & "-" & Space(41) & "Suministro Nª      :" & "   " & rsDetalle("NUM_SERVICIO").Value
         Printer.Print
         Printer.Print "Importe Pagado     : " & "   " & Format(rsDetalle("MTO_SUBTOTAL").Value, "#0.00") & Space(44) & "Importe Pagado     : " & "   " & Format(rsDetalle("MTO_SUBTOTAL").Value, "#0.00")
         Printer.Print
         Printer.Print
         Printer.Print
         'Printer.Print "(*)  Incluye IGV" & Space(56) & "(*)  Incluye IGV"
         Printer.Print String(40, "-") & Space(34) & String(40, "-")
         Printer.EndDoc
         
    rsDetalle.MoveNext
Wend

End Function

'Agregado para controlar el cambio de ticketera local a red 21/08/2012 mlevano
Public Sub LocalORed(ByVal Cia As String, ByVal Maquina As String, ByVal TipoDocumento As String)
    On Error GoTo CntrlErr
        
        Dim NomImpresora As String
        Dim n As String
        
        n = "" & gclsOracle.FN_Valor("BTLPROD.FN_TKB_LOCAL_O_RED", Maquina, TipoDocumento, Cia)
        
        If (Trim(gclsOracle.FN_Valor("BTLPROD.FN_NOM_IMPRESORA_TKB", Maquina))) <> "0" Then
            '1 = Red ; 0 = Local
            If n = "1" Then
                NomImpresora = "" & gclsOracle.FN_Valor("BTLPROD.FN_NOM_IMPRESORA_TKB", Maquina)
                Open "c:\red.bat" For Output As #1
                Print #1, "@echo off"
                Print #1, "NET USE LPT1 /D /Y"
                Print #1, "NET USE LPT1 "; NomImpresora; " /PERSISTENT:YES"
                Print #1, "TASKKILL /IM cmd.exe /F"
                Close #1
                Shell "c:\red.bat", vbHide
                'Shell "c:\red.bat", vbHide
            Else
                Open "c:\local.bat" For Binary As #1
                Put #1, , "@echo off" & Chr$(13) & Chr$(10)
                Put #1, , "NET USE LPT1 /D /Y" & Chr$(13) & Chr$(10)
                Put #1, , "TASKKILL /IM cmd.exe /F"
                Close #1
                Shell "c:\local.bat", vbHide
            End If
        End If
CntrlErr:
    n = ""
    
'Fin de actualizacion 21/08/2012
End Sub


Public Function ListaNumeroDisponible(ByVal Cia As String, ByVal Maquina As String, ByVal TipoDocumento As String) As String

    On Error GoTo CntrlErr
        
        Call LocalORed(Cia, Maquina, TipoDocumento)
        ListaNumeroDisponible = "" & gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO.FN_LISTA_CORRELATIVO", Maquina, TipoDocumento, Cia)
        
        Exit Function
CntrlErr:
    'Err.Raise Err.Number, "clsDocumento", Err.Description
    ListaNumeroDisponible = ""
End Function

Public Function ListaDetalleCXR(ByVal vstrCia As String, _
                              ByVal vstrNumDoc As String, _
                              ByVal vstrTipoDoc As String, _
                              ByVal vstrCodLocal As String) As oraDynaset

    On Error GoTo CntrlErr
    
    Set ListaDetalleCXR = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_DET_DOCUMENTO_CXR", 0, vstrCia, vstrNumDoc, vstrTipoDoc, vstrCodLocal)
    
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaDetalleCXR", Err.Description
End Function

'Se agregaron los parametros CadNroLote, CadFchVencimiento, TipDocReferencia y NumDocReferencia
'para proyecto Logistica III - DJara - 22/10/2008
Public Function GeneraGuiaLocal(ByRef pRetNumDocumento As String, _
                            ByVal pCodCIA As String, _
                            ByVal pCodLocal As String, _
                            ByVal pCodLocalDestino As String, _
                            ByVal pCodMaquina As String, _
                            ByVal pCodTipoDoc As String, _
                            ByVal pCodMotivoGuia As String, _
                            ByVal pCodUsuario As String, _
                            ByVal pObservaciones As String, _
                            ByVal pCadCodProducto As String, _
                            ByVal pCadCtdUEnvio As String, _
                            ByVal pCadCtdFEnvio As String, _
                            ByVal pCadProdFormulario As String, _
                            ByVal pCadNumFormulario As String, _
                            ByVal CadNroLote As String, _
                            ByVal CadFchVencimiento As String, _
                            ByVal TipDocReferencia As String, _
                            ByVal NumDocReferencia As String, _
                            ByVal vCodTipoDev As String, _
                            ByVal vCodMotivoDev As String, _
                            ByVal pCodFacSap As String) As String
                            

On Error GoTo CtrlErr

    Dim gvarValores As Variant
    Dim gvarIO  As Variant
                          
    
                      
    gvarValores = Array(pRetNumDocumento, pCodCIA, pCodLocal, pCodLocalDestino, pCodMaquina, _
                        pCodTipoDoc, pCodMotivoGuia, pCodUsuario, pObservaciones, pCadCodProducto, _
                        pCadCtdUEnvio, pCadCtdFEnvio, pCadProdFormulario, pCadNumFormulario, "1", _
                        "005", CadNroLote, CadFchVencimiento, TipDocReferencia, NumDocReferencia, _
                        vCodTipoDev, vCodMotivoDev, pCodFacSap)
                                            
    gvarIO = Array(Salida, entrada, entrada, entrada, entrada, _
                   entrada, entrada, entrada, entrada, entrada, _
                   entrada, entrada, entrada, entrada, entrada, _
                   entrada, entrada, entrada, entrada, entrada, _
                   entrada, entrada, entrada)
                                       
    GeneraGuiaLocal = gclsOracle.SP("NUEVO.PKG_GUIA.SP_GENERA_GUIA_LOCAL", gvarValores, gvarIO)
    'GeneraGuiaLocal = gclsOracle.SP("NUEVO.PKG_GUIA.SP_GENERA_GUIA_DEV", gvarValores, gvarIO)
    If GeneraGuiaLocal <> "" Then
        Err.Raise 2000, "Function GeneraGuiaLocal", GeneraGuiaLocal
    End If
    pRetNumDocumento = gvarValores(0)
Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.GeneraGuiaLocal", Err.Description
End Function

Public Sub ImprimirGuiaTransferencia(ByVal numGuia As String)

    Dim objGuia As New clsGuia
    Dim rsCab As oraDynaset
    Dim rsDet As oraDynaset

    Dim SumItem As Long
    Dim SumCant As Long
    Dim SumCodi As Long
    Dim a As String
    Dim p As Printer
    Dim strSQL As String
 On Error GoTo ErrorImpresora

    For Each p In Printers
       If UCase(p.PORT) = "LPT1:" Then
          Set Printer = p
          Exit For
       End If
    Next p
   
    'jrazuri 18/07/2007
    '------------------------------------------
    
    If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
        Dim objImpresion As New clsImpresion
     
        If Not objImpresion.cargaParametros("005") Then
            Exit Sub
        End If
    Else
        Printer.Width = 1440 * 8.5
        Printer.Height = 1440 * 8
    End If
    
    '------------------------------------------
   
    Set rsCab = objGuia.ImprimeCabecera(numGuia)
    If rsCab.EOF Then Exit Sub
    Set rsDet = objGuia.ImprimeDetalle(numGuia)
    If rsDet.EOF Then Exit Sub
        
        SumItem = 0
        SumCant = 0
        SumCodi = 0

        'Printer.Width = 225 * 56.7
        'Printer.Height = 152 * 56.7
        '''''''''''''''''''''''''''''''''''
                Printer.FontName = "Draft 17cpi"
                Printer.Print "."
                Printer.FontName = "Draft 10cpi"
                Printer.Print
                Printer.Print
                'Printer.Print
                Printer.Print
                Printer.Print '"Local Origen : " & objUsuario.CodigoLocal
                Printer.Print
                Printer.FontName = "Draft 15cpi"
                Printer.Print Space$(10) & IIf(IsNull(rsCab("FCH_EMI").Value), "", rsCab("FCH_EMI").Value)
                Printer.Print Space$(13) & objUsuario.CodigoLocal & " " & IIf(IsNull(rsCab("DIR_ORIGEN").Value), "", rsCab("DIR_ORIGEN").Value)
                Printer.Print
                Printer.Print Space$(12) & Trim(rsCab(3).Value) & " - " & Trim(rsCab("DES_DESTINO").Value) '& Space$(50) & IIf(IsNull(rsCab("NUM_GUIA").Value), "", rsCab("NUM_GUIA").Value)
                Printer.Print Space$(18) & IIf(IsNull(rsCab("RS_ORIGEN").Value), "", rsCab("RS_DESTINO").Value) & Space$(35) & IIf(IsNull(rsCab("NUM_GUIA").Value), "", rsCab("NUM_GUIA").Value)
                Printer.Print Space$(20) & IIf(IsNull(rsCab("DIR_DESTINO").Value), "", rsCab("DIR_DESTINO").Value)
                Printer.Print Space$(10) & Trim(rsCab("RUC_DESTINO").Value) & Space(10) & "Motivo: " & rsCab("MOTIVO").Value
                Printer.Print
                Printer.FontName = "Draft 12cpi"
                'Printer.Print
                'Printer.Print
                'Printer.Print 'Space$(7) & IIf(IsNull(rsCab("RS_ORIGEN").Value), "", rsCab("RS_ORIGEN").Value) & " - " & IIf(IsNull(rsCab("DES_DESTINO").Value), "", rsCab("DES_DESTINO").Value)
                'Printer.Print 'Space$(7) & IIf(IsNull(rsCab("RUC_ORIGEN").Value), "", rsCab("RUC_ORIGEN").Value) & Space$(8) & "Movimiento : " & IIf(IsNull(rsCab("DES_MOTIVO_GUIA").Value), "", rsCab("DES_MOTIVO_GUIA").Value) & "    ORIGEN: BTL " & objUsuario.CodigoLocal & " - DESTINO: BTL " & Trim(rsCab(3).Value)
                'Printer.Print 'Space$(7) & IIf(IsNull(rsCab("DIR_DESTINO").Value), "", rsCab("DIR_DESTINO").Value)
                'Printer.Print
                'Printer.Print
                Printer.Print
                
                
                
                a = CME("CODIGO", 5, "I") & Space(1) & _
                    CME("PCJE", 4, "I") & Space(1) & _
                    CME("DESCRIPCION", 39, "I") & Space(1) & _
                    CME("CA", 5, "D") & _
                    CME("NTIDAD", 6, "I") & Space(3) & _
                    CME("LOTE", 20, "I") & CME("FEC.VEN", 8, "D")
                        
                Printer.Print a
                Printer.Print String(5 + 3 + 40 + 3 + 5 + 6 + 3 + 20 + 8, "=")
                
                
                While Not (rsDet.EOF)
                    SumItem = SumItem + 1
                    SumCant = SumCant + Val(IIf(IsNull(rsDet("CTD_PRODUCTO").Value), 0, rsDet("CTD_PRODUCTO").Value))
                    SumCodi = SumCodi + Val(IIf(IsNull(rsDet("COD_PRODUCTO").Value), "", rsDet("COD_PRODUCTO").Value))
                    
'                    a = CME(rsDet("NUM_ITEM").Value, 5, "I") & _
'                            IIf(Val(rsDet("CTD_PRODUCTO").Value) = 0, " ", rsDet("CTD_PRODUCTO").Value) & _
'                            IIf(Val(rsDet("CTD_PRODUCTO_FRAC").Value) = 0, " ", "f" & CME(rsDet("CTD_PRODUCTO_FRAC").Value, 5, "I")) & _
'                        CME(rsDet("COD_PRODUCTO").Value, 5, "I") & Space(5) & _
'                        CME(IIf(rsDet.EOF, "NO EXISTE PRODUCTO", IIf(IsNull(rsDet("DES_PRODUCTO").Value), "", rsDet("DES_PRODUCTO").Value)), 40, "I")
                    a = CME(rsDet("COD_PRODUCTO").Value, 5, "I") & Space(1) & _
                        CME(rsDet("DES_POLITICA").Value, 4, "I") & Space(1) & _
                        CME(IIf(rsDet.EOF, "NO EXISTE PRODUCTO", IIf(IsNull(rsDet("DES_PRODUCTO").Value), "", rsDet("DES_PRODUCTO").Value)), 39, "I") & Space(1) & _
                        CME(IIf(Val(rsDet("CTD_PRODUCTO_LOTE").Value) = 0, " ", rsDet("CTD_PRODUCTO_LOTE").Value), 5, "D") & _
                        CME(IIf(Val(rsDet("CTD_PRODUCTO_FRAC_LOTE").Value) = 0, " ", "f" & rsDet("CTD_PRODUCTO_FRAC_LOTE").Value), 6, "I") & Space(3) & _
                        CME(rsDet("NUM_LOTE").Value, 20, "I") & CME(rsDet("FCH_VEN").Value, 8, "D")

                    Printer.Print a
                    rsDet.MoveNext
                Wend
                
                
                
                
                Printer.Print
                strSQL = Trim(IIf(IsNull(rsCab("DES_OBSERVACIONES").Value), "", rsCab("DES_OBSERVACIONES").Value))
                Printer.Print Space(4) & IIf(Len(Trim(strSQL)) = 0, " ", "Observaciones: " & left(strSQL, 70 + InStr(Mid(strSQL, 71), " ")))
                If Len(strSQL) > 70 Then
                    Printer.Print Space(14) + Mid(strSQL, 70 + InStr(Mid(strSQL, 71), " "), 70)
                End If
                'Printer.Print Space(4) & "Items : " & Trim(Str(SumItem)) & "     Suma de cantidades : " & Trim(Str(SumCant)) & "     Suma de códigos : " & Trim(Str(SumCodi))
                Printer.Print Space(4) & "Responsable : " & rsCab("NOM_USUARIO").Value
                Printer.Print
                
                
                Printer.EndDoc

Exit Sub
ErrorImpresora:
    Err.Raise Err.Number, "clsDocumento.ImprimirGuiaTransferencia", Err.Description
End Sub

Public Function GeneraGuiaAjuste(ByRef pRetNumDocumento As String, _
                            ByVal pCodCIA As String, _
                            ByVal pCodLocal As String, _
                            ByVal pTipoAjuste As String, _
                            ByVal pCodUsuario As String, _
                            ByVal pObservaciones As String, _
                            ByVal pCadCodProducto As String, _
                            ByVal pCadCtdUEnvio As String, _
                            ByVal pCadCtdFEnvio As String, _
                            ByVal pCadImpProducto As String, _
                            ByVal pMotivoAjuste As String, _
                            Optional pFlgValidar As String = "0") As String

On Error GoTo CtrlErr

Dim gvarValores As Variant
Dim gvarIO  As Variant
                        gvarValores = Array(pRetNumDocumento, pCodCIA, pCodLocal, pTipoAjuste, pCodUsuario, pObservaciones, pCadCodProducto, pCadCtdUEnvio, pCadCtdFEnvio, pCadImpProducto, "", pMotivoAjuste, pFlgValidar)
                                            
                        gvarIO = Array(Salida, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada, entrada)
                                       
                        GeneraGuiaAjuste = gclsOracle.SP("BTLPROD.PKG_AJUSTE.SP_GRABA", gvarValores, gvarIO)

        pRetNumDocumento = gvarValores(0)
Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.GeneraGuiaAjuste", Err.Description
End Function

Public Sub NotaCred(ByVal Cia As String, ByVal CodLocal As String, ByVal TipoDoc As String, NumDoc As String)
Dim rsCab As oraDynaset
Dim rsDet As oraDynaset
Dim rsRef As oraDynaset
Dim a As String
Dim p As Printer
Dim i As Integer
Dim k As Integer
Dim strFchRef As String
   
    For Each p In Printers
       If UCase(p.PORT) = "LPT1:" Then
          Set Printer = p
          Exit For
       End If
    Next p
   
    Set rsCab = ListaCabecera(Cia, NumDoc, TipoDoc, CodLocal)
    If rsCab.EOF Then Exit Sub
    Set rsDet = ListaDetalle(Cia, NumDoc, TipoDoc, CodLocal)
    If rsDet.EOF Then Exit Sub
   
    Set rsRef = ListaCabecera(Cia, rsCab("NUM_DOCUMENTO_REFER").Value, rsCab("COD_TIPO_DOCUMENTO_REFER").Value, CodLocal)
    If rsRef.EOF Then
        strFchRef = ""
    Else
        strFchRef = IIf(IsNull(rsRef("FCH_REGISTRA").Value), "", Format(rsRef("FCH_REGISTRA").Value, "dd/mm/yyyy"))
    End If
   
   On Error GoTo ErrorImpresora
  
   Printer.Width = 220 * 56.7
   Printer.Height = 170 * 56.7
   
   Printer.Print
   Printer.FontName = "Draft 17cpi"
   Printer.Print "."
   Printer.FontName = "Draft 12cpi"
   Printer.Print
   Printer.Print
   Printer.Print
   Printer.Print Space$(59) & "N.I.T. " & objUsuario.Ruc
   Printer.Print
   Printer.Print
   Printer.Print Space$(62) & Mid(NumDoc, 1, 3) & "-" & Mid(NumDoc, 4, 7)
        
    ' AGREGADO POR DJARA PARA IMPRIMIR EN CASO DE CAMBIO DE DIRECCION
    If objUsuario.ImprimirDireccion = True Then
        Printer.Font.name = "Draft 20cpi"
        Printer.Print "NUEVA DIR: " + objUsuario.Direccion
        Printer.Font.name = "Draft 17cpi"
    Else
        Printer.Print
    End If
    ' DJARA - 03/09/2009
   
   Printer.Print
   Printer.Print Space$(11) & rsCab("FCH_REGISTRA").Value
   Printer.Print
   Printer.Print
   Printer.Print Space$(2) & CME(IIf(IsNull(rsCab("DES_RAZON_SOCIAL").Value), "", rsCab("DES_RAZON_SOCIAL").Value), 58) & Space$(10) & strFchRef
   Printer.Print
   Printer.Print Space$(2) & CME(IIf(IsNull(rsCab("NUM_RUC_EMPRESA").Value), "", rsCab("NUM_RUC_EMPRESA").Value), 62) & Space$(6) & Mid(rsCab("NUM_DOCUMENTO_REFER").Value, 1, 3) & "-" & Mid(rsCab("NUM_DOCUMENTO_REFER").Value, 4, 7)
   Printer.Print
   Printer.Print Space$(2) & IIf(IsNull(rsCab("DES_AUX_CLI_DIRECC").Value), "", rsCab("DES_AUX_CLI_DIRECC").Value)
   Printer.Print
   Printer.Print
   Printer.Print Space$(7) & rsCab("DES_SUBCONCEPTO").Value
   Printer.Print
   Printer.Print
   Printer.Print
   
   i = 0
   
                While Not (rsDet.EOF)
                    a = CME(rsDet("CTD_PRODUCTO").Value & IIf(rsDet("CTD_PRODUCTO_FRAC").Value > 0, "f" & Trim(str(rsDet("CTD_PRODUCTO_FRAC").Value)), ""), 7, "D") & Space$(3)
                    a = a & CME(rsDet("COD_PRODUCTO").Value, 5, "I") & Space$(2)
                    a = a & CME(rsDet("DES_PRODUCTO").Value, 53, "I")
                    a = a & CME(Format(rsDet("MTO_SUBTOTAL").Value, "##0.#0"), 12, "D")
                    i = i + 1
                    Printer.Print a
                    rsDet.MoveNext
                Wend
   
   For k = 0 To (11 - i)
       Printer.Print
   Next k
   
   Printer.Print "SON : " & Trim(rsCab("LETRAS").Value) & " NUEVOS SOLES"
   Printer.Print
   Printer.Print Space(69) & "S/." & CME(Format(rsCab("MTO_BASE_IMP").Value, "##0.#0"), 10, "D")
   Printer.Print Space(69) & "S/." & CME(Format(rsCab("MTO_IMPUESTO").Value, "##0.#0"), 10, "D")
   Printer.Print
   Printer.Print Space(69) & "S/." & CME(Format(rsCab("MTO_TOTAL").Value, "##0.#0"), 10, "D")
   Printer.EndDoc
   
   On Error GoTo 0
   
   Exit Sub
   
ErrorImpresora:
    Err.Raise Err.Number, "clsDocumento.NotaCred", Err.Description
End Sub

Public Function DevTextoDocCnv(ByVal Cia As String, _
                        ByVal TipoDocumento As String, _
                        ByVal numDocumento As String, _
                        Optional ByVal flgBeneficiario = 0, _
                        Optional CodLocal As String = "") As String

On Error GoTo CtrlErr

    DevTextoDocCnv = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO.FN_DEV_TEXTO_DOC_CNV", Cia, numDocumento, TipoDocumento, flgBeneficiario, CodLocal)

    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.DevTextoDocCnv", Err.Description
End Function

Public Function ListaDoc(ByVal Cia As String, _
                        ByVal TipoDocumento As String, _
                        ByVal numDocumento As String) As oraDynaset
On Error GoTo CtrlErr
    Set ListaDoc = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_LISTA_DOC", 0, Cia, numDocumento, TipoDocumento)
    
    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.ListaDoc", Err.Description
End Function

Public Function ListaFormato(ByVal Cia As String, _
                        ByVal TipoDocumento As String, _
                        ByVal numDocumento As String) As oraDynaset
On Error GoTo CtrlErr
    Set ListaFormato = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_LISTA_FORMATO", 0, Cia, numDocumento, TipoDocumento)
    
    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.ListaFormato", Err.Description
End Function

Public Function ImprimeRecibo(ByVal Cia As String, _
                        ByVal TipoDocumento As String, _
                        ByVal numDocumento As String) As String
On Error GoTo CtrlErr

    ImprimeRecibo = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO.FN_IMPRIME_RECIBO", Cia, TipoDocumento, numDocumento)
    
    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.ImprimeRecibo", Err.Description
End Function

Public Function ListaTipoDocumento(ByVal codTipoDoc As String, Optional ByVal Tip As String = "0") As oraDynaset
On Error GoTo CtrlErr
    Set ListaTipoDocumento = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_LISTA_TIPO_DOCUMENTO", 0, codTipoDoc, Tip)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaTipoDocumento", Err.Description
End Function

Public Function ListaControlSOAT(ByVal NumFormulario As String, ByVal CodBtl As String) As oraDynaset
On Error GoTo CtrlErr
    Set ListaControlSOAT = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_LISTA_CONTROL_SOAT", 0, NumFormulario, CodBtl)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaControlSOAT", Err.Description
End Function

Public Function DevMensaje(ByVal Cia As String, ByVal numDocumento As String, ByVal TipoDocumento As String, ByVal CodLocal As String) As String
On Error GoTo CtrlErr
    DevMensaje = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO.FN_DEV_MENSAJE", Cia, numDocumento, TipoDocumento, CodLocal)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.DevMensaje", Err.Description
End Function

Public Function DevImpDocCnv(ByVal Cia As String, ByVal numDocumento As String, ByVal TipoDocumento As String, ByVal CodLocal As String) As Double
On Error GoTo CtrlErr
    DevImpDocCnv = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO.FN_DEV_IMP_DOC_CNV", Cia, numDocumento, TipoDocumento, CodLocal)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.DevImpDocCnv", Err.Description
End Function

Sub Imprime_Ajuste_Cje(ByVal vstrNum_Ajuste$)
Dim strSQL$
Dim rstCab As oraDynaset
Dim rstDet As oraDynaset
Dim rstTipo As oraDynaset
Dim rstMotivo As oraDynaset
Dim Copia As Double
Dim strSigno$
Dim strCodigo$
Dim objAjuste As New clsAjuste
Dim gstrSin_Lote As String
Dim gstrSin_Fecha  As String
'Cabecera
    
    gstrSin_Lote = objUsuario.valor(objUsuario.ValParametroSinLote, objUsuario.CodigoEmpresa)
    gstrSin_Fecha = objUsuario.valor(objUsuario.ValParametroSinFecha, objUsuario.CodigoEmpresa)
    
    Set rstCab = objAjuste.ListaCabAjuste(vstrNum_Ajuste, objUsuario.CodigoLocal)
    
    If rstCab.RecordCount = 0 Then
        MsgBox "NO SE UBICO AJUSTE ( EL AJUSTE DEBE ESTAR PROCESADO ) : " & vstrNum_Ajuste, vbCritical, App.ProductName
        Exit Sub
    End If
    
    If rstCab.RecordCount > 1 Then
        MsgBox "SE ENCONTRARON VARIOS REGISTROS CABECERA PARA " & vstrNum_Ajuste, vbCritical, App.ProductName
        Exit Sub
    End If
 
'Detalle
    strSigno = IIf(rstCab("TIP_AJUSTE").Value = "JCA", "-", "")
    
    Set rstDet = objAjuste.ListaDetAjuste(rstCab("NUM_AJUSTE").Value, rstCab("COD_ZONA").Value, strSigno)
    
    If rstDet.RecordCount = 0 Then
        MsgBox "NO SE UBICO DETALLE DE AJUSTE : " & vstrNum_Ajuste, vbCritical, App.ProductName
        Exit Sub
    End If
    
    Set rstTipo = objAjuste.ListaTipoAjuste(rstCab("NUM_AJUSTE").Value, rstCab("COD_ZONA").Value)
    
    Set rstMotivo = objAjuste.ListaMotivo(rstCab("NUM_AJUSTE").Value, rstCab("COD_ZONA").Value)
        
    Dim intLinea%
    Copia = 0
    intLinea = 0
    
       rstDet.MoveFirst
       
       Printer.FontName = "Draft 17cpi"
       While Not rstDet.EOF
            
            '-----------------------------------------------------------------
            ''If Not pfbln_SetForm("003") Then
            ''    Exit Sub
            ''End If
            '-----------------------------------------------------------------
            
            Printer.FontBold = True
            Printer.Print objUsuario.Empresa     ' "BOTICA TORRES DE LIMATAMBO SAC"
            Printer.FontBold = False
            Printer.FontName = "Draft 17cpi"
            Printer.Print Space(105) & "Fecha: " & Format(rstCab("FCH_IMPRESION").Value, "dd/mm/yyyy")
            Printer.Print Space(105) & "Hora : " & Format(rstCab("FCH_IMPRESION").Value, "HH:MM:SS AM/PM ")
            Printer.Print Space(105) & "Página : " & Printer.Page
            Printer.FontBold = True
            Printer.FontName = "Draft 10cpi"
            Printer.Print fstr_Centrar("AJUSTE Nº " & Format(vstrNum_Ajuste, "000-0000000"), 81)
            Printer.Print
            Printer.Print fstr_Centrar(rstCab("DES_MOVIMIENTO").Value, 81)
            Printer.FontName = "Draft 12cpi"
            Printer.Print
            Printer.Print Space(10) & "Sub Almacén    : " & rstCab("DES_ZONA").Value
            Printer.Print
            Printer.Print Space(10) & "Fch. Emisión   : " & Format(rstCab("FCH_EMISION").Value, "DD/MM/YYYY hh:MM:SS AM/PM")
            Printer.Print
            Printer.Print Space(10) & "Usuario        : " & IIf(IsNull(rstCab("USU_EMISOR").Value), " ", rstCab("USU_EMISOR").Value) & " - " & IIf(IsNull(rstCab("NOM_USUARIO").Value), " ", rstCab("NOM_USUARIO").Value)
            Printer.Print
            Printer.Print Space(10) & "Observaciones  : " & IIf(IsNull(rstCab("OBS_AJUSTE").Value), " ", rstCab("OBS_AJUSTE").Value)
            Printer.Print
            Printer.FontBold = False
            Printer.FontName = "Draft 17cpi"
            Printer.Print
            Printer.Print String(136, "-")
            Printer.Print right(Space(3) & "#", 3) & Space(1) & _
                           left("COD." & Space(5), 5) & Space(1) & _
                           left("DESCRIPCIÓN" & Space(53), 53) & Space(1) & _
                           right(Space(4) & "C.F.", 4) & Space(1) & _
                           left("LOTE" & Space(10), 10) & Space(1) & _
                           left("F.VEN." & Space(7), 7) & Space(1) & _
                           left("TIP" & Space(3), 3) & Space(1) & _
                           left("MOT" & Space(3), 3) & Space(1) & _
                           right(Space(5) & "SI.U.", 5) & Space(2) & _
                           right(Space(5) & "SI.F.", 5) & Space(2) & _
                           right(Space(5) & "AJ.U.", 5) & Space(2) & _
                           right(Space(5) & "AJ.F.", 5) & Space(2) & _
                           right(Space(5) & "SF.U.", 5) & Space(2) & _
                           right(Space(5) & "SF.F.", 5) & Space(2)
            Printer.Print String(136, "-")
            '---
            intLinea = 20
            strCodigo = "*"
            While intLinea <= (70 - 7 - rstMotivo.RecordCount - rstTipo.RecordCount) And Not rstDet.EOF
                'Detalle
                If strCodigo <> rstDet("COD_PRODUCTO").Value Then
                    Printer.Print right(Space(3) & rstDet("NUM_ITEM").Value, 3) & Space(1) & _
                                   left(rstDet("COD_PRODUCTO").Value & Space(5), 5) & Space(1) & _
                                   left(rstDet("DES").Value & Space(53), 53) & Space(1) & _
                                   right(Space(4) & rstDet("CF").Value, 4) & Space(1) & _
                                   left(IIf(rstDet("NUM_LOTE").Value = gstrSin_Lote, " ", rstDet("NUM_LOTE").Value) & Space(10), 10) & Space(1) & _
                                   left(IIf(rstDet("FCH_VENCIMIENTO").Value = gstrSin_Fecha, " ", Format(rstDet("FCH_VENCIMIENTO").Value, "MM/YYYY")) & Space(7), 7) & Space(1) & _
                                   left(rstDet("COD_TIPODEV").Value & Space(3), 3) & Space(1) & _
                                   left(rstDet("COD_MOTIVODEV").Value & Space(3), 3) & Space(1) & _
                                   right(Space(5) & rstDet("SI_UND").Value, 5) & Space(2) & _
                                   right(Space(5) & rstDet("SI_FRAC").Value, 5) & Space(2) & _
                                   right(Space(5) & rstDet("CTD_PRODUCTO").Value, 5) & Space(2) & _
                                   right(Space(5) & rstDet("CTD_PRODUCTO_FRAC").Value, 5) & Space(2) & _
                                   right(Space(5) & rstDet("SF_UND").Value, 5) & Space(2) & _
                                   right(Space(5) & rstDet("SF_FRAC").Value, 5) & Space(2)
                Else
                    Printer.Print Space(69) & _
                                   left(IIf(rstDet("NUM_LOTE").Value = gstrSin_Lote, " ", rstDet("NUM_LOTE").Value) & Space(10), 10) & Space(1) & _
                                   left(IIf(rstDet("FCH_VENCIMIENTO").Value = gstrSin_Fecha, " ", Format(rstDet("FCH_VENCIMIENTO").Value, "MM/YYYY")) & Space(7), 7) & Space(1) & _
                                   left(rstDet("COD_TIPODEV").Value & Space(3), 3) & Space(1) & _
                                   left(rstDet("COD_MOTIVODEV").Value & Space(3), 3) & Space(1) & _
                                   right(Space(5) & rstDet("SI_UND").Value, 5) & Space(2) & _
                                   right(Space(5) & rstDet("SI_FRAC").Value, 5) & Space(2) & _
                                   right(Space(5) & rstDet("CTD_PRODUCTO").Value, 5) & Space(2) & _
                                   right(Space(5) & rstDet("CTD_PRODUCTO_FRAC").Value, 5) & Space(2) & _
                                   right(Space(5) & rstDet("SF_UND").Value, 5) & Space(2) & _
                                   right(Space(5) & rstDet("SF_FRAC").Value, 5) & Space(2)
                    
                End If
                strCodigo = rstDet("COD_PRODUCTO").Value
                intLinea = intLinea + 1
                rstDet.MoveNext
            Wend
            Printer.Print String(136, "-")
            Printer.Print
            '++++++++++++++++
            Printer.Print "TIPOS"
            Printer.Print "-----"
            rstTipo.MoveFirst
            While Not rstTipo.EOF
                Printer.Print Space(3) & rstTipo("TIPO").Value
                rstTipo.MoveNext
            Wend
            '++++++++++++++++
            Printer.Print
            Printer.Print "MOTIVOS"
            Printer.Print "-------"
            rstCab.MoveFirst
            Printer.Print Space(3) & rstCab("DES_MOTIVO_AJUSTE").Value
'''''''            rstMotivo.MoveFirst
'''''''            While Not rstMotivo.EOF
'''''''                Printer.Print Space(3) & rstMotivo("MOTIVO").Value
'''''''                rstMotivo.MoveNext
'''''''            Wend
            Printer.NewPage
        Wend
'        Copia = Copia + 1
        Printer.EndDoc
'    Loop
On Error GoTo 0
Exit Sub
 
ErrorImpresora:
    If MsgBox("Existe un Problema con la Impresora" & Chr(13) & _
             "Error: " & CStr(Err.Number) & " - " & Err.Description & Chr(13) & _
             "¿Desea Esperar a ser Resuelto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Resume
   Else
       Exit Sub
   End If
End Sub

Public Property Get AnchoBoleta() As Integer
    AnchoBoleta = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_ANCHO_BOLETA")
End Property

Public Property Get AnchoFactura() As Integer
    AnchoFactura = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_ANCHO_FACTURA")
End Property

''''''Public Sub Proforma(ByVal vstrCia As String, _
''''''                    ByVal vstrCodBtl As String, _
''''''                    ByVal vstrProforma As String, _
''''''                    ByVal vintNumCopia As String)
''''''
''''''   Dim dblSumaSubTotal As Double
''''''   Dim Intlinea As Integer
''''''   Dim intCopia As Integer
''''''   Dim i As Integer
''''''   Dim objImpresiones As New clsProforma
''''''   Dim odynCab As OraDynaset
''''''   Dim odynDet As OraDynaset
''''''   Dim intItems As Integer
''''''   Dim TotItems As Integer
''''''
''''''   Set odynCab = objImpresiones.fnImpCabPed(vstrCia, vstrCodBtl, vstrProforma)
''''''   Set odynDet = objImpresiones.fnImpDetPed(vstrCia, vstrProforma)
''''''
''''''    'Frm_SIS_Impresoras.Show vbModal
''''''    Printer.PaperSize = vbPRPSA4
''''''
''''''    TotItems = "0" & odynCab("ITEM").Value
''''''
''''''    intCopia = 0: TotItems = 0
''''''    If vintNumCopia = 0 Then Exit Sub
''''''
''''''    '-- Generar Copia --'
''''''    Do While vintNumCopia > intCopia
''''''        odynCab.MoveFirst
''''''        odynDet.MoveFirst
''''''        While Not odynDet.EOF
''''''            'Cabecera
''''''
''''''            On Error GoTo Mal_Seteo
''''''            GoTo OK_Seteo
''''''Mal_Seteo:
''''''            MsgBox Err.Description, vbExclamation, "Error de Impresión"
''''''            Exit Sub
''''''OK_Seteo:
''''''            On Error GoTo 0
''''''            On Error GoTo ErrorImpresora
''''''
''''''            Printer.FontName = "Draft 15cpi"
''''''            Printer.FontSize = 9
''''''            Printer.Print Space(155) & "Fecha :" & " " & odynCab("FECHA").Value
''''''            Printer.Print Space(155) & "Hora  :" & " " & odynCab("HORA").Value
''''''            Printer.Print
''''''            Printer.FontName = "Draft 15cpi"
''''''            Printer.Font.Bold = True
''''''            Printer.FontSize = 15
''''''            Printer.Print pfstr_Alineado("PROFORMA", 125, centro)
''''''            Printer.FontName = "Draft 17cpi"
''''''            Printer.FontSize = 9
''''''            Printer.FontBold = False
''''''            Printer.Print
''''''            Printer.Print Space(2) & "Atendido        : " & pfstr_Alineado(odynCab("USUARIO").Value, 50) & " " & "Motorizado:" & " " & pfstr_Alineado(odynCab("MOTORIZADO").Value, 30)
''''''            Printer.Print Space(2) & "Nro.Proforma    : " & pfstr_Alineado(odynCab("NUM_PROFORMA").Value, 50)
''''''            Printer.Print Space(2) & "Codigo Cliente  : " & pfstr_Alineado(odynCab("COD_CLIENTE").Value, 50)
''''''            Printer.Print Space(2) & "Telefono        : " & pfstr_Alineado("xxxxxxxx", 15) & Space(35) & "Tarjeta" & " " & pfstr_Alineado(IIf(IsNull(odynCab("NUM_TARJETA").Value), "", odynCab("NUM_TARJETA").Value), 20) & " " & "F.V" & " " & pfstr_Alineado(IIf(IsNull(odynCab("FCH_VENCIMIENTO").Value), "", odynCab("FCH_VENCIMIENTO").Value), 15)
''''''            Printer.Print Space(2) & "Nombre          : " & pfstr_Alineado(odynCab("NOMBRE").Value, 85) & " " & "D.N.I." & " " & pfstr_Alineado(IIf(IsNull(odynCab("NUM_DOCUMENTO_ID").Value), "", odynCab("NUM_DOCUMENTO_ID").Value), 13)
''''''            Printer.Print Space(2) & "Dirección       : " & pfstr_Alineado(odynCab("DIRECCION").Value, 80)
''''''            Printer.Print Space(2) & "Distrito        : " & pfstr_Alineado(odynCab("DISTRITO").Value, 50) & " " & "T.C" & " " & pfstr_Alineado(odynCab("IMP_TIP_CAMBIO").Value, 8) & Space(15) & "Forma Pago" & " " & pfstr_Alineado(odynCab("DES_HIJO").Value, 15)
''''''            'Printer.Print Space(2) & "Observación     : " & pfstr_Alineado(odynCab_OC("DIR_ENTREGA").Value, 91)
''''''            '-------------------------------'
''''''            Printer.FontSize = 7
''''''            Printer.FontName = "Draft 15cpi"
''''''            Printer.Font.Bold = True
''''''            Printer.FontSize = 11
''''''            Printer.Print
''''''            Printer.Print pfstr_Alineado(odynCab("TIPODOC").Value, 135, centro)
''''''            Printer.FontName = "Draft 17cpi"
''''''            Printer.FontSize = 9
''''''            Printer.FontBold = False
''''''            '-------------------------------'
''''''            Printer.Print String(210, "_")
''''''            Printer.Print pfstr_Alineado("Codigo", 8, centro) & Space(1) & _
''''''                          pfstr_Alineado("Lab", 8, Izquierda) & Space(1) & _
''''''                          pfstr_Alineado("Descripción", 50, Izquierda) & Space(1) & _
''''''                          pfstr_Alineado("Cant.", 8, Derecha) & Space(1) & _
''''''                          pfstr_Alineado("PVP", 8, Derecha) & Space(1) & _
''''''                          pfstr_Alineado("PVPU.", 8, Derecha) & Space(1) & _
''''''                          pfstr_Alineado("Sub Total.", 10, Derecha) & Space(1)
''''''            Printer.Print ; String(210, "_")
''''''          '---
''''''            intItems = 0
''''''            'Printer.FontName = "Draft 17cpi"
''''''            While Not odynDet.EOF
''''''                  'Detalle
''''''                Printer.Print pfstr_Alineado(odynDet("COD_PRODUCTO").Value, 8) & Space(1) & _
''''''                              pfstr_Alineado(odynDet("DES_LABORATORIO").Value, 8) & Space(1) & _
''''''                              pfstr_Alineado(odynDet("DES_PRODUCTO").Value, 50) & Space(1) & _
''''''                              pfstr_Alineado(odynDet("CTD_PRODUCTO").Value, 8, Derecha) & _
''''''                              pfstr_Alineado(odynDet("PRC_UNIT_KAIRO").Value, 8, Derecha) & _
''''''                              pfstr_Alineado(odynDet("IMP_UNIT_BASE_IMP").Value, 8, Derecha) & _
''''''                              pfstr_Alineado(odynDet("MTO_SUBTOTAL").Value, 10, Derecha)
''''''
''''''                dblSumaSubTotal = dblSumaSubTotal + Format(Val(odynDet("MTO_SUBTOTAL").Value), "#,###,##0.00")
''''''                intItems = intItems + 1
''''''                odynDet.MoveNext
''''''            Wend
''''''            'Printer.FontBold = False
''''''            For i = 1 To TotItems
''''''                Printer.Print
''''''            Next i
''''''            'Printer.Print "detalle aqui"
''''''            Printer.Print String(210, "_")
''''''            Printer.Font.name = "Draft 15cpi"
''''''            Printer.FontSize = 8
''''''            Printer.Print
''''''            Printer.Print Space(2) & "Total de Registro => " & " " & intItems & Space(90) & "Total a Cancelar S/. " & " " & Right(Space(15) & Format(dblSumaSubTotal, "#,###,##0.00"), 15)
''''''            Printer.Print
''''''            Printer.Print Space(2) & "Paga con S/. " & " " & Right(Space(15) & Format(odynCab("IMP_CON_REDONDEO").Value, "#,###,##0.00"), 15)
''''''            Printer.Print Space(2) & "Vuelto   s/. " & " " & Right(Space(15) & Format(odynCab("IMP_VUELTO").Value, "#,###,##0.00"), 15)
''''''
''''''            'Printer.NewPage
''''''        Wend
''''''        intCopia = intCopia + 1
''''''        Printer.EndDoc
''''''    Loop
''''''    On Error GoTo 0
''''''    Exit Sub
''''''ErrorImpresora:
''''''    If MsgBox("Existe un Problema con la Impresora" & Chr(13) & _
''''''         "Error: " & CStr(Err.Number) & " - " & Err.Description & Chr(13) & _
''''''         "¿Desea Esperar a ser Resuelto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
''''''        Resume
''''''    Else
''''''      Exit Sub
''''''    End If
''''''End Sub
''''''
''''''

Private Sub Class_Initialize()
    lsCodClaseCOM = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONST_TIP_CLASE_COM")
    lsCodSubClaseCom = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONST_TIP_SUB_CLASE_COM")
    lsCodFamiliaCom = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONST_TIP_FAMILIA_COM")
    lsCodCategoriaCom = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONST_TIP_CATEGORIA_COM")
    lsCodProdTrans = gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONST_COD_PROD_TRANSF")
End Sub

'*********************************************************************************************'
'*********************************************************************************************'
'**** FORMATOS DE TICKETS ****'

Private Function proximoSalto(ByVal vData As String, ByVal vMax As Integer) As Integer
    Dim tmp As String
    Dim i As Integer
    
    tmp = vData
    i = Len(tmp)
    Do While Len(tmp) > vMax
        i = InStrRev(tmp, " ")
        tmp = Mid$(tmp, 1, i - 1)
        
        If Len(tmp) < 42 Then
            Exit Do
        End If
    Loop
    proximoSalto = i
End Function

Private Sub Ticket(ByVal TipoDocumento As String, ByVal numDocumento As String, Optional PagTkt As String, Optional TPagTkt As String, Optional FlgAnulado As Boolean = False, Optional flgProformaDLV As Boolean = False)
    Dim objImpresiones As New clsImpresiones
    Dim objServicio As New clsServicio
    Dim rsServicio As oraDynaset
    Dim rstImpresion As oraDynaset
    Dim rstDocumentoMaquina As oraDynaset
    Dim rstFormaPago As oraDynaset
    Dim rstDonacion As oraDynaset
    Dim rstConvenio As oraDynaset
    'Dim objDocumento As New clsDocumento
    Dim dblDonacion As Double
    Dim strGlosaConv As String
    Dim strMensaje As String
    Dim dblTotal As Double
    Dim i As Integer
    Dim j As Integer
    '''''''''''''''''''''''''''''''''''''''''''''''
    'Dim p As Printer
    Dim intRespuesta As Integer
    Dim intNumeroLineas As Integer
    Dim dblMtoTotalRef As Double
    Dim dblTotUni As Double
    Dim vInd As Integer
    Dim strlinea1 As String
    Dim strlinea2 As String
    Dim strlinea3 As String
    Dim strlinea4 As String
    Dim strlinea5 As String
    Dim strlinea6 As String
    Dim strlinea7 As String
    'Dim dblPagoEfect As Double
    Dim strDireccFiscal As String
    'Dim dblVuelto As Double
    '**************
    Dim sNetUse As String, vInd2 As Integer
    '**************
    Dim oFP As New clsFarmaPuntos
    Dim oTarjetaPuntos As clsTarjetaBean
    Dim rstClienteF As oraDynaset
    'Dim oClienteF As New clsClienteD
    Dim strExperto As String
    Dim dblPuntosDinero As Double, dbl12mahorro As Double
    Dim MontoAhorro As Double, MtoAhorroCampana As Double, MtoAhorroPuntos As Double, CtdCupones As Double
    '**************
    vInd = 0
    MontoAhorro = 0
    MtoAhorroCampana = 0
    MtoAhorroPuntos = 0
    CtdCupones = 0
    
'''    sNetUse = "NET USE LPT1 /DELETE"
'''    vInd2 = Shell(sNetUse)

'''    For Each p In Printers
'''        Debug.Print p.PORT & "/" & p.Devicename & "/" & p.DriverName
'''
'''        If UCase(p.PORT) = Me.PORT(objUsuario.NombrePC, TipoDocumento) Then '"LPT1:" Then
'''            Set Printer = p
'''            vInd = "1"
'''            Exit For
'''        End If
'''    Next p

    'If vInd <> 1 Then MsgBox "No Tiene Conectada la Ticketera al Computador, " & Chr(13) & "Revise la Conexión Física", vbCritical, "": Exit Sub
    
    Set rstConvenio = ListaDoc(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
        
    intNumeroLineas = 0
    dblDonacion = 0
    Set rstDocumentoMaquina = ListaDocumentoMaquina(objUsuario.CodigoEmpresa, TipoDocumento, objUsuario.NombrePC)
    Set rstImpresion = Impresion_Ticket(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
    Set rstFormaPago = Lista_FPagos_x_Ticket(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
    
    If Not IsNull(rstImpresion("NUM_TARJ_PUNTOS").Value) And rstImpresion("EST_TRX_ORBIS").Value <> "D" Then
        Set oTarjetaPuntos = oFP.ValidarTarjetaAsociada(rstImpresion("NUM_TARJ_PUNTOS").Value, objUsuario.Codigo)
        Set rstClienteF = objVenta.fnListaAfiliadoMonedero(oTarjetaPuntos.TipoDNI, oTarjetaPuntos.DNI)
    End If
    
    If Not rstDocumentoMaquina.EOF Then intNumeroLineas = rstDocumentoMaquina("NUM_LINEAS_DOC").Value
    
    On Error GoTo 0
        On Error GoTo ErrorImpresora
            On Error GoTo Mal_Seteo
            GoTo OK_Seteo
Mal_Seteo:
            Close #1
            MsgBox Err.Description, vbExclamation, "Error de Ticketera.....????"
            Exit Sub
OK_Seteo:

            '***********************************************************************'
            dblTotUni = 0
            Open "lpt1" For Output As #1
            'Open "XPSPort" For Output As #1
            
            '--- EXPERTO DEL AHORRO ----
            '****************************************
            If Not IsNull(rstImpresion("NUM_TARJ_PUNTOS").Value) And rstImpresion("EST_TRX_ORBIS").Value <> "D" Then
                strExperto = "Expert" & IIf("" & rstClienteF("FLG_SEXO").Value = "1", "o", "a") & " del Ahorro"
                dblPuntosDinero = Val(oTarjetaPuntos.AhorroTotal) * Val(rstImpresion("FARMA_PUNTOS_REDIME").Value)
                dbl12mahorro = Val("" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "INDFV474"))
                Print #1, Chr$(27); Chr$(58); '16CPI
                Print #1, pfstr_Alineado(UCase("" & rstClienteF("DES_NOM_CLIENTE")), 23, centro)
                Print #1, pfstr_Alineado(strExperto, 23, centro)
                Print #1, Chr$(27); Chr$(77); '12CPI
                strExperto = "Ahorraste S/. " & Format$(dblPuntosDinero, "###,##0.00") & " en los ultimos 12 meses"
                If dblPuntosDinero >= dbl12mahorro Then
                    i = Len(strExperto)
                    Do While strExperto <> ""
                        i = proximoSalto(strExperto, 42)
                        Print #1, pfstr_Alineado(Trim(Mid$(strExperto, 1, i)), 42, centro)
                        strExperto = Trim(Mid$(strExperto, i + 1))
                    Loop
                End If
                Print #1, ""
            End If
            '****************************************
            '--- FIN EXPERTO DEL AHORRO ----
            
            '**** 16-DOT PITCH FONT *******'
            Print #1, Chr$(27); Chr$(58);
            
            'If objVenta.EsMFA(objUsuario.CodigoLocal) Then
                'Print #1, "  BOTICAS MIFARMA  "
            'Else
                Print #1, NombreComercial(objUsuario.CodigoLocal) '"    BOTICAS BTL    "
            'End If
            '****** Standar 12CPI ********'
            Print #1, Chr$(27); Chr$(77);
            '-----------------------------'
            'Print #1, "   BOTICA TORRES DE LIMATAMBO SAC    "
            
            'Print #1, "    " & Replace(UCase(objUsuario.Empresa), "Í", "I")'comentado por arturo escate 19/03/2009
            Print #1, pfstr_Alineado(Replace(UCase(objUsuario.Empresa), "Í", "I"), 42, centro)
            
            'Print #1, Space(1) & objUsuario.Direccion 'comentado por arturo escate 19/03/2009
            
            strDireccFiscal = Replace(objUsuario.Direccion, " - ", "-")
            While strDireccFiscal <> ""
                Print #1, pfstr_Alineado(Mid(strDireccFiscal, 1, 42), 42, centro)
                If Len(strDireccFiscal) > 42 Then
                    strDireccFiscal = Mid(strDireccFiscal, 43, Len(strDireccFiscal) - 42)
                Else
                    strDireccFiscal = ""
                End If
            Wend
             
            '*******************************************************
            
            'Print #1, Space(8) & "RUC :" & objUsuario.Ruc 'comentado por arturo escate 19/03/2009
            Print #1, pfstr_Alineado("RUC :" & objUsuario.Ruc, 42, centro)
            
            'Print #1, "BTL" & objUsuario.CodigoLocal & Space(2) & gclsOracle.FN_Valor("CMR.PKG_LOCAL.FN_DIR", objUsuario.CodigoLocal)
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            'PREGUNTAR PORQUE ESTA ESTO AQUI, SI ABAJO IMPRIME LOS DATOS DE LA RAZON SOCIAL Y EL RUC
''''            If rstImpresion("COD_TIPO_DOCUMENTO").Value = "TKB" Then
''''               Print #1, left(IIf(IsNull(rstImpresion("DES_RAZON_SOCIAL").Value), "", rstImpresion("DES_RAZON_SOCIAL").Value) & Space(25), 25)
''''            End If''ESTO SE COMENTO PORQUE ABAJO SE VA PONER LOS DATOS DEL CLIENTE
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

            '******************************************************
            'pherrera 08/03/10 cambio para que la direccion se imprima completa.
'            Dim strDireccion As String
'            Dim bytLineas As Single
'            Dim strDirPrint As String
'            strDireccion = objUsuario.Direccion
'            bytLineas = Len(strDireccion) / 42
'            bytLineas = CInt(bytLineas) - (bytLineas <> CInt(bytLineas))
'
'            For i = 0 To bytLineas - 1
'                strDirPrint = Mid(strDireccion, (42 * i) + 1, 42)
'                Print #1, pfstr_Alineado(Replace(strDirPrint, " - ", "-"), 42, centro)
'            Next
            '*******************************************************

            Print #1, "BTL" & objUsuario.CodigoLocal & Space(2) & CadenaDirecc_Part1(gclsOracle.FN_Valor("CMR.PKG_LOCAL.FN_DIR", objUsuario.CodigoLocal))
            
            Print #1, Space(8) & CadenaDirecc_Part2(gclsOracle.FN_Valor("CMR.PKG_LOCAL.FN_DIR", objUsuario.CodigoLocal))
            Print #1, "FECHA:" & rstImpresion("FCH_REGISTRA").Value & Space(1) & "CAJA:" & " " & rstImpresion("COD_MAQUINA").Value
            Print #1, "Ticket" & " " & "Nº" & Space(1) & rstImpresion("NUM_DOCUMENTO").Value & " " & "Serie:" & " " & rstImpresion("COD_SERIE_REL").Value
            Print #1, "Cajero:" & Space(1) & rstImpresion("COD_USUARIO").Value & Space(1) & rstImpresion("NOMB_CAJERO").Value
            'Print #1, ""
            Print #1, left("Codigo" & Space(5), 5) & Space(1) & _
                      left("Descripcion" & Space(20), 20) & Space(1) & _
                      right(Space(5) & "Cant.", 5) & Space(1) & _
                      right(Space(8) & "Importe", 8) & Space(1)
            Print #1, String(42, "-")
            Printer.FontBold = False
            
            While Not rstImpresion.EOF
            
                'Detalle
                If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio And rstConvenio("FLG_DOC_COPAGO").Value = "1" Then
                    If Len(rstImpresion("DES_PRODUCTO").Value) > 20 Then ' SI LA DESCRIPCION DE PRODUCTO ES MAYOR A 20 CARACTERES'
                        Print #1, left(rstImpresion("COD_PRODUCTO").Value & Space(5), 5) & Space(1) & _
                                  left(Mid(rstImpresion("DES_PRODUCTO").Value, 1, 20) & Space(20), 20) & Space(1) & _
                                  right(Space(5) & rstImpresion("CANTIDAD").Value, 5) & Space(2) & _
                                  pfstr_Alineado(Format(rstImpresion("PRC_VTA_TOTAL").Value, "#0.00"), 7, Derecha) & Space(1)
                                  'right(Space(7) & rstImpresion("SUB_TOTAL").Value, 7) & Space(1)
                                                
                        Print #1, Space(6) & left(Mid(rstImpresion("DES_PRODUCTO").Value, 21, 40) & Space(20), 20) & Space(3)
                                
                        'mlaguna: Agrega numero de partida a inafectos 19/04/2010
                        If rstImpresion("PCT_IGV").Value = 0 Then
                            Print #1, Space(6) & "PA: " & left(rstImpresion("NUM_PART_ARANCEL").Value, 20)
                        End If
                    Else
                        Print #1, left(rstImpresion("COD_PRODUCTO").Value & Space(5), 5) & Space(1) & _
                                  left(Mid(rstImpresion("DES_PRODUCTO").Value, 1, 30) & Space(30), 30) & Space(2) & _
                                  Space(30) & right(Space(5) & rstImpresion("CANTIDAD").Value, 5) & Space(1) & _
                                  pfstr_Alineado(Format(rstImpresion("PRC_VTA_TOTAL").Value, "#0.00"), 7, Derecha) & Space(1)
                                           
                        'mlaguna: Agrega numero de partida a inafectos 19/04/2010
                        If rstImpresion("PCT_IGV").Value = 0 Then
                            Print #1, Space(6) & "PA: " & left(rstImpresion("NUM_PART_ARANCEL").Value, 20)
                        End If
                    End If
                Else
                    If Len(rstImpresion("DES_PRODUCTO").Value) > 20 Then ' SI LA DESCRIPCION DE PRODUCTO ES MAYOR A 20 CARACTERES'
                        Print #1, left(rstImpresion("COD_PRODUCTO").Value & Space(5), 5) & Space(1) & _
                                  left(Mid(rstImpresion("DES_PRODUCTO").Value, 1, 20) & Space(20), 20) & Space(1) & _
                                  right(Space(5) & rstImpresion("CANTIDAD").Value, 5) & Space(2) & _
                                  pfstr_Alineado(Format$(Val(rstImpresion("SUB_TOTAL").Value) + Val(rstImpresion("MTO_AHORRO").Value), "#0.00"), 7, Derecha) & Space(1)
                                  'pfstr_Alineado(rstImpresion("SUB_TOTAL").Value, 7, Derecha) & Space(1)
                                  'right(Space(7) & rstImpresion("SUB_TOTAL").Value, 7) & Space(1)
                                                
                        Print #1, Space(6) & left(Mid(rstImpresion("DES_PRODUCTO").Value, 21, 40) & Space(20), 20) & Space(3)
                                
                        'mlaguna: Agrega numero de partida a inafectos 19/04/2010
                        If rstImpresion("PCT_IGV").Value = 0 Then
                            Print #1, Space(6) & "PA: " & left(rstImpresion("NUM_PART_ARANCEL").Value, 20)
                        End If
                    Else
                        Print #1, left(rstImpresion("COD_PRODUCTO").Value & Space(5), 5) & Space(1) & _
                                  left(Mid(rstImpresion("DES_PRODUCTO").Value, 1, 30) & Space(30), 30) & Space(2) & _
                                  Space(30) & right(Space(5) & rstImpresion("CANTIDAD").Value, 5) & Space(1) & _
                                  pfstr_Alineado(Format$(Val(rstImpresion("SUB_TOTAL").Value) + Val(rstImpresion("MTO_AHORRO").Value), "#0.00"), 7, Derecha) & Space(1)
                                  'pfstr_Alineado(rstImpresion("SUB_TOTAL").Value, 7, Derecha) & Space(1)
                                           
                        'mlaguna: Agrega numero de partida a inafectos 19/04/2010
                        If rstImpresion("PCT_IGV").Value = 0 Then
                            Print #1, Space(6) & "PA: " & left(rstImpresion("NUM_PART_ARANCEL").Value, 20)
                        End If
                    End If
                End If
                
                Dim strLote, strFecVencimiento As String
                strLote = Trim("" & rstImpresion("NUM_LOTE").Value)
                strFecVencimiento = Trim("" & rstImpresion("FCH_VENCIMIENTO").Value)
                If Not (strLote & strFecVencimiento) = "" Then
                    Print #1, Space(6) & left(IIf(strLote = "", "", "Lote:" & strLote) & Space(25), 25) & Space(3)
                    Print #1, Space(6) & left(IIf(strFecVencimiento = "", "", "F. Venc.: " & strFecVencimiento) & Space(25), 25) & Space(3)
                End If
                
                MontoAhorro = MontoAhorro + Val(rstImpresion("MTO_AHORRO").Value)
                MtoAhorroPuntos = MtoAhorroPuntos + Val(rstImpresion("AHORRO_PUNTOS").Value)
                MtoAhorroCampana = MtoAhorroCampana + Val(rstImpresion("AHORRO_CAMP").Value)
                CtdCupones = CtdCupones + Val(rstImpresion("CTD_CUPON").Value)
                rstImpresion.MoveNext
            Wend
            
            dblTotUni = rstImpresion.RecordCount
            rstImpresion.MoveFirst
            'Print #1, ""
            'Print #1, Space(5) & "Total Productos" & " " & Val(dblTotUni)'comentado por arturo escate 19/03/2009
            'Print #1, pfstr_Alineado("Total Productos" & " " & Val(dblTotUni), 42, Derecha)
            'Print #1, pfstr_Alineado(String(22, "-"), 42, Derecha)
            
            Print #1, pfstr_Alineado("Item(s):" & Val(dblTotUni), 42, Izquierda)
            
            '--- EXPERTO DEL AHORRO ----
            '****************************************
            If Not IsNull(rstImpresion("NUM_TARJ_PUNTOS").Value) And rstImpresion("EST_TRX_ORBIS").Value <> "D" Then
                If MtoAhorroCampana > 0 Then
                    Print #1, pfstr_Alineado(String(22, "-"), 42, Derecha)
                    Print #1, left("AHORRO DCTO PTOS INMEDIATOS: S/. " & Space(34), 34) & _
                              pfstr_Alineado(Format$(MtoAhorroCampana * -1, "#0.00"), 7, Derecha) & Space(1)
                End If
                If MtoAhorroPuntos > 0 Then
                    If MtoAhorroCampana = 0 Then Print #1, pfstr_Alineado(String(22, "-"), 42, Derecha)
                    Print #1, left("AHORRO X DCTO 1000 PTOS: S/. " & Space(34), 34) & _
                              pfstr_Alineado(Format$(MtoAhorroPuntos * -1, "#0.00"), 7, Derecha) & Space(1)
                End If
            End If
            '****************************************
            '--- FIN EXPERTO DEL AHORRO ----
            
            If rstImpresion("COD_TIPO_DOCUMENTO").Value = "TKF" Then
                If rstImpresion("MTO_EXONERADO").Value = 0 Then
                    Print #1, pfstr_Alineado(left("Valor Venta" & Space(11), 11) & Space(7) & "S/." & right(Space(10) & Format(rstImpresion("MTO_BASE_IMP").Value, "#0.00"), 10) & Space(1), 42, Derecha)
                Else
                    Print #1, pfstr_Alineado(left("Valor venta" & Space(11), 11) & Space(7) & "S/." & right(Space(10) & Format(rstImpresion("MTO_EXONERADO").Value, "#0.00"), 10) & Space(1), 42, Derecha)
                End If
                Print #1, pfstr_Alineado(left("IGV " & objUsuario.IGV & "%." & Space(11), 11) & Space(7) & "S/." & right(Space(10) & Format(rstImpresion("MTO_IMPUESTO").Value, "#0.00"), 10) & Space(1), 42, Derecha)
            End If
            Print #1, pfstr_Alineado(String(22, "-"), 42, Derecha)
            
            ''''''' NUEVO 27/04/2011
            Dim XRedondeo As Double
            
            While Not rstFormaPago.EOF
                If rstFormaPago("COD_FORMA_PAGO") = "008" Then
                    XRedondeo = XRedondeo + rstFormaPago("IMP_MONEDA_NAC").Value
                End If
                rstFormaPago.MoveNext
            Wend
            rstFormaPago.MoveFirst
            ''''''' NUEVO 27/04/2011
            'Print #1, pfstr_Alineado("Redondeo:" & Format(Val(XRedondeo), "#0.00") & Space(10) & left("Total" & Space(6), 6) & Space(2) & "S/." & right(Space(1) & Format(Val(rstImpresion("MTO_TOTAL").Value) + Val(XRedondeo), "#0.00"), 10) & Space(1), 42, Derecha)
            'Print #1, pfstr_Alineado(Space(10) & left("Total" & Space(6), 6) & Space(2) & "S/." & right(Space(1) & Format(Val(rstImpresion("MTO_TOTAL").Value), "#0.00"), 10) & Space(1), 42, Derecha)
            
            If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio And rstConvenio("FLG_DOC_COPAGO").Value = "1" Then
                Print #1, pfstr_Alineado(Space(10) & left("Total" & Space(6), 6) & Space(2) & "S/." & right(Space(1) & Format(Val(rstImpresion("MTO_TOTAL").Value) + DevImpDocCnv(objUsuario.CodigoEmpresa, numDocumento, TipoDocumento, objUsuario.CodigoLocal), "#0.00"), 10) & Space(1), 42, Derecha)
                Print #1, "COPAGO - " & " (" & rstConvenio("PCT_BENEFICIARIO_ORIG").Value & "%)              S/.   " & " " & Format(Val(rstImpresion("MTO_TOTAL").Value), "#0.00")
            Else
                Print #1, pfstr_Alineado(Space(10) & left("Total" & Space(6), 6) & Space(2) & "S/." & right(Space(1) & Format(Val(rstImpresion("MTO_TOTAL").Value), "#0.00"), 10) & Space(1), 42, Derecha)
            End If
            
            
            'Print #1, ""
            ''Aca imprime las formas de pago
            Dim dblMonNac As Double
            Dim dblVuelto As Double
            'Dim montoRedoneot As Double
            dblMonNac = 0: dblVuelto = 0
            rstFormaPago.MoveFirst
            While Not rstFormaPago.EOF
                Print #1, left(rstFormaPago("DES_HIJO") & Space(25), 25) & Space(4) & _
                          "S/." & " " & right(Space(8) & Format(rstFormaPago("IMP_MONEDA_NAC").Value, "#0.00"), 8)
                'If rstFormaPago("COD_FORMA_PAGO") = "008" Then
                          'montoRedoneot = Val(montoRedoneot + "" & rstFormaPago("IMP_MONEDA_NAC").Value)
                'End If
                dblMonNac = dblMonNac + Val(rstFormaPago("IMP_MONEDA_NAC").Value)
                dblVuelto = dblVuelto + Val(rstFormaPago("IMP_VUELTO").Value)
                rstFormaPago.MoveNext
            Wend
            Print #1, pfstr_Alineado(String(22, "-"), 42, Derecha)
            Print #1, "Total a pagar                S/.   " & " " & Format(Val(rstImpresion("MTO_TOTAL").Value) + XRedondeo, "#0.00")
            Print #1, "Vuelto                       S/.   " & " " & Format(dblVuelto, "#0.00")
            
            '& Space(10) & "" & "Total", 26, Derecha) & Space(1) & _
                      right(Space(5) & "", 5) & Space(2) & _
                      pfstr_Alineado(Format(rstImpresion("MTO_TOTAL").Value, "#0.00"), 7, Derecha) & Space(1)


''''HASTA ACA TODO ALINEADO
            
            
            
            Set rsServicio = objServicio.ListaServiciosAsoc(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
            While Not rsServicio.EOF
                If objVenta.EsNavsat("" & rsServicio("COD_TIPO_SERVICIO")) = "1" Then
                    Print #1, "PIN o Telefono : " & rsServicio("NUM_SUMINISTRO")
                    Print #1, "Nº Aprobacion:" & rsServicio("NUM_RECIBO")
                    If Not IsNull(rsServicio("DES_LEYENDA")) Then
                        If InStr(rsServicio("DES_LEYENDA"), "1[") > 0 Then
                            strlinea1 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "1[") + 2), InStr(rsServicio("DES_LEYENDA"), "]1") - ((InStr(rsServicio("DES_LEYENDA"), "1[") + 2)))
                            If strlinea1 <> "" Then Print #1, strlinea1
                        End If
                        If InStr(rsServicio("DES_LEYENDA"), "2[") > 0 Then
                            strlinea2 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "2[") + 2), InStr(rsServicio("DES_LEYENDA"), "]2") - ((InStr(rsServicio("DES_LEYENDA"), "2[") + 2)))
                            If strlinea2 <> "" Then Print #1, strlinea2
                        End If
                        If InStr(rsServicio("DES_LEYENDA"), "3[") > 0 Then
                            strlinea3 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "3[") + 2), InStr(rsServicio("DES_LEYENDA"), "]3") - ((InStr(rsServicio("DES_LEYENDA"), "3[") + 2)))
                            If strlinea3 <> "" Then Print #1, strlinea3
                        End If
                        If InStr(rsServicio("DES_LEYENDA"), "4[") > 0 Then
                            strlinea4 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "4[") + 2), InStr(rsServicio("DES_LEYENDA"), "]4") - ((InStr(rsServicio("DES_LEYENDA"), "4[") + 2)))
                            If strlinea4 <> "" Then Print #1, strlinea4
                        End If
                        If InStr(rsServicio("DES_LEYENDA"), "5[") > 0 Then
                            strlinea5 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "5[") + 2), InStr(rsServicio("DES_LEYENDA"), "]5") - ((InStr(rsServicio("DES_LEYENDA"), "5[") + 2)))
                            If strlinea5 <> "" Then Print #1, strlinea5
                        End If
                        If InStr(rsServicio("DES_LEYENDA"), "6[") > 0 Then
                            strlinea6 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "6[") + 2), InStr(rsServicio("DES_LEYENDA"), "]6") - ((InStr(rsServicio("DES_LEYENDA"), "6[") + 2)))
                            If strlinea6 <> "" Then Print #1, strlinea6
                        End If
                        If InStr(rsServicio("DES_LEYENDA"), "7[") > 0 Then
                            strlinea7 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "7[") + 2), InStr(rsServicio("DES_LEYENDA"), "]7") - ((InStr(rsServicio("DES_LEYENDA"), "7[") + 2)))
                            If strlinea7 <> "" Then Print #1, strlinea7
                        End If
                    End If
                End If
                rsServicio.MoveNext
            Wend
            Set objServicio = Nothing
            
'''''            dblTotUni = rstImpresion.RecordCount
'''''            Print #1, ""
'''''            'Print #1, Space(5) & "Total Productos" & " " & Val(dblTotUni)'comentado por arturo escate 19/03/2009
'''''            Print #1, pfstr_Alineado("Total Productos" & " " & Val(dblTotUni), 42, Derecha)
'''''
'''''            Print #1, ""
          
'            dblPagoEfect = DevPagoEfectVuelto_Ticket(objUsuario.CodigoEmpresa, _
'                                                     TipoDocumento, _
'                                                     NumDocumento, _
'                                                     "1")
'
'            dblVuelto = DevPagoEfectVuelto_Ticket(objUsuario.CodigoEmpresa, _
'                                                  TipoDocumento, _
'                                                  NumDocumento, _
'                                                  "2")
            
            
''''''''''            If rstImpresion("COD_TIPO_DOCUMENTO").Value = "TKF" Then
''''''''''               If rstImpresion("MTO_EXONERADO").Value = 0 Then
''''''''''                  'Print #1, Space(15) & "Valor venta   S/. " & " " & Format(rstImpresion("MTO_BASE_IMP").Value, "#0.00") 'comentado por arturo escate 19/03/2009
''''''''''                   'Print #1, pfstr_Alineado("Valor venta   S/. " & " " & Format(rstImpresion("MTO_BASE_IMP").Value, "#0.00"), 42, Derecha)
''''''''''                   Print #1, pfstr_Alineado(left("Valor Venta" & Space(11), 11) & Space(7) & "S/." & right(Space(10) & Format(rstImpresion("MTO_BASE_IMP").Value, "#0.00"), 10) & Space(1), 42, Derecha)
''''''''''               Else
''''''''''                  'Print #1, Space(15) & "Valor venta   S/. " & " " & Format(rstImpresion("MTO_EXONERADO").Value, "#0.00") 'comentado por arturo escate 19/03/2009
''''''''''                  'Print #1, pfstr_Alineado("Valor venta   S/. " & " " & Format(rstImpresion("MTO_EXONERADO").Value, "#0.00"), 42, Derecha)
''''''''''                  Print #1, pfstr_Alineado(left("Valor venta" & Space(11), 11) & Space(7) & "S/." & right(Space(10) & Format(rstImpresion("MTO_EXONERADO").Value, "#0.00"), 10) & Space(1), 42, Derecha)
''''''''''               End If
''''''''''                  'Print #1, Space(15) & "IGV " & objUsuario.IGV & "%.      S/. " & " " & Format(rstImpresion("MTO_IMPUESTO").Value, "#0.00")'comentado por arturo escate 19/03/2009
''''''''''                  'Print #1, pfstr_Alineado("IGV " & objUsuario.IGV & "%.      S/. " & " " & Format(rstImpresion("MTO_IMPUESTO").Value, "#0.00"), 42, Derecha)
''''''''''                  Print #1, pfstr_Alineado(left("IGV " & objUsuario.IGV & "%." & Space(11), 11) & Space(7) & "S/." & right(Space(10) & Format(rstImpresion("MTO_IMPUESTO").Value, "#0.00"), 10) & Space(1), 42, Derecha)
''''''''''            End If
                  'Print #1, Space(15) & "Total         S/. " & " " & Format(rstImpresion("MTO_TOTAL").Value, "#0.00") 'comentado por arturo escate 19/03/2009
                  'Print #1, pfstr_Alineado("Total         S/. " & " " & Format(rstImpresion("MTO_TOTAL").Value, "#0.00"), 42, Derecha)
                  'Print #1, pfstr_Alineado(left("Total" & Space(11), 11) & Space(7) & "S/." & right(Space(10) & Format(rstImpresion("MTO_TOTAL").Value, "#0.00"), 10) & Space(1), 42, Derecha)
                  'Print #1, Space(15) & "Total         S/. " & " " & Format(rstCAB_TiketFac("MTO_TOTAL").Value, "#0.00")
                  'Print #1, ""
                  
                  'Print #1, "Pago Efectivo                S/. " & " " & Format(dblPagoEfect, "#0.00")
                  'Print #1, "Vuelto                       S/. " & " " & Format(dblVuelto, "#0.00")

                    
            If Round(MontoAhorro, 2) > 0 Then
                If Not DevImpDocCnv(objUsuario.CodigoEmpresa, numDocumento, TipoDocumento, objUsuario.CodigoLocal) = MontoAhorro Then
                    ''Print #1, ""
                    'Print #1, String(42, "-")
                    'Print #1, "USTED HA AHORRADO           S/.   " & " " & Format(MontoAhorro, "#0.00")
                    ''Print #1, ""
                    'Print #1, String(42, "-")
                    Print #1, ""
                    Print #1, pfstr_Alineado("!!!FELICITACIONES!!!", 42, centro)
                    Print #1, pfstr_Alineado("EN ESTA COMPRA AHORRASTE S/. " & Format(MontoAhorro, "#0.00"), 42, centro)
                    Print #1, ""
                End If
            End If
                  
            If Not rstImpresion("DES_RAZON_SOCIAL").Value = "" Then
                If rstImpresion("COD_TIPO_DOCUMENTO").Value = "TKF" Then
                    Print #1, "Ruc Ciente   :" & " " & rstImpresion("COD_CLIENTE").Value
                    Print #1, "Razon Social:" & " " & rstImpresion("DES_RAZON_SOCIAL").Value
                Else
                    If Len("" & rstImpresion("COD_CLIENTE").Value) = 11 Then
                        Print #1, "Ruc Ciente   :" & " " & rstImpresion("COD_CLIENTE").Value
                        Print #1, "Razon Social:" & " " & rstImpresion("DES_RAZON_SOCIAL").Value
                    Else
                        Print #1, "Doc. Ident :" & " " & rstImpresion("COD_CLIENTE").Value
                        Print #1, "Nombre     :" & " " & rstImpresion("DES_RAZON_SOCIAL").Value
                    End If
                    
                End If
            End If
            
            
            If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio Then
                Dim objConvenio As New clsConvenio
                Dim objCliente  As New clsCliente
                If objConvenio.ListaConvenioCompetencia(rstConvenio("COD_CONVENIO").Value) = 0 Then
                    dblMtoTotalRef = 0
                    dblMtoTotalRef = DevImpDocCnv(objUsuario.CodigoEmpresa, numDocumento, TipoDocumento, objUsuario.CodigoLocal)
                
                    Print #1, ""
                    Print #1, "Convenio: " & rstImpresion("DES_CONVENIO").Value
                    '''''''''''''''''''''''''''''''
                    
                    If rstConvenio("FLG_DOC_COPAGO").Value = "1" Then
                        Print #1, "B E N E F I C I A R I O  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO_ORIG").Value & "%)"
                        objVenta.Beneficiario = strGlosaConv
                        If objConvenio.FnDev_Deducible(rstConvenio("COD_CONVENIO").Value) = "0" Then
                            Print #1, rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value & " " & rstConvenio("NUM_DOCUMENTO_REFER").Value & " (" & 100 - rstConvenio("PCT_BENEFICIARIO_ORIG").Value & "%)" & IIf("" & rstConvenio("FLG_IMPRIME_COPAGO").Value = "1", " - S/.: " & Format(dblMtoTotalRef, "###,##0.00"), "") & "  "
                        End If
                    Else
                        Set objCliente = New clsCliente
                        objVenta.Beneficiario = objCliente.Nombre(IIf(IsNull(rstConvenio("COD_CLIENTE").Value), "", rstConvenio("COD_CLIENTE").Value))
                        Set objCliente = Nothing
                        If dblMtoTotalRef > -1 And rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value <> "" Then
                            Print #1, "#REF " & rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value & " " & rstConvenio("NUM_DOCUMENTO_REFER").Value & " (" & 100 - rstConvenio("PCT_BENEFICIARIO").Value & "%) - S/.: " & Format(dblMtoTotalRef, "###,##0.00") & Space(2)
                            Print #1, "Beneficiario:" & Space(2) & objVenta.Beneficiario
                        Else
                            Print #1, "Beneficiario:" & Space(2) & objVenta.Beneficiario
                        End If
                        Print #1, "I N S T I T U C I O N  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%)"
                    End If
    
                    Set objConvenio = Nothing
                End If
            End If
            
            If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio Then
                Dim strMensajeConvenio As String
                If objConvenio.ListaConvenioCompetencia(rstConvenio("COD_CONVENIO").Value) = 0 Then
                    strMensajeConvenio = "" & Me.MensajeImpresionConvenio(rstConvenio("COD_CONVENIO").Value)
                    If Not strMensajeConvenio = "" Then
                        Print #1, strMensajeConvenio
                    End If
                    ''11/03/2010 SE AGREGO PARA QUE EL CONVENIO BTL EN EFECTIVO TAMBIEN SALGA LA FIRMA
                    If rstImpresion("COD_TIPO_VTAFIN").Value = "2" Or gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_CNV_BTL_EFECTIVO") = rstConvenio("COD_CONVENIO").Value Then
                        Print #1, ""
                        'Print #1, ""
                        'Print #1, ""
                        'Print #1, "    " & String(32, "-") & "    "
                        Print #1, "Nombre: " & rstImpresion("DES_AUX_CLI_NOMBRE").Value
                        Print #1, "DNI: " & rstConvenio("DNI_BENEFICIARIO").Value
                    End If
                End If
            End If
               
            '--- EXPERTO DEL AHORRO ----
            ' *****************************************
            If Not IsNull(rstImpresion("NUM_TARJ_PUNTOS").Value) And rstImpresion("EST_TRX_ORBIS").Value <> "D" Then
                Print #1, String(42, "*")
                Print #1, Chr$(27); Chr$(58); '16CPI
                Print #1, pfstr_Alineado("MONEDERO MIFARMA", 23, centro)
                
                Print #1, Chr$(27); Chr$(77); '12CPI
    
                Print #1, pfstr_Alineado(oTarjetaPuntos.NombreCompleto, 42, centro)
                i = 42 - (Len(EncriptarNumeroTarjeta(oTarjetaPuntos.NumeroTarjeta)) + Len(oTarjetaPuntos.DNI) + 17)
                strExperto = "Tarjeta: " & EncriptarNumeroTarjeta(oTarjetaPuntos.NumeroTarjeta) & Space(i) & _
                             "DNI/CE: " & oTarjetaPuntos.DNI
                Print #1, strExperto
                If MontoAhorro > 0.5 Then _
                    Print #1, "AHORRASTE              S/.: " & pfstr_Alineado(Format(MontoAhorro, "#0.00"), 14, Derecha)
                If Val(rstImpresion("PT_REDIMIDO").Value) > 0 Then _
                    Print #1, "HAS CANJEADO (PTOS DINERO): " & pfstr_Alineado(Format(Val(rstImpresion("PT_REDIMIDO").Value), "#0"), 14, Derecha)
                If CtdCupones > 0 Then _
                    Print #1, "HAS GANADO   (CUPON DSCTO): " & pfstr_Alineado(Format(CtdCupones, "#0"), 14, Derecha)
                If Val(rstImpresion("PT_ACUMULADO").Value) > 0 Then _
                    Print #1, "HAS GANADO   (PTOS DINERO): " & pfstr_Alineado(Format(Val(rstImpresion("PT_ACUMULADO").Value), "#0"), 14, Derecha)
                If Val(rstImpresion("PT_TOTAL").Value) > 0 Then _
                    Print #1, "ACUMULADOS   (PTOS DINERO): " & pfstr_Alineado(Format(Val(rstImpresion("PT_TOTAL").Value), "#0"), 14, Derecha)
                If dblPuntosDinero >= dbl12mahorro Then _
                    Print #1, "EN 12 MESES AHORRASTE  S/.: " & pfstr_Alineado(Format$(dblPuntosDinero, "###,##0.00"), 14, Derecha)
                Print #1, "Sigue usando tu Monedero del Ahorro"
                Print #1, "Mifarma!!!"
                
                Print #1, String(42, "*")
                Print #1, ""
                Print #1, pfstr_Alineado("Para saber su estado actual de puntos,", 42, centro)
                Print #1, pfstr_Alineado("consultar en la pagina www.mifarma.com.pe", 42, centro)
                Print #1, pfstr_Alineado("Si la tarjeta fue reportada perdida, no", 42, centro)
                Print #1, pfstr_Alineado("procede la acumulación de puntos", 42, centro)
                Print #1, ""
                Print #1, String(42, "-")
            End If
            ' *****************************************
            '--- FIN EXPERTO DEL AHORRO ----
               
            '*/
            '   Se define el mensaje de venta por cada local
            '   23/07/2009 Crueda
            '/*
               
            Dim strMensajeVta As String
            objMensaje.Modalidad = objVenta.CodModalidadVenta   'rstConvenio("COD_MODALIDAD_VENTA").Value
            objMensaje.Tienda = objUsuario.CodigoLocal
            strMensajeVta = objMensaje.Devuelve_Mensaje
               
            If Not strMensajeVta = "" Then
                Print #1, strMensajeVta
            End If
               
            'Print #1, Chr$(27); Chr$(77);
               
            '*/ ..............................
               
            Print #1, String(42, "-")
            Print #1, "*********** www.mifarma.com.pe ***********" & _
                      "           Gracias por su compra          "
            Print #1, pfstr_Alineado(objVenta.Apostrofe & objVenta.SloganBTL & objVenta.Apostrofe, 42, centro)
            Print #1, "       DELIVERY  " & objUsuario.NumeroDelivery & " LAS 24 HORAS   "
'               Print #1, "       DELIVERY LIMA " & objUsuario.NumeroDelivery
            If Val(TPagTkt) > 1 Then
                Print #1,
                Print #1, "Ticket" & Space(1) & Val(PagTkt) & "/" & Val(TPagTkt)
            End If
             
            dblTotal = Format(rstImpresion("MTO_TOTAL").Value, "#0.00")
            'dblMtoTotalRef = 0
       
            'dblMtoTotalRef = DevImpDocCnv(objUsuario.CodigoEmpresa, NumDocumento, TipoDocumento, objUsuario.CodigoLocal)
            If FlgAnulado = True Then
                Print #1, "DOCUMENTO ANULADO"
            End If
            If Not rstConvenio.EOF Then
                strGlosaConv = DevTextoDocCnv(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento, , objUsuario.CodigoLocal)
                If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio Then
'''                  If rstConvenio("FLG_DOC_COPAGO").Value = "1" Then
'''                      Print #1, "B E N E F I C I A R I O  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%)"
'''                  Else
'''                      Print #1, "I N S T I T U C I O N  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%)"
'''                  End If
'''
'''                  If Not "" & rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value = "" And dblMtoTotalRef >= 0 Then
'''                    Print #1, "#REF " & rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value & " " & rstConvenio("NUM_DOCUMENTO_REFER").Value & " (" & 100 - rstConvenio("PCT_BENEFICIARIO").Value & "%) - S/.: " & Format(dblMtoTotalRef, "###,##0.00")
'''                  End If
                End If
            End If
              
'            Set rstDonacion = Impresion(objUsuario.CodigoEmpresa, TipoDocumento, NumDocumento, "1")
'            If rstDonacion.EOF Then
'                Print #1,
'                Print #1,
'             Else
'                dblDonacion = 0
'                While Not rstDonacion.EOF
'                   dblDonacion = dblDonacion + rstDonacion("IMP_CON_REDONDEO").Value
'                   Print #1, left(rstDonacion("DES_HIJO").Value, 54) & "   " & Trim(objUsuario.SmbMoneda) & right("        " & Format(rstDonacion("IMP_CON_REDONDEO").Value, "####0.#0"), 8)
'                   rstDonacion.MoveNext
'                Wend
'                Print #1, Space$(30) & "TOTAL VENTA Y DONACION " & Trim(objUsuario.SmbMoneda) & "    " & right("        " & Format(dblTotal + dblDonacion, "####0.#0"), 8)
'            End If
              
            ''AQUI VA EL MENSAJE
            strMensaje = DevMensaje(objUsuario.CodigoEmpresa, numDocumento, TipoDocumento, objUsuario.CodigoLocal)
            If Len(strMensaje) < 40 Then
                Print #1, strMensaje
            Else
                For i = 0 To Val(Len(strMensaje)) / 40
                    Print #1, strMensaje
                Next i
            End If
            
            For j = 1 To 5
                Print #1, ""
            Next j
      
            'Copia = Copia + 1
            
            '** Corte Parcial **'
            'If flgProformaDLV = False Then
            Print #1, Chr(27) + "d" + "1"
            'End If
            
            '** Corte Total **'
            'Print #1, Chr$(27); "d"; Chr$(2)
            
            '' para TicketProfroma
            'If flgProformaDLV = True Then
            '    'mlaguna
            '    objProforma.ImprimirTicket_Delivery objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, objVenta.NumPedidoPadre, "1"
            '    'imprime proforma de delivery
            'End If
            
            '*********************************************************************************************************
            Close #1
                  
    On Error GoTo 0
    Exit Sub
ErrorImpresora:
    If MsgBox("Existe un Problema con la Impresora" & Chr(13) & _
             "Error: " & CStr(Err.Number) & " - " & Err.Description & Chr(13) & _
             "¿Desea Esperar a ser Resuelto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
          Resume
       Else
          Exit Sub
     End If
End Sub



Private Sub TicketNew(ByVal TipoDocumento As String, ByVal numDocumento As String, Optional PagTkt As String, Optional TPagTkt As String, Optional FlgAnulado As Boolean = False, Optional flgProformaDLV As Boolean = False)

Dim objImpresiones As New clsImpresiones
Dim objServicio As New clsServicio
Dim rsServicio As oraDynaset
Dim rstImpresion As oraDynaset
Dim rstDocumentoMaquina As oraDynaset
Dim rstFormaPago As oraDynaset
Dim rstDonacion As oraDynaset
Dim rstConvenio As oraDynaset
'Dim objDocumento As New clsDocumento
Dim dblDonacion As Double
Dim strGlosaConv As String
Dim strMensaje As String
Dim dblTotal As Double
Dim i As Integer
Dim j As Integer
'''''''''''''''''''''''''''''''''''''''''''''''
Dim p As Printer
Dim intRespuesta As Integer
Dim intNumeroLineas As Integer
Dim dblMtoTotalRef As Double
Dim dblTotUni As Double
Dim strlinea1 As String
Dim strlinea2 As String
Dim strlinea3 As String
Dim strlinea4 As String
Dim strlinea5 As String
Dim strlinea6 As String
Dim strlinea7 As String
'Dim dblPagoEfect As Double
'Dim dblVuelto As Double
Dim MontoAhorro As Double
'**************
    
    MontoAhorro = 0
    
    Set rstConvenio = ListaDoc(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
        
    intNumeroLineas = 0
    dblDonacion = 0
    Set rstDocumentoMaquina = ListaDocumentoMaquina(objUsuario.CodigoEmpresa, TipoDocumento, objUsuario.NombrePC)
    If Not rstDocumentoMaquina.EOF Then intNumeroLineas = rstDocumentoMaquina("NUM_LINEAS_DOC").Value
    Set rstImpresion = Impresion_Ticket(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
       
    Set rstFormaPago = Lista_FPagos_x_Ticket(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
    
    Call frm_VTA_ElegirImpresoras.spCargaDefault(objUsuario.CodigoEmpresa, TipoDocumento, objUsuario.NombrePC)
    
    If gNroCopia = 0 Then
       frm_VTA_ElegirImpresoras.Show vbModal
    End If
       
    On Error GoTo 0
        On Error GoTo ErrorImpresora
            On Error GoTo Mal_Seteo
            GoTo OK_Seteo
Mal_Seteo:
            MsgBox Err.Description, vbExclamation, "Error de Ticketera.....????"
            Exit Sub
OK_Seteo:

'              If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
'                 Dim objImpresion As New clsImpresion
'                 If Not objImpresion.cargaParametros("004") Then
'                     Exit Sub
'                 End If
'            Else
'                 Printer.Width = 225 * 56.7
'                 Printer.Height = (280 / 3) * 56.7
'                 Printer.Print "."
'            End If
      
            '***********************************************************************'
            dblTotUni = 0
            'Open Printer.Port For Output As #1
            '''Open "COM3" For Output As #1
            '**** 16-DOT PITCH FONT *******'
            'Printer.Print Chr$(27); Chr$(58);
            '-------------------------------'
            
            
            'Printer.Font.name = "Arial"
            'Printer.Font.name = "Printer 9cpi"
            'Printer.FontSize = 9
            
            Printer.Font.name = "Verdana"
            Printer.FontSize = 11
'             Printer.Font.Bold = False
            
            Printer.Print "            BOTICAS BTL    "
            '****** Standar ********'
            'Printer.FontSize = 9
            'Printer.Font.name = "Arial"
            
            Printer.FontSize = 8
            
            Printer.Print pfstr_Alineado(Replace(UCase(objUsuario.Empresa), "Í", "I"), 42, centro)
            
            'Printer.Print pfstr_Alineado(Replace(objUsuario.direccion, " - ", "-"), 42, centro)
            Dim strDireccFiscal As String
             strDireccFiscal = Replace(objUsuario.Direccion, " - ", "-")
             While strDireccFiscal <> ""
                
                Print #1, pfstr_Alineado(Mid(strDireccFiscal, 1, 42), 42, centro)
                If Len(strDireccFiscal) > 42 Then
                strDireccFiscal = Mid(strDireccFiscal, 43, Len(strDireccFiscal) - 42)
                Else
                strDireccFiscal = ""
                
                End If
             Wend
            '*******************************************************
            
            Printer.Print pfstr_Alineado("RUC :" & objUsuario.Ruc, 42, centro)
            
            Printer.FontSize = 7
            
            Printer.Print "BTL" & objUsuario.CodigoLocal & Space(1) & CadenaDirecc_Part1(gclsOracle.FN_Valor("CMR.PKG_LOCAL.FN_DIR", objUsuario.CodigoLocal))
            
            Printer.FontSize = 8
            
            Printer.Print Space(8) & CadenaDirecc_Part2(gclsOracle.FN_Valor("CMR.PKG_LOCAL.FN_DIR", objUsuario.CodigoLocal))
            Printer.FontSize = 7
            Printer.Print "FECHA:" & rstImpresion("FCH_REGISTRA").Value & Space(1) & "CAJA:" & " " & rstImpresion("COD_MAQUINA").Value
            Printer.Print "Ticket" & " " & "Nº" & Space(1) & rstImpresion("NUM_DOCUMENTO").Value & " " & "Serie:" & " " & rstImpresion("COD_SERIE_REL").Value
            Printer.FontSize = 8
            Printer.Print "Cajero:" & Space(1) & rstImpresion("COD_USUARIO").Value & Space(1) & rstImpresion("NOMB_CAJERO").Value
            Printer.Print ""
            Printer.Print left("CODIGO" & Space$(5), 5) & Space$(1) & _
                      left("DESCRIPCION" & Space$(20), 20) & Space$(1) & _
                      right(Space$(5) & "CANT.", 5) & Space$(1) & _
                      right(Space$(7) & "IMPORTE", 7) & Space$(1)
            Printer.Print String(46, "-")
            'Printer.FontBold = False
            'Printer.FontSize = 7
            While Not rstImpresion.EOF
                'Detalle
               If Len(rstImpresion("DES_PRODUCTO").Value) > 20 Then ' SI LA DESCRIPCION DE PRODUCTO ES MAYOR A 20 CARACTERES'
                         Printer.Print left(rstImpresion("COD_PRODUCTO").Value & Space(5), 5) & Space(1) & _
                                         left(Mid(rstImpresion("DES_PRODUCTO").Value, 1, 20) & Space(20), 20) & Space(1) & _
                                         right(Space(5) & rstImpresion("CANTIDAD").Value, 5) & Space(2) & _
                                         pfstr_Alineado(rstImpresion("SUB_TOTAL").Value, 7, Derecha) & Space(1)
                                         'right(Space(7) & rstImpresion("SUB_TOTAL").Value, 7) & Space(1)
                                         
                         Printer.Print Space(6) & left(Mid(rstImpresion("DES_PRODUCTO").Value, 21, 40) & Space(20), 20) & Space(3)
                        
                         'mlaguna: Agrega numero de partida a inafectos 19/04/2010
                         If rstImpresion("PCT_IGV").Value = 0 Then
                         Printer.Print Space(6) & "PA: " & left(rstImpresion("NUM_PART_ARANCEL").Value, 20)
                         End If
             Else
                         Printer.Print left(rstImpresion("COD_PRODUCTO").Value & Space(5), 5) & Space(1) & _
                                   left(Mid(rstImpresion("DES_PRODUCTO").Value, 1, 30) & Space(30), 30) & Space(2) & _
                                   Space(30) & right(Space(5) & rstImpresion("CANTIDAD").Value, 5) & Space(1) & _
                                   pfstr_Alineado(rstImpresion("SUB_TOTAL").Value, 7, Derecha) & Space(1)
                                   
                         'mlaguna: Agrega numero de partida a inafectos 19/04/2010
                         If rstImpresion("PCT_IGV").Value = 0 Then
                         Printer.Print Space(6) & "PA: " & left(rstImpresion("NUM_PART_ARANCEL").Value, 20)
                         End If
            End If
            Dim strLote, strFecVencimiento As String
            strLote = Trim("" & rstImpresion("NUM_LOTE").Value)
            strFecVencimiento = Trim("" & rstImpresion("FCH_VENCIMIENTO").Value)
            If Not (strLote & strFecVencimiento) = "" Then
                Printer.Print Space(6) & left(IIf(strLote = "", "", "Lote:" & strLote) & Space(25), 25) & Space(3)
                Printer.Print Space(6) & left(IIf(strFecVencimiento = "", "", "F. Venc.: " & strFecVencimiento) & Space(25), 25) & Space(3)
            End If
            
            
                MontoAhorro = MontoAhorro + Val(rstImpresion("MTO_AHORRO").Value)
                rstImpresion.MoveNext
            Wend
            dblTotUni = rstImpresion.RecordCount
            rstImpresion.MoveFirst
            'Print #1, ""
            'Print #1, Space(5) & "Total Productos" & " " & Val(dblTotUni)'comentado por arturo escate 19/03/2009
            'Print #1, pfstr_Alineado("Total Productos" & " " & Val(dblTotUni), 42, Derecha)
            'Print #1, pfstr_Alineado(String(22, "-"), 42, Derecha)
            Printer.Print pfstr_Alineado("Item(s):" & Val(dblTotUni), 42, Izquierda)
            If rstImpresion("COD_TIPO_DOCUMENTO").Value = "TKF" Then
               If rstImpresion("MTO_EXONERADO").Value = 0 Then
                   Printer.Print pfstr_Alineado(left("Valor Venta" & Space(11), 11) & Space(7) & "S/." & right(Space(10) & Format(rstImpresion("MTO_BASE_IMP").Value, "#0.00"), 10) & Space(1), 42, Derecha)
               Else
                  Printer.Print pfstr_Alineado(left("Valor venta" & Space(11), 11) & Space(7) & "S/." & right(Space(10) & Format(rstImpresion("MTO_EXONERADO").Value, "#0.00"), 10) & Space(1), 42, Derecha)
               End If
                  Printer.Print pfstr_Alineado(left("IGV " & objUsuario.IGV & "%." & Space(11), 11) & Space(7) & "S/." & right(Space(10) & Format(rstImpresion("MTO_IMPUESTO").Value, "#0.00"), 10) & Space(1), 42, Derecha)
            End If
            Printer.Print pfstr_Alineado(String(22, "-"), 42, Derecha)
            ''''''' NUEVO 27/04/2011
            Dim XRedondeo As Double
            
            While Not rstFormaPago.EOF
                If rstFormaPago("COD_FORMA_PAGO") = "008" Then
                    XRedondeo = XRedondeo + rstFormaPago("IMP_MONEDA_NAC").Value
                End If
                rstFormaPago.MoveNext
            Wend
            rstFormaPago.MoveFirst
            ''''''' NUEVO 27/04/2011
            'Print #1, pfstr_Alineado("Redondeo:" & Format(Val(XRedondeo), "#0.00") & Space(10) & left("Total" & Space(6), 6) & Space(2) & "S/." & right(Space(1) & Format(Val(rstImpresion("MTO_TOTAL").Value) + Val(XRedondeo), "#0.00"), 10) & Space(1), 42, Derecha)
            Printer.Print pfstr_Alineado(Space(10) & left("Total" & Space(6), 6) & Space(2) & "S/." & right(Space(1) & Format(Val(rstImpresion("MTO_TOTAL").Value), "#0.00"), 10) & Space(1), 42, Derecha)
            'Print #1, ""
            ''Aca imprime las formas de pago
            Dim dblMonNac As Double
            Dim dblVuelto As Double
            'Dim montoRedoneot As Double
            dblMonNac = 0: dblVuelto = 0
            rstFormaPago.MoveFirst
            While Not rstFormaPago.EOF
                Printer.Print left(rstFormaPago("DES_HIJO") & Space(25), 25) & Space(4) & _
                          "S/." & " " & right(Space(8) & Format(rstFormaPago("IMP_MONEDA_NAC").Value, "#0.00"), 8)
                'If rstFormaPago("COD_FORMA_PAGO") = "008" Then
                          'montoRedoneot = Val(montoRedoneot + "" & rstFormaPago("IMP_MONEDA_NAC").Value)
                'End If
                dblMonNac = dblMonNac + Val(rstFormaPago("IMP_MONEDA_NAC").Value)
                dblVuelto = dblVuelto + Val(rstFormaPago("IMP_VUELTO").Value)
                rstFormaPago.MoveNext
            Wend
            Printer.Print pfstr_Alineado(String(22, "-"), 42, Derecha)
            Printer.Print "Total a pagar                S/.   " & " " & Format(Val(rstImpresion("MTO_TOTAL").Value) + XRedondeo, "#0.00")
            Printer.Print "Vuelto                       S/.   " & " " & Format(dblVuelto, "#0.00")
            
            '& Space(10) & "" & "Total", 26, Derecha) & Space(1) & _
                      right(Space(5) & "", 5) & Space(2) & _
                      pfstr_Alineado(Format(rstImpresion("MTO_TOTAL").Value, "#0.00"), 7, Derecha) & Space(1)


''''HASTA ACA TODO ALINEADO
            
            
            
            Set rsServicio = objServicio.ListaServiciosAsoc(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
            While Not rsServicio.EOF
             If objVenta.EsNavsat("" & rsServicio("COD_TIPO_SERVICIO")) = "1" Then
                 Printer.Print "PIN o Telefono : " & rsServicio("NUM_SUMINISTRO")
                 Printer.Print "Nº Aprobacion:" & rsServicio("NUM_RECIBO")
                 If Not IsNull(rsServicio("DES_LEYENDA")) Then
                     If InStr(rsServicio("DES_LEYENDA"), "1[") > 0 Then
                         strlinea1 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "1[") + 2), InStr(rsServicio("DES_LEYENDA"), "]1") - ((InStr(rsServicio("DES_LEYENDA"), "1[") + 2)))
                         If strlinea1 <> "" Then Printer.Print strlinea1
                     End If
                     If InStr(rsServicio("DES_LEYENDA"), "2[") > 0 Then
                         strlinea2 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "2[") + 2), InStr(rsServicio("DES_LEYENDA"), "]2") - ((InStr(rsServicio("DES_LEYENDA"), "2[") + 2)))
                         If strlinea2 <> "" Then Printer.Print strlinea2
                     End If
                     If InStr(rsServicio("DES_LEYENDA"), "3[") > 0 Then
                         strlinea3 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "3[") + 2), InStr(rsServicio("DES_LEYENDA"), "]3") - ((InStr(rsServicio("DES_LEYENDA"), "3[") + 2)))
                         If strlinea3 <> "" Then Printer.Print strlinea3
                     End If
                     If InStr(rsServicio("DES_LEYENDA"), "4[") > 0 Then
                         strlinea4 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "4[") + 2), InStr(rsServicio("DES_LEYENDA"), "]4") - ((InStr(rsServicio("DES_LEYENDA"), "4[") + 2)))
                         If strlinea4 <> "" Then Printer.Print strlinea4
                     End If
                     If InStr(rsServicio("DES_LEYENDA"), "5[") > 0 Then
                         strlinea5 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "5[") + 2), InStr(rsServicio("DES_LEYENDA"), "]5") - ((InStr(rsServicio("DES_LEYENDA"), "5[") + 2)))
                         If strlinea5 <> "" Then Printer.Print strlinea5
                     End If
                     If InStr(rsServicio("DES_LEYENDA"), "6[") > 0 Then
                         strlinea6 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "6[") + 2), InStr(rsServicio("DES_LEYENDA"), "]6") - ((InStr(rsServicio("DES_LEYENDA"), "6[") + 2)))
                         If strlinea6 <> "" Then Printer.Print strlinea6
                     End If
                     If InStr(rsServicio("DES_LEYENDA"), "7[") > 0 Then
                         strlinea7 = Mid(rsServicio("DES_LEYENDA"), (InStr(rsServicio("DES_LEYENDA"), "7[") + 2), InStr(rsServicio("DES_LEYENDA"), "]7") - ((InStr(rsServicio("DES_LEYENDA"), "7[") + 2)))
                         If strlinea7 <> "" Then Printer.Print strlinea7
                     End If
                 End If
             End If
             rsServicio.MoveNext
            Wend
            Set objServicio = Nothing
            
'''''            dblTotUni = rstImpresion.RecordCount
'''''            Print #1, ""
'''''            'Print #1, Space(5) & "Total Productos" & " " & Val(dblTotUni)'comentado por arturo escate 19/03/2009
'''''            Print #1, pfstr_Alineado("Total Productos" & " " & Val(dblTotUni), 42, Derecha)
'''''
'''''            Print #1, ""
          
'            dblPagoEfect = DevPagoEfectVuelto_Ticket(objUsuario.CodigoEmpresa, _
'                                                     TipoDocumento, _
'                                                     NumDocumento, _
'                                                     "1")
'
'            dblVuelto = DevPagoEfectVuelto_Ticket(objUsuario.CodigoEmpresa, _
'                                                  TipoDocumento, _
'                                                  NumDocumento, _
'                                                  "2")
            
            
''''''''''            If rstImpresion("COD_TIPO_DOCUMENTO").Value = "TKF" Then
''''''''''               If rstImpresion("MTO_EXONERADO").Value = 0 Then
''''''''''                  'Print #1, Space(15) & "Valor venta   S/. " & " " & Format(rstImpresion("MTO_BASE_IMP").Value, "#0.00") 'comentado por arturo escate 19/03/2009
''''''''''                   'Print #1, pfstr_Alineado("Valor venta   S/. " & " " & Format(rstImpresion("MTO_BASE_IMP").Value, "#0.00"), 42, Derecha)
''''''''''                   Print #1, pfstr_Alineado(left("Valor Venta" & Space(11), 11) & Space(7) & "S/." & right(Space(10) & Format(rstImpresion("MTO_BASE_IMP").Value, "#0.00"), 10) & Space(1), 42, Derecha)
''''''''''               Else
''''''''''                  'Print #1, Space(15) & "Valor venta   S/. " & " " & Format(rstImpresion("MTO_EXONERADO").Value, "#0.00") 'comentado por arturo escate 19/03/2009
''''''''''                  'Print #1, pfstr_Alineado("Valor venta   S/. " & " " & Format(rstImpresion("MTO_EXONERADO").Value, "#0.00"), 42, Derecha)
''''''''''                  Print #1, pfstr_Alineado(left("Valor venta" & Space(11), 11) & Space(7) & "S/." & right(Space(10) & Format(rstImpresion("MTO_EXONERADO").Value, "#0.00"), 10) & Space(1), 42, Derecha)
''''''''''               End If
''''''''''                  'Print #1, Space(15) & "IGV " & objUsuario.IGV & "%.      S/. " & " " & Format(rstImpresion("MTO_IMPUESTO").Value, "#0.00")'comentado por arturo escate 19/03/2009
''''''''''                  'Print #1, pfstr_Alineado("IGV " & objUsuario.IGV & "%.      S/. " & " " & Format(rstImpresion("MTO_IMPUESTO").Value, "#0.00"), 42, Derecha)
''''''''''                  Print #1, pfstr_Alineado(left("IGV " & objUsuario.IGV & "%." & Space(11), 11) & Space(7) & "S/." & right(Space(10) & Format(rstImpresion("MTO_IMPUESTO").Value, "#0.00"), 10) & Space(1), 42, Derecha)
''''''''''            End If
                  'Print #1, Space(15) & "Total         S/. " & " " & Format(rstImpresion("MTO_TOTAL").Value, "#0.00") 'comentado por arturo escate 19/03/2009
                  'Print #1, pfstr_Alineado("Total         S/. " & " " & Format(rstImpresion("MTO_TOTAL").Value, "#0.00"), 42, Derecha)
                  'Print #1, pfstr_Alineado(left("Total" & Space(11), 11) & Space(7) & "S/." & right(Space(10) & Format(rstImpresion("MTO_TOTAL").Value, "#0.00"), 10) & Space(1), 42, Derecha)
                  'Print #1, Space(15) & "Total         S/. " & " " & Format(rstCAB_TiketFac("MTO_TOTAL").Value, "#0.00")
                  Printer.Print ""
                  
                  'Print #1, "Pago Efectivo                S/. " & " " & Format(dblPagoEfect, "#0.00")
                  'Print #1, "Vuelto                       S/. " & " " & Format(dblVuelto, "#0.00")

                  If Round(MontoAhorro, 2) > 0 Then
                    If Not DevImpDocCnv(objUsuario.CodigoEmpresa, numDocumento, TipoDocumento, objUsuario.CodigoLocal) = MontoAhorro Then
                        Printer.Print ""
                        Printer.Print "USTED HA AHORRADO           S/.   " & " " & Format(MontoAhorro, "#0.00")
                        Printer.Print ""
                    End If
                  End If
                  
            If Not rstImpresion("DES_RAZON_SOCIAL").Value = "" Then
                If rstImpresion("COD_TIPO_DOCUMENTO").Value = "TKF" Then
                   Printer.Print "Ruc Ciente   :" & " " & rstImpresion("COD_CLIENTE").Value
                   Printer.Print "Razon Social:" & " " & rstImpresion("DES_RAZON_SOCIAL").Value
                Else
                    If Len("" & rstImpresion("COD_CLIENTE").Value) = 11 Then
                        Printer.Print "Ruc Ciente   :" & " " & rstImpresion("COD_CLIENTE").Value
                        Printer.Print "Razon Social:" & " " & rstImpresion("DES_RAZON_SOCIAL").Value
                    Else
                        Printer.Print "Doc. Ident :" & " " & rstImpresion("COD_CLIENTE").Value
                        Printer.Print "Nombre     :" & " " & rstImpresion("DES_RAZON_SOCIAL").Value
                    End If
                    
                End If
            End If
            
           
            
            
            If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio Then
                    Dim objConvenio As New clsConvenio
                    Dim objCliente  As New clsCliente
                If objConvenio.ListaConvenioCompetencia(rstConvenio("COD_CONVENIO").Value) = 0 Then
                    'dblMtoTotalRef = 0
       
                    'dblMtoTotalRef = DevImpDocCnv(objUsuario.CodigoEmpresa, NumDocumento, TipoDocumento, objUsuario.CodigoLocal)
                
                    Printer.Print ""
                    Printer.Print "Convenio: " & rstImpresion("DES_CONVENIO").Value
                    '''''''''''''''''''''''''''''''
                    
                    
                    If rstConvenio("FLG_DOC_COPAGO").Value = "1" Then
                        Printer.Print "B E N E F I C I A R I O  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO_ORIG").Value & "%)"
                        objVenta.Beneficiario = strGlosaConv
                        If objConvenio.FnDev_Deducible(rstConvenio("COD_CONVENIO").Value) = "0" Then
                           Printer.Print rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value & " " & rstConvenio("NUM_DOCUMENTO_REFER").Value & " (" & 100 - rstConvenio("PCT_BENEFICIARIO_ORIG").Value & "%)" & IIf("" & rstConvenio("FLG_IMPRIME_COPAGO").Value = "1", " - S/.: " & Format(dblMtoTotalRef, "###,##0.00"), "") & "  "
                        End If
                    Else
                        Set objCliente = New clsCliente
                        objVenta.Beneficiario = objCliente.Nombre(IIf(IsNull(rstConvenio("COD_CLIENTE").Value), "", rstConvenio("COD_CLIENTE").Value))
                        Set objCliente = Nothing
                        If dblMtoTotalRef > -1 And rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value <> "" Then
                            Printer.Print "#REF " & rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value & " " & rstConvenio("NUM_DOCUMENTO_REFER").Value & " (" & 100 - rstConvenio("PCT_BENEFICIARIO").Value & "%) - S/.: " & Format(dblMtoTotalRef, "###,##0.00") & Space(2)
                            Printer.Print "Beneficiario:" & Space(2) & objVenta.Beneficiario
                        Else
                            Printer.Print "Beneficiario:" & Space(2) & objVenta.Beneficiario
                        End If
                        Printer.Print "I N S T I T U C I O N  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%)"
                    End If
    
                    Set objConvenio = Nothing
                End If
            End If
            
               
               If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio Then
                    Dim strMensajeConvenio As String
                    If objConvenio.ListaConvenioCompetencia(rstConvenio("COD_CONVENIO").Value) = 0 Then
                       strMensajeConvenio = "" & Me.MensajeImpresionConvenio(rstConvenio("COD_CONVENIO").Value)
                       If Not strMensajeConvenio = "" Then
                              Printer.Print strMensajeConvenio
                       End If
                       ''11/03/2010 SE AGREGO PARA QUE EL CONVENIO BTL EN EFECTIVO TAMBIEN SALGA LA FIRMA
                       If rstImpresion("COD_TIPO_VTAFIN").Value = "2" Or gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_CNV_BTL_EFECTIVO") = rstConvenio("COD_CONVENIO").Value Then
                            Printer.Print ""
                            Printer.Print ""
                            Printer.Print ""
                            'Printer.Print "    " & String(32, "-") & "    "
                            Printer.Print "Nombre: " & rstImpresion("DES_AUX_CLI_NOMBRE").Value
                            Printer.Print "DNI: " & rstConvenio("DNI_BENEFICIARIO").Value
                       End If
                    End If
               End If
               
               
               '*/
               '   Se define el mensaje de venta por cada local
               '   23/07/2009 Crueda
               '/*
               
               Dim strMensajeVta As String
               objMensaje.Modalidad = objVenta.CodModalidadVenta   'rstConvenio("COD_MODALIDAD_VENTA").Value
               objMensaje.Tienda = objUsuario.CodigoLocal
               strMensajeVta = objMensaje.Devuelve_Mensaje
               
               If Not strMensajeVta = "" Then
                    Printer.Print strMensajeVta
               End If
               
               '*/ ..............................
               
               Printer.Print String(42, "-")
               Printer.Print "************* www.btl.com.pe *************" & _
                         "           Gracias por su compra          "
               Printer.Print pfstr_Alineado(objVenta.Apostrofe & objVenta.SloganBTL & objVenta.Apostrofe, 42, centro)
               Printer.Print "       DELIVERY  " & objUsuario.NumeroDelivery & " LAS 24 HORAS   "
            If Val(TPagTkt) > 1 Then
               Printer.Print
               Printer.Print "Ticket" & Space(1) & Val(PagTkt) & "/" & Val(TPagTkt)
            End If
             
            dblTotal = Format(rstImpresion("MTO_TOTAL").Value, "#0.00")
            dblMtoTotalRef = 0
       
            dblMtoTotalRef = DevImpDocCnv(objUsuario.CodigoEmpresa, numDocumento, TipoDocumento, objUsuario.CodigoLocal)
          
            If FlgAnulado = True Then
                Printer.Print "DOCUMENTO ANULADO"
            End If
            If Not rstConvenio.EOF Then
              strGlosaConv = DevTextoDocCnv(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento, , objUsuario.CodigoLocal)
              If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio Then
              
              
              

              
              
              
              
'''                  If rstConvenio("FLG_DOC_COPAGO").Value = "1" Then
'''                      Print #1, "B E N E F I C I A R I O  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%)"
'''                  Else
'''                      Print #1, "I N S T I T U C I O N  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%)"
'''                  End If
'''
'''                  If Not "" & rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value = "" And dblMtoTotalRef >= 0 Then
'''                    Print #1, "#REF " & rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value & " " & rstConvenio("NUM_DOCUMENTO_REFER").Value & " (" & 100 - rstConvenio("PCT_BENEFICIARIO").Value & "%) - S/.: " & Format(dblMtoTotalRef, "###,##0.00")
'''                  End If
                  
              End If
            End If
              
'            Set rstDonacion = Impresion(objUsuario.CodigoEmpresa, TipoDocumento, NumDocumento, "1")
'            If rstDonacion.EOF Then
'                Print #1,
'                Print #1,
'             Else
'                dblDonacion = 0
'                While Not rstDonacion.EOF
'                   dblDonacion = dblDonacion + rstDonacion("IMP_CON_REDONDEO").Value
'                   Print #1, left(rstDonacion("DES_HIJO").Value, 54) & "   " & Trim(objUsuario.SmbMoneda) & right("        " & Format(rstDonacion("IMP_CON_REDONDEO").Value, "####0.#0"), 8)
'                   rstDonacion.MoveNext
'                Wend
'                Print #1, Space$(30) & "TOTAL VENTA Y DONACION " & Trim(objUsuario.SmbMoneda) & "    " & right("        " & Format(dblTotal + dblDonacion, "####0.#0"), 8)
'            End If
              
            ''AQUI VA EL MENSAJE
            strMensaje = DevMensaje(objUsuario.CodigoEmpresa, numDocumento, TipoDocumento, objUsuario.CodigoLocal)
            If Len(strMensaje) < 40 Then
                    Printer.Print strMensaje
              Else
                For i = 0 To Val(Len(strMensaje)) / 40
                    Printer.Print strMensaje
                Next i
            End If
            
            For j = 1 To 5
               Printer.Print ""
            Next j
      
            'Copia = Copia + 1
            
            '** Corte Parcial **'
            'If flgProformaDLV = False Then
           'Print #1, Chr(27) + "d" + "1"
            'End If
            
            '** Corte Total **'
            'Print #1, Chr$(27); "d"; Chr$(2)
            
            '' para TicketProfroma
            'If flgProformaDLV = True Then
            '    'mlaguna
            '    objProforma.ImprimirTicket_Delivery objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, objVenta.NumPedidoPadre, "1"
            '    'imprime proforma de delivery
            'End If
            
            '*********************************************************************************************************
            
            Printer.EndDoc
                  
    On Error GoTo 0
    Exit Sub
ErrorImpresora:
    If MsgBox("Existe un Problema con la Impresora" & Chr(13) & _
             "Error: " & CStr(Err.Number) & " - " & Err.Description & Chr(13) & _
             "¿Desea Esperar a ser Resuelto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
          Resume
       Else
          Exit Sub
     End If

End Sub




Public Function MaqEsCajeroCorresponsal(ByVal Cia As String, _
                        ByVal codMaquina As String) As Integer
                        
On Error GoTo CtrlErr

    MaqEsCajeroCorresponsal = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO.FN_MAQ_ES_CAJERO_CORRESPONSAL", Cia, codMaquina)
    
    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.MaqEsCajeroCorresponsal", Err.Description
End Function

Public Function DevNomRepartidor(ByVal vstrCia As String, _
                                 ByVal vstrLocal As String, _
                                 ByVal vstrTipDoc As String, _
                                 ByVal vstrDocumento As String) As String

On Error GoTo CtrlErr

    DevNomRepartidor = gclsOracle.FN_Valor("BTLPROD.PKG_REPARTIDOR.FN_NOM_REPARTIDOR", vstrCia, vstrLocal, vstrTipDoc, vstrDocumento)
    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.DevNomRepartidor", Err.Description

End Function

Public Function DocumentoEsConvenio(ByVal vstrCodLocal As String, _
                                    ByVal vstrNumProforma As String, _
                                    ByVal vstrFlgProforma As String) As String
On Error GoTo CtrlErr
        DocumentoEsConvenio = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO.FN_C_DOCUMENTOS_ES_CONVENIO", vstrCodLocal, _
                                                                                                       vstrNumProforma, _
                                                                                                       vstrFlgProforma)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.DocumentoEsConvenio", Err.Description
End Function

Public Function FindConvenioEnBtlprod(ByVal vstrCia As String, _
                                      ByVal vstrCodCnvGLM As String) As String
                                    
On Error GoTo CtrlErr
        FindConvenioEnBtlprod = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO.FN_FIND_CNV_EN_BTLPROD", vstrCia, _
                                                                                                    vstrCodCnvGLM)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.DocumentoEsConvenio", Err.Description
End Function
    
Public Function DevPagoEfectVuelto_Ticket(ByVal vstrCia As String, _
                                          ByVal vstrCodTipoDoc As String, _
                                          ByVal vstrNumDoc As String, _
                                          ByVal vstrTipoSal As String) As Double
On Error GoTo CtrlErr
        DevPagoEfectVuelto_Ticket = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO.FN_DEV_PAGEFEC_VUELTO_TICKET", vstrCia, _
                                                                                                              vstrCodTipoDoc, _
                                                                                                              vstrNumDoc, _
                                                                                                              vstrTipoSal)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.DevPagoEfectVuelto_Ticket", Err.Description
End Function

Public Function ListaDocAnu_x_Dlv(ByVal vstrCia As String, _
                                  ByVal vstrLocal As String)
On Error GoTo CtrlErr
    Set ListaDocAnu_x_Dlv = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.LISTA_DOC_ANULADOS_X_DLV", 0, vstrCia, _
                                                                                                     vstrLocal)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaDocAnu_x_Dlv", Err.Description
End Function

Public Function LstDetDocumentoxTipo(ByVal vstrCia As String, _
                                     ByVal vstrCodTipoDoc As String, _
                                     ByVal vstrNumDocumento As String) As oraDynaset
    On Error GoTo CtrlErr
    Set LstDetDocumentoxTipo = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_DET_DOCUMENTO_X_TIPO", 0, vstrCia, _
                                                                                                          vstrCodTipoDoc, _
                                                                                                          vstrNumDocumento)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.LstDetDocumentoxTipo", Err.Description
End Function

Public Function ListaCobros(ByVal vstrCia As String, _
                            ByVal vstrCodTipoDoc As String, _
                            ByVal vstrNumDocumento As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaCobros = gclsOracle.FN_Cursor("BTLPROD.pkg_documento.fn_lista_cobro", 0, vstrCia, _
                                                                                      vstrCodTipoDoc, _
                                                                                      vstrNumDocumento)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaCobros", Err.Description
End Function


Public Function ListaDetDocEmitidos(ByVal vstrCia As String, _
                                    ByVal vstrCodLocal As String, _
                                    ByVal vstrCodTipoDoc As String, _
                                    ByVal vstrNumDocIni As String, _
                                    ByVal vstrNumDocFin As String, _
                                    ByVal vstrCodLiquidacion As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaDetDocEmitidos = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_LISTA_DET_DOC_EMITIDOS", 0, vstrCia, _
                                                                                                         vstrCodLocal, _
                                                                                                         vstrCodTipoDoc, _
                                                                                                         vstrNumDocIni, _
                                                                                                         vstrNumDocFin, _
                                                                                                         vstrCodLiquidacion)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaDetDocEmitidos", Err.Description
End Function

Public Function ListaProdDocumento(ByVal vstrCia As String, _
                                    ByVal vstrCodTipoDoc As String, _
                                    ByVal vstrNumDocumento As String, _
                                    ByVal vstrCodLocal As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaProdDocumento = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_LISTA_PROD_DOCUMENTO", 0, vstrCia, _
                                                                                                         vstrCodTipoDoc, _
                                                                                                         vstrNumDocumento, _
                                                                                                         vstrCodLocal)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaProdDocumento", Err.Description
End Function

Public Sub ImprimeDocAlamcenLocal(ByVal Cia As String, ByVal codTipoDoc As String, ByVal numDocumento As String, ByVal CodLocal As String)
Dim rsListaProdDocumento As oraDynaset
Dim strCia As String
Dim strFchRegistra As String
Dim strCodLocal As String
Dim strCodUsuario As String
Dim strNombre As String
Dim strCodTipoDocumento As String
Dim strNumDocumento As String
Dim strCodProducto As String
Dim strDesProducto As String
Dim strDesLaboratorio As String
Dim strDesLinea As String
Dim strCant As String

    On Error GoTo handle

    Set rsListaProdDocumento = ListaProdDocumento(Cia, codTipoDoc, numDocumento, CodLocal)
    
    If Not rsListaProdDocumento.EOF Then
            strCia = ""
            strFchRegistra = ""
            strCodLocal = ""
            strCodUsuario = ""
            strNombre = ""
            strCodTipoDocumento = ""
            strNumDocumento = ""
    
            If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
                 Dim objImpresion As New clsImpresion
                 If Not objImpresion.cargaParametros("004") Then
                     Exit Sub
                 End If
            Else
                 Printer.Width = 225 * 56.7
                 Printer.Height = (280 / 3) * 56.7
                 Printer.Print "."
            End If

            strCia = IIf(IsNull(rsListaProdDocumento("CIA").Value), "", rsListaProdDocumento("CIA").Value)
            strFchRegistra = IIf(IsNull(rsListaProdDocumento("FCH_REGISTRA").Value), "", rsListaProdDocumento("FCH_REGISTRA").Value)
            strCodLocal = IIf(IsNull(rsListaProdDocumento("COD_LOCAL").Value), "", rsListaProdDocumento("COD_LOCAL").Value)
            strCodUsuario = IIf(IsNull(rsListaProdDocumento("COD_USUARIO").Value), "", rsListaProdDocumento("COD_USUARIO").Value)
            strNombre = IIf(IsNull(rsListaProdDocumento("NOMBRE").Value), "", rsListaProdDocumento("NOMBRE").Value)
            strCodTipoDocumento = IIf(IsNull(rsListaProdDocumento("COD_TIPO_DOCUMENTO").Value), "", rsListaProdDocumento("COD_TIPO_DOCUMENTO").Value)
            strNumDocumento = IIf(IsNull(rsListaProdDocumento("NUM_DOCUMENTO").Value), "", rsListaProdDocumento("NUM_DOCUMENTO").Value)
            
            Printer.FontBold = True
            Printer.Print objUsuario.Empresa
            Printer.FontBold = False
            Printer.FontName = "Draft 17cpi"
            Printer.Print Space(105) & "Fecha: " & Format(strFchRegistra, "DD-MM-YYYY")
            Printer.Print Space(105) & "Hora : " & Format(strFchRegistra, "HH:MM:SS AM/PM")
            Printer.Print
            Printer.Print "Usuario : " & right(Space(5) & strCodUsuario, 5) & Space(1) & left(strNombre & Space(50), 50) & Space(39) & "Página : " & Printer.Page
            Printer.FontBold = True
            Printer.FontName = "Draft 10cpi"
            Printer.Print fstr_Centrar(strCodTipoDocumento & Space(1) & " Nº " & Format(strNumDocumento, "000-0000000"), 81)
            Printer.Print
            Printer.Print
            Printer.FontName = "Draft 12cpi"
            Printer.Print "CODIGO     PRODUCTO                                        LABORATORIO                 CANT  "
            Printer.Print "----- --------------------------------------------------  -------------------------  --------"
            Printer.Print
            strCodProducto = ""
            strDesProducto = ""
            strDesLaboratorio = ""
            strDesLinea = ""
            strCant = ""
            While Not rsListaProdDocumento.EOF
                    strCodProducto = IIf(IsNull(rsListaProdDocumento("COD_PRODUCTO").Value), "", rsListaProdDocumento("COD_PRODUCTO").Value)
                    strDesProducto = IIf(IsNull(rsListaProdDocumento("DES_PRODUCTO").Value), "", rsListaProdDocumento("DES_PRODUCTO").Value)
                    strDesLaboratorio = IIf(IsNull(rsListaProdDocumento("DES_LABORATORIO").Value), "", rsListaProdDocumento("DES_LABORATORIO").Value)
                    '''strDesLinea = IIf(IsNull(rsListaProdDocumento("DES_LINEA").Value), "", rsListaProdDocumento("DES_LINEA").Value)
                    strCant = IIf(IsNull(rsListaProdDocumento("CANT").Value), "", rsListaProdDocumento("CANT").Value)
                Printer.Print CME(strCodProducto, 5, "D") & Space(1) & _
                        CME(strDesProducto, 50, "I") & Space(2) & _
                        CME(strDesLaboratorio, 24, "I") & Space(2) & _
                        CME(strCant, 8, "D")
                rsListaProdDocumento.MoveNext
            Wend
            Printer.EndDoc
    End If
    
    On Error GoTo 0
    
    Exit Sub
    
handle:
    If MsgBox("Existe un Problema con la Impresora" & Chr(13) & _
             "Error: " & CStr(Err.Number) & " - " & Err.Description & Chr(13) & _
             "¿Desea Esperar a ser Resuelto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
          Resume
       Else
          Exit Sub
     End If

End Sub

Public Function ListaInsumos_x_Doc(ByVal vstrCia As String, _
                                   ByVal vstrTipoDoc As String, _
                                   ByVal vstrNumDoc As String) As oraDynaset

    On Error GoTo CtrlErr
    Set ListaInsumos_x_Doc = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.LISTA_INSUMOS_RM_X_DOC", 0, vstrCia, vstrTipoDoc, vstrNumDoc)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaInsumos_x_Doc", Err.Description
End Function

Public Function Lista_FPagos_x_Ticket(ByVal vstrCia As String, _
                                      ByVal vstrTipoDoc As String, _
                                      ByVal vstrNumDoc As String) As oraDynaset
    On Error GoTo CtrlErr
    Set Lista_FPagos_x_Ticket = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_DEV_FP_TIKET", 0, vstrCia, vstrTipoDoc, vstrNumDoc)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaInsumos_x_Doc", Err.Description
End Function

Public Function CadenaDirecc_Part1(ByVal vstrCadDirecc As String)
    Dim i As Integer
    Dim C1 As String
    C1 = ""
    For i = 1 To Len(vstrCadDirecc)
       If i <= 31 Then
           C1 = C1 + Mid(vstrCadDirecc, i, 1)
       End If
    Next i
    CadenaDirecc_Part1 = C1
End Function

Public Function CadenaDirecc_Part2(ByVal vstrCadDirecc As String)
    Dim i As Integer
    Dim C2 As String
    C2 = ""
    For i = 1 To Len(vstrCadDirecc)
       If i > 31 Then
           C2 = C2 + Mid(vstrCadDirecc, i, 1)
       End If
    Next i
    CadenaDirecc_Part2 = C2
End Function

Public Function CadenaDirecc_Part3(ByVal vstrCadDirecc As String)
    Dim i As Integer
    Dim C1 As String
    C1 = ""
    For i = 1 To Len(vstrCadDirecc)
       If i <= 53 Then
           C1 = C1 + Mid(vstrCadDirecc, i, 1)
       End If
    Next i
    CadenaDirecc_Part3 = C1
End Function

Public Function CadenaDirecc_Part4(ByVal vstrCadDirecc As String)
    Dim i As Integer
    Dim C2 As String
    C2 = ""
    For i = 1 To Len(vstrCadDirecc)
       If i > 53 Then
           C2 = C2 + Mid(vstrCadDirecc, i, 1)
       End If
    Next i
    CadenaDirecc_Part4 = C2
End Function

Public Function ValidaExisteLiquidacion(ByVal Cia As String, _
                        ByVal codMaquina As String, _
                        ByVal CodLocal As String, _
                        ByVal CodUsuario As String) As String
                        
On Error GoTo CtrlErr

    ValidaExisteLiquidacion = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO.FN_VALIDA_EXISTE_LIQUIDACION", Cia, codMaquina, CodLocal, CodUsuario)

    Exit Function
CtrlErr:
        If Err.Number = 94 Then Err.Description = "La Caja " & codMaquina & " no esta Aperturada "
        Err.Raise Err.Number, "clsDocumento.ValidaExisteLiquidacion", Err.Description
End Function

Public Sub GuiaNew(ByVal Cia As String, ByVal TipDoc As String, numGuia As String)
Dim recImpGuia As oraDynaset
Dim rstConvenio As oraDynaset

'Dim devReferencias As oraDynaset

Dim strGlosaConv  As String
Dim i, j As Integer
Dim dblImpTotal As Double

Dim strNumDocumentoRef As String
Dim strCodTipodocRef As String
Dim dblPctCopago As Double
               
Dim dblIGV As Double
Dim p As Printer
Dim dblMtoTotalRef As Double

'-- Variables de cambio 04/12/2008 --'
Dim dblImpSubTotal As Double
Dim dblMtoTotalRefSinIgv As Double
Dim dblMtoSaldo As Double
Dim dblCoaseguro As Double
Dim dlblCoaseguroConIGV As Double
Dim dblImpIgv As Double
Dim dblPrecio As Double
Dim objproducto As New clsProducto
Dim objConvenio As New clsConvenio
Dim strFlgImp As String

 On Error GoTo ErrorImpresora

    For Each p In Printers
       If UCase(p.PORT) = "LPT1:" Then
          Set Printer = p
          Exit For
       End If
    Next p
    
    Dim odynCab As oraDynaset
    Set odynCab = gclsOracle.FN_Cursor("BTLCERO.PKG_GUIA_CLIENTE.FN_CABECERA", 0, numGuia, objUsuario.CodigoLocal)
    If odynCab.RecordCount = 0 Then MsgBox "No se encontró Guía en la BD", vbCritical, "ERROR": Exit Sub
    
    If odynCab("COD_FORMATO").Value = "005" Then 'formato corto
    
        'jrazuri 18/07/2007
        '------------------------------------------
        
        If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
            Dim objImpresion As New clsImpresion
         
            If Not objImpresion.cargaParametros("005") Then
                Exit Sub
            End If
        Else
            
            Printer.Width = 225 * 56.7
            Printer.Height = 152 * 56.7
            
        End If
        
        dblIGV = PctImpuesto(Cia, TipDoc, numGuia)
        
        Set recImpGuia = ImpresionGuia(Cia, numGuia, objUsuario.CodigoLocal)
        
        ' Se agrego esta linea para que realice la impresion del coaseguro segun el documento
        ' por que cuando estaba facturado la guia tomaba omo referencia la factura y no el documento relacionado
        ' Autor : Ccieza
        ' Fecha : 22/04/2010
        '(1) comentado por pherrera 19/07/10 se cambio para que el proceso del financiero no cambie el campo de
        ' documento de referencia y se creo otra columna adicional
        '--------------------------------------------------------------------------------------------------------'
        'Set devReferencias = DocumentosReferencia(Cia, objUsuario.CodigoLocal, TipDoc, NumGuia)
        '--------------------------------------------------------------------------------------------------------'
        
        dblImpTotal = 0
        dblImpSubTotal = 0
        dblMtoTotalRef = 0
        dblMtoTotalRefSinIgv = 0
        dblMtoSaldo = 0
        dblCoaseguro = 0
        dlblCoaseguroConIGV = 0
        dblImpIgv = 0
        dblPrecio = 0
        
        If Not recImpGuia.EOF Then
                    '(1)strNumDocumentoRef = IIf(IsNull(devReferencias("NUM_DOCUMENTO").Value), "", devReferencias("NUM_DOCUMENTO").Value)
                    '(1)strCodTipodocRef = IIf(IsNull(devReferencias("COD_TIPO_DOCUMENTO").Value), "", devReferencias("COD_TIPO_DOCUMENTO").Value)
                    strNumDocumentoRef = IIf(IsNull(recImpGuia!NUM_DOCUMENTO_REF), "", recImpGuia!NUM_DOCUMENTO_REF)
                    strCodTipodocRef = IIf(IsNull(recImpGuia!COD_TIPODOC_REF), "", recImpGuia!COD_TIPODOC_REF)
                    Set rstConvenio = ListaDoc(objUsuario.CodigoEmpresa, strCodTipodocRef, strNumDocumentoRef)
                    dblMtoTotalRef = DevImpDocCnv(objUsuario.CodigoEmpresa, numGuia, TipDoc, objUsuario.CodigoLocal)
                    
                    strGlosaConv = DevTextoDocCnv(objUsuario.CodigoEmpresa, "" & recImpGuia("COD_TIPODOC_REF").Value, "" & recImpGuia("NUM_DOCUMENTO_REF").Value, , objUsuario.CodigoLocal)
                    
                    If strGlosaConv = "*" Then
                        strGlosaConv = UCase(objCliente.Nombre("" & recImpGuia("COD_BENEFICIARIO").Value))
                    End If
                    
                    Set objCliente = Nothing
                    
                    dblPctCopago = IIf(IsNull(recImpGuia("PCT_COPAGO").Value), 0, recImpGuia("PCT_COPAGO").Value)
                    'dblImpTotal = recImpGuia("CAB_NETO").Value
                    Printer.FontName = "Draft 17cpi"
                    Printer.Print "."
                    Printer.Print
                    Printer.Print
                    Printer.Print Space$(61) & "Nueva Direccion: " & CadenaDirecc_Part3(gclsOracle.FN_Valor("CMR.PKG_LOCAL.FN_DIR", objUsuario.CodigoLocal))
                    'Printer.Print
                    'Printer.Print
                    Printer.Print Space$(61) & CadenaDirecc_Part4(gclsOracle.FN_Valor("CMR.PKG_LOCAL.FN_DIR", objUsuario.CodigoLocal))
                    Printer.Print
                    'Printer.Print Space$(13) & recImpGuia("FCH_EMISION").Value
        
                    ' AGREGADO POR DJARA PARA IMPRIMIR EN CASO DE CAMBIO DE DIRECCION
                    If objUsuario.ImprimirDireccion = True Then
                        Printer.Font.name = "Draft 20cpi"
                        Printer.Print Space$(5) & "NUEVA DIR: " + objUsuario.Direccion
                        Printer.Font.name = "Draft 17cpi"
                    Else
                        Printer.Print
                    End If
                    ' DJARA - 03/09/2009
                    
                    Printer.Print Space$(20) & recImpGuia("FCH_EMISION").Value
                    'Printer.Print
                    Printer.Print Space$(60) & recImpGuia("NUM_GUIA").Value
                    Printer.Print Space$(22) & recImpGuia("COD_BENEFICIARIO").Value & "-" & strGlosaConv
                    Printer.Print Space$(22) & recImpGuia("DES_RAZON_SOCIAL").Value
                    Printer.Print Space$(28) & recImpGuia("DIR_ENTREGA").Value & Space(35) & objUsuario.LabelImpConvGuia02 & " " & Trim(recImpGuia("PCT_COPAGO").Value) & "%"
                    Printer.Print Space$(20) & recImpGuia("NUM_DOCUMENTO_ID").Value
                    Printer.Print
                    Printer.Print
                    Printer.Print
                    i = 0
                    Printer.Print Space(63) & CME("Cantidad", 8, "D") & CME("  ", 15, "C") & CME("P.Unit", 15, "D") & CME("  ", 3, "C") & CME("Imp.Neto", 15, "D")
                    While Not recImpGuia.EOF
                
                        'Impresión del precio verdadero del convenio total x producto cuando es co-pago
                        '04/12/2008 Cristhian
                        If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio And (rstConvenio("PCT_BENEFICIARIO").Value > 0 And rstConvenio("PCT_BENEFICIARIO").Value < 100) Then
                                dblImpSubTotal = Format((FnDev_DetPrecioCnv(objUsuario.CodigoEmpresa, _
                                                                          strNumDocumentoRef, _
                                                                          strCodTipodocRef, _
                                                                          recImpGuia("COD_PRODUCTO").Value)), "###,##0.00")
                                dblImpTotal = dblImpTotal + dblImpSubTotal
                          Else
                                dblImpTotal = recImpGuia("CAB_NETO").Value
                                dblImpSubTotal = Format(recImpGuia("DET_NETO").Value, "###,##0.00")
                        End If
                        
                        'Devuelve el precio del producto para poder imprimirlo en la guía
                        '04/12/2008 Cristhian
                        
                        'Corrección en el cáculo del precio final para que distinga el producto exonerado y ademas tome le precio desde la guía grabada en la tabla y tome el precio en línea
                        '02/02/2010 Cristhian R.
                        
                        dblPrecio = objproducto.Precio(objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, recImpGuia("COD_PRODUCTO").Value, "001") / (1 + recImpGuia("IGV_NEW").Value)
                        dblPrecio = dblImpSubTotal / recImpGuia("CTD_UNI_NEW").Value
                        
                        If "" & objVenta.ParametroValor("FLGPRCFRAC") = "1" Then
                                dblPrecio = dblPrecio / Val("" & recImpGuia("PRC_CTD_UNIFICADA").Value)
                                If strNumDocumentoRef = "" And strCodTipodocRef = "" Then
                                       If objConvenio.FnDev_ImprimeImp("" & odynCab("COD_CONVENIO").Value) = "1" Then
                                                Printer.Print Space$(5) & CME(recImpGuia("COD_PRODUCTO").Value, 8, "I") & _
                                                                    CME(recImpGuia("DES_PRODUCTO").Value, 50, "I") & _
                                                                    CME(recImpGuia("CTD_UNIFICADA").Value & " " & recImpGuia("TIPO_CTD_UNIFICADA").Value, 8, "I") & _
                                                                    CME("  ", 15, "C") & _
                                                                    CME(Format(dblPrecio, "###,##0.000"), 15, "D") & _
                                                                    CME(Format(dblImpSubTotal, "###,##0.00"), 18, "D")
                                                                    
                                          Else
                                               Printer.Print Space$(5) & CME(recImpGuia("COD_PRODUCTO").Value, 8, "I") & _
                                                                    CME(recImpGuia("DES_PRODUCTO").Value, 50, "I") & _
                                                                    CME(recImpGuia("CTD_UNIFICADA").Value & " " & recImpGuia("TIPO_CTD_UNIFICADA").Value, 8, "I")
                                       End If
                                  Else
                                       If objConvenio.FnDev_ImprimeImp(odynCab("COD_CONVENIO").Value) = "1" Then
                                                Printer.Print Space$(5) & CME(recImpGuia("COD_PRODUCTO").Value, 8, "I") & _
                                                                        CME(recImpGuia("DES_PRODUCTO").Value, 50, "I") & _
                                                                        CME(recImpGuia("CTD_UNIFICADA").Value & " " & recImpGuia("TIPO_CTD_UNIFICADA").Value, 8, "I") & _
                                                                        CME("  ", 15, "C") & _
                                                                        CME(Format(dblPrecio, "###,##0.000"), 15, "D") & _
                                                                        CME(Format(dblImpSubTotal, "###,##0.00"), 18, "D")
                                           Else
                                                Printer.Print Space$(5) & CME(recImpGuia("COD_PRODUCTO").Value, 8, "I") & _
                                                                        CME(recImpGuia("DES_PRODUCTO").Value, 50, "I") & _
                                                                        CME(recImpGuia("CTD_UNIFICADA").Value & " " & recImpGuia("TIPO_CTD_UNIFICADA").Value, 8, "I")
                                        End If
                                End If
                        
                        
                        Else
                                If strNumDocumentoRef = "" And strCodTipodocRef = "" Then
                                       If objConvenio.FnDev_ImprimeImp(odynCab("COD_CONVENIO").Value) = "1" Then
                                                Printer.Print Space$(5) & CME(recImpGuia("COD_PRODUCTO").Value, 8, "I") & _
                                                                    CME(recImpGuia("DES_PRODUCTO").Value, 50, "I") & _
                                                                    CME(recImpGuia("STOCK").Value, 8, "I") & _
                                                                    CME("  ", 15, "C") & _
                                                                    CME(Format(dblPrecio, "###,##0.00"), 15, "D") & _
                                                                    CME(Format(dblImpSubTotal, "###,##0.00"), 18, "D")
                                                                    
                                          Else
                                               Printer.Print Space$(5) & CME(recImpGuia("COD_PRODUCTO").Value, 8, "I") & _
                                                                    CME(recImpGuia("DES_PRODUCTO").Value, 50, "I") & _
                                                                    CME(recImpGuia("STOCK").Value, 8, "I")
                                       End If
                                  Else
                                       If objConvenio.FnDev_ImprimeImp(odynCab("COD_CONVENIO").Value) = "1" Then
                                                Printer.Print Space$(5) & CME(recImpGuia("COD_PRODUCTO").Value, 8, "I") & _
                                                                        CME(recImpGuia("DES_PRODUCTO").Value, 50, "I") & _
                                                                        CME(recImpGuia("STOCK").Value, 8, "I") & _
                                                                        CME("  ", 15, "C") & _
                                                                        CME(Format(dblPrecio, "###,##0.00"), 15, "D") & _
                                                                        CME(Format(dblImpSubTotal, "###,##0.00"), 18, "D")
                                           Else
                                                Printer.Print Space$(5) & CME(recImpGuia("COD_PRODUCTO").Value, 8, "I") & _
                                                                        CME(recImpGuia("DES_PRODUCTO").Value, 50, "I") & _
                                                                        CME(recImpGuia("STOCK").Value, 8, "I")
                                        End If
                                End If
                        End If
                        'Impresión de la partida arancelaria
                        'jlopez
                        '06/05/2008
                        If dblIGV = 0 Then
                           If Not IsNull(recImpGuia("NUM_PART_ARANCEL").Value) Then
                              Printer.Print "        P.A. : " & recImpGuia("NUM_PART_ARANCEL").Value
                              i = i + 1
                           End If
                        End If
                        
                        recImpGuia.MoveNext
                        i = i + 1
                        
                    Wend
                    recImpGuia.MoveFirst
                    'Printer.Print Space$(5) & CME("---------", 115, "D")
                    
                    
                    'If strNumDocumentoRef <> "" And strCodTipodocRef <> "" Then  --- Se comento para que no valide por los numeros y tipos de referencia 07/04/2009
                        If objConvenio.FnDev_ImprimeImp("" & odynCab("COD_CONVENIO").Value) = "1" Then
                            Printer.Print Space$(5) & CME("---------", 115, "D")
                            
                            Printer.Print Space$(95) & "Sub-Total" & CME(Format(dblImpTotal, "###,##0.00"), 16, "D")
                        
                            '-- Muestra el Coaseguro sin IGV --'
                            If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio Then
                                 'If rstConvenio("FLG_DOC_COPAGO").Value = "1" Then
                                 If rstConvenio("PCT_BENEFICIARIO").Value > 0 And rstConvenio("PCT_BENEFICIARIO").Value < 100 Then
                                      'dblCoaseguro = Val(rstConvenio("COASEGURO").Value)
                                      'dlblCoaseguroConIGV = Val(rstConvenio("COASEGURO_CONIGV").Value)
                                      dblCoaseguro = Val(dblImpTotal) - Val(recImpGuia("CAB_NETO").Value)
                                      If recImpGuia("FLG_DOC_COPAGO").Value = "0" Then
                                        Printer.Print Space$(95) & "Coaseguro" & CME("(" & (Format(dblCoaseguro, "###,##0.00") * -1) & ")", 17, "D")
                                      End If
                                 End If
                            End If
                                 
                             dblMtoSaldo = dblImpTotal - dblCoaseguro
                            ' Printer.Print Space$(5) & "Saldo" & CME(Format(dblMtoSaldo, "###,##0.00"), 110, "D")
                             'Printer.Print
                             'Printer.Print Space$(95) & "SUB-TOTAL" & CME(Format(dblMtoSaldo, "###,##0.00"), 16, "D")
                             Printer.Print Space$(95) & "SUB-TOTAL" & CME(Format(recImpGuia("CAB_NETO").Value, "###,##0.00"), 16, "D")
                                                 
                             If Val(recImpGuia("IMPUESTO_GUIA").Value) > 0 Then
                                 dblImpIgv = Val(recImpGuia("IMPUESTO_GUIA").Value) '+ Val(rstConvenio("IMPUESTO_DOC").Value)
                             Else
                                 dblImpIgv = 0
                             End If
                             
                             Printer.Print Space$(95) & "IGV" & CME(Format(dblImpIgv, "###,##0.00"), 22, "D")
                             
                             Printer.Print Space$(5) & CME("---------", 115, "D")
                             Printer.Print Space$(95) & "TOTAL" & CME(Format(recImpGuia("MTO_TOTAL").Value, "###,##0.00"), 20, "D") '-- se subio 14/01/2008 PSALAZAR
                        End If
                     'End If
                     
                     
                    'Printer.Print Space$(95) & "TOTAL" & CME(Format(dblMtoSaldo + dblImpIgv, "###,##0.00"), 20, "D")
                    
                        
                    For j = 1 To 15 - i
                        Printer.Print
                    Next
                    
                    If strNumDocumentoRef <> "" And strCodTipodocRef <> "" Then
                        '' Se movio hacia arriba --- Set rstConvenio = ListaDoc(objUsuario.CodigoEmpresa, strCodTipodocRef, strNumDocumentoRef)
                        '' Se movio hacia arriba --- dblMtoTotalRef = 0
                        '' Se movio hacia arriba --- dblMtoTotalRef = DevImpDocCnv(objUsuario.CodigoEmpresa, NumGuia, TipDoc, objUsuario.CodigoLocal)
                        
                        If Not rstConvenio.EOF Then
                             strGlosaConv = "" & DevTextoDocCnv(objUsuario.CodigoEmpresa, TipDoc, numGuia, , objUsuario.CodigoLocal)
                             If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio Then
                                 If rstConvenio("FLG_DOC_COPAGO").Value = "1" Then
                                     'Printer.Print "I N S T I T U C I O N  - " & strGlosaConv & " (" & 100 - rstConvenio("PCT_BENEFICIARIO").Value & "%)"
                                     Printer.Print "I N S T I T U C I O N  - " & strGlosaConv & " (" & 100 - rstConvenio("PCT_BENEFICIARIO_ORIG").Value & "%)"
                                     Printer.Print "#REF " & strCodTipodocRef & " " & strNumDocumentoRef & " (" & rstConvenio("PCT_BENEFICIARIO_ORIG").Value & "%) - S/.: " & Format(dblMtoTotalRef, "###,##0.00") & "- BENEF " & DevTextoDocCnv(objUsuario.CodigoEmpresa, TipDoc, numGuia, 1, objUsuario.CodigoLocal)
                                 Else
                                     Printer.Print "B E N E F I C I A R I O  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%)"
                                     Printer.Print "#REF " & strCodTipodocRef & " " & strNumDocumentoRef & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%) - S/.: " & Format(dblMtoTotalRef, "###,##0.00") & "- BENEF " & DevTextoDocCnv(objUsuario.CodigoEmpresa, TipDoc, numGuia, 1, objUsuario.CodigoLocal)
                                 End If
                                 'Printer.Print "#REF " & strCodTipodocRef & " " & strNumDocumentoRef & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%) - S/.: " & Format(dblMtoTotalRef, "###,##0.00") & "- BENEF " & DevTextoDocCnv(objUsuario.CodigoEmpresa, TipDoc, NumGuia, 1, objUsuario.CodigoLocal)
    
                                 If rstConvenio("COD_CONVENIO") = objUsuario.CodConvenioSenati Then
                                    Printer.Print Space$(5) & CME("Crédito Senati: Institución 80% : S/. ", 37, "I") & CME(Round(dblMtoTotalRef * 0.8, 2), 50, "I")
                                    Printer.Print Space$(5) & CME("Crédito Senati: Beneficiario 20%: S/. ", 37, "I") & CME(Round(dblMtoTotalRef * 0.2, 2), 50, "I")
                                 End If
                             End If

                        End If
                    End If
    
                    '**********************************************
                    '**********************************************
                    '**********************************************
                        Printer.FontSize = 10
                        Dim objDocumento As New clsDocumento
                        Dim rstTipoCampo As oraDynaset
                        
                        Set rstTipoCampo = objDocumento.ListaTipoCampo(Cia, TipDoc, numGuia, objUsuario.CodigoLocal)
                        
                        With rstTipoCampo
                          i = 0
                          If Not .EOF Then
                              While Not .EOF
                                  Printer.Print Space(10) & rstTipoCampo("DES_TIPO_CAMPO").Value & " " & rstTipoCampo("DES_VALOR").Value
                                  i = i + 1
                                  .MoveNext
                              Wend
                          End If
                        End With
                        Set objDocumento = Nothing
                        
                        If Not "" & odynCab("COD_CONVENIO").Value = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_CNV_RIMAC") Then  ''CAMBIAR POR CONSTANTE EN RIMAC
                                Printer.Print DevNomRepartidor(objUsuario.CodigoEmpresa, _
                                                   objUsuario.CodigoLocal, _
                                                   TipDoc, _
                                                   numGuia)
                        End If
                        If "" & odynCab("COD_CONVENIO").Value = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_CNV_PLAVITAL") Then  ''CAMBIAR POR CONSTANTE EN PLAN VITAL
                                Printer.Print DevTipoFacturacion(objUsuario.CodigoEmpresa, _
                                                   objUsuario.CodigoLocal, _
                                                   TipDoc, _
                                                   numGuia)
                        End If
                        
                        'Dim objConvenio As New clsConvenio
                        Dim rsCabeceraReceta As oraDynaset
                        Set rsCabeceraReceta = objConvenio.ListaCabeceraReceta(TipDoc, numGuia)
                        
                        Dim o As Integer
                        'este se agrego
                        Do Until rsCabeceraReceta.EOF
                            While o < rsCabeceraReceta.Fields.Count
                                Printer.Print Space(10) & rsCabeceraReceta.Fields(o).name & ": " & rsCabeceraReceta(o).Value
                                o = o + 1
                            Wend
                            rsCabeceraReceta.MoveNext
                        Loop
    
                        Printer.FontName = "Draft 12cpi"
                        Set objConvenio = Nothing
                    '**********************************************
                    '**********************************************
                    '**********************************************
    
                    For j = 1 To 4
                        Printer.Print
                    Next
    
                    Printer.FontName = "Draft 17cpi"
    
                    Printer.Print CME("X", 135, "D")
                    Printer.Print CME(objUsuario.LabelImpConvGuia01, 135, "D")
                    Dim strMensaje  As String
                    strMensaje = DevMensaje(objUsuario.CodigoEmpresa, numGuia, "GRL", objUsuario.CodigoLocal)
                    Printer.Print Space(50) + IIf(IsNull(strMensaje), "", strMensaje)
                    Printer.EndDoc
    
            End If
            
        Set rstConvenio = Nothing
        Set recImpGuia = Nothing
    Else
        'IMPRESION LARGA
        ImprimeGuiaLarga odynCab, numGuia
    End If
        Exit Sub
    
ErrorImpresora:
    Err.Raise Err.Number, "clsDocumento.Guia", Err.Description
End Sub

Function FnDev_DetPrecioCnv(ByVal vstrCia As String, _
                            ByVal vstrNumDoc As String, _
                            ByVal vstrTipDoc As String, _
                            ByVal vstrCodProducto As String) As Double
 On Error GoTo CtrlErr
    FnDev_DetPrecioCnv = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO.FN_DEV_PRECIO_CNV", vstrCia, _
                                                                                       vstrNumDoc, _
                                                                                       vstrTipDoc, _
                                                                                       vstrCodProducto)

  Exit Function
CtrlErr:
  Err.Raise Err.Number, "clsDocumento", Err.Description
End Function

Private Sub FacturaNew(ByVal TipoDocumento As String, ByVal numDocumento As String)

Dim rstTipoCampo As oraDynaset
Dim rstImpresion As oraDynaset
Dim rstDocumentoMaquina As oraDynaset
Dim rstDonacion As oraDynaset
Dim rstConvenio As oraDynaset
Dim objDocumento As New clsDocumento

Dim dblIGV As Double

Dim strEntidadDonacion As String

Dim i As Integer
Dim j As Integer

Dim arrMeses(12) As String
Dim intNumeroLineas As Integer
Dim dblDonacion As Double
Dim strLetras As String
Dim dblMtoBaseImp As Double
Dim dblMtoExonerado As Double
Dim dblMtoImpuesto As Double
Dim dblMtoTotal As Double
Dim strMensaje As String
Dim intTamañoEspacio  As Integer
Dim intFormato  As Integer
Dim objImpresiones As New clsProforma
Dim p As Printer
Dim strGlosaConv As String
Dim dblMtoTotalRef As Double
'**Modified Martinnets
Dim rsFormato As oraDynaset
Dim objFormato As New clsImpresion
Dim lsFormato As String
'**************
Dim strDesAuxCliNombre As String
Dim strCodMaquina As String
Dim strCodUsuarioDependiente As String
Dim strCodConvenio As String

Dim dblCoaseguro As Double
Dim dblMtoSubtotal As Double

   arrMeses(1) = "Enero"
   arrMeses(2) = "Febrero"
   arrMeses(3) = "Marzo"
   arrMeses(4) = "Abril"
   arrMeses(5) = "Mayo"
   arrMeses(6) = "Junio"
   arrMeses(7) = "Julio"
   arrMeses(8) = "Agosto"
   arrMeses(9) = "Setiembre"
   arrMeses(10) = "Octubre"
   arrMeses(11) = "Noviembre"
   arrMeses(12) = "Diciembre"
   
   On Error GoTo ErrorImpresora
   
    For Each p In Printers
       If UCase(p.PORT) = "LPT1:" Then
          Set Printer = p
          Exit For
       End If
    Next p

    'COMENTADO POR JLOPEZ PARA VERIFICAR LA IMPRESION EN XP

'    Set rsFormato = ListaFormato(objUsuario.CodigoEmpresa, TipoDocumento, NumDocumento)
'    lsFormato = rsFormato.Fields("COD_FORMATO")
'    objFormato.cargaParametros lsFormato

    Set rstConvenio = ListaDoc(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
    If Not rstConvenio.EOF Then strCodConvenio = "" & rstConvenio("COD_CONVENIO").Value
   
    Set rstImpresion = objDocumento.Impresion(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)

    If rstImpresion("COD_FORMATO").Value = "002" Then 'factura corta
   
       If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
    '        intFormato = objImpresiones.CrearFormato("Formato BTL - 3", 225, 152)
    '        Set objImpresiones = Nothing
    '        Printer.PaperSize = intFormato
            If Not objFormato.cargaParametros("002") Then
                Exit Sub
            End If
             ': Printer.Print
       Else
            Printer.Width = 225 * 56.7
            Printer.Height = 153 * 56.7 '152*56.7, ' 93.3 * 56.7 ' 280
       End If
       
       Printer.Print: Printer.Print
       Printer.FontName = "Draft 17cpi"
       Printer.Print "."
       Printer.FontName = "Draft 10cpi"
       Printer.Print Space(55) & "NIT " & objUsuario.Ruc
       Printer.FontName = "Draft 12cpi"
    '   Printer.Print
    '   Printer.Print
    
          dblIGV = objDocumento.PctImpuesto(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
          Set rstTipoCampo = objDocumento.ListaTipoCampo(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento, objUsuario.CodigoLocal)
    
          With rstTipoCampo
            i = 0
            If Not .EOF Then
                While Not .EOF
                    Printer.Print Space(28) & rstTipoCampo("DES_TIPO_CAMPO").Value & " " & rstTipoCampo("DES_VALOR").Value
                    i = i + 1
                    .MoveNext
                Wend
            End If
    
    '      i = i + 1
    
          End With
    '      If i < 4 Then
    '        For j = i To 4
    '           Printer.Print
    '        Next j
    '      End If
    
        For j = 1 To 4 - i
            Printer.Print
        Next j
    
        intNumeroLineas = 0
        dblDonacion = 0
        'Set rstDocumentoMaquina = objDocumento.ListaDocumentoMaquina(objUsuario.CodigoEmpresa, TipoDocumento, objUsuario.NombrePC)
        'If Not rstDocumentoMaquina.EOF Then intNumeroLineas = rstDocumentoMaquina("NUM_LINEAS_DOC").Value
        
        Dim objFormDoc As New clsFormato
        intNumeroLineas = objFormDoc.NumLineas("002")
        Set objFormDoc = Nothing
        
        dblMtoBaseImp = 0
        dblMtoExonerado = 0
        dblMtoImpuesto = 0
        dblMtoTotal = 0
        dblCoaseguro = 0
        dblMtoSubtotal = 0
        
        strCodMaquina = rstImpresion("COD_MAQUINA").Value
        strCodUsuarioDependiente = rstImpresion("COD_USUARIO_DEPENDIENTE").Value
        
        '*****************************************************************************'
        '***** Datos de proforma para imprimir cuando el la venta
        '***** se haya originado por delivery aqui saldra impreso la BOL/FAC con PRO
        '***** Hecho el 19/07/2007
    
         objVenta.NumPedidoPadre = "" & rstImpresion("NUM_PEDIDO_PADRE").Value
         
        '******************************************************************************'
    
        If Not rstImpresion.EOF Then
            'Printer.FontName = "Draft 12cpi"
            Printer.Print
            'Printer.Print Space(18) & rstImpresion("COD_MAQUINA").Value & "   " & rstImpresion("COD_USUARIO_DEPENDIENTE").Value & "   " & rstImpresion("COD_USUARIO").Value
            ' AGREGADO POR DJARA PARA IMPRIMIR EN CASO DE CAMBIO DE DIRECCION
            If objUsuario.ImprimirDireccion = True Then
                Printer.Font.name = "Draft 20cpi"
'                Printer.Print Space(16) & "NUEVA DIR: " + left(objUsuario.direccion, AnchoFactura - 11)
                Printer.Print "" & Mensaje1(objUsuario.CodigoLocal, "FAC")
                Printer.Print "" & Mensaje2(objUsuario.CodigoLocal, "FAC")
                Printer.Font.name = "Draft 17cpi"
            Else
                Printer.Print
            End If
            ' DJARA - 03/09/2009
            
            Printer.Print Space(8) & CStr(Day(rstImpresion("FCH_REGISTRA").Value)) & " de " & arrMeses(Month(rstImpresion("FCH_REGISTRA").Value)) & " de " & CStr(Year(rstImpresion("FCH_REGISTRA").Value))
            Printer.Print Space(8) & left(rstImpresion("DES_RAZON_SOCIAL").Value & Space(60), 60) & " " & rstImpresion("NUM_DOCUMENTO").Value
            Printer.Print Space(8) & rstImpresion("NUM_RUC_EMPRESA").Value & Space(10) & IIf(strCodConvenio = objUsuario.CodConvenioScotiaBank, "Crédito 30 días", "") & Space(10) & rstImpresion("NUM_OC").Value
            Printer.Print Space(8) & rstImpresion("DES_AUX_CLI_DIRECC").Value
    
            Printer.Print
            'Printer.Print
            'Printer.Print
            'Printer.FontName = "Draft 10cpi"
            strLetras = rstImpresion("LETRAS").Value
            'dblMtoBaseImp = rstImpresion("MTO_NETO_SINIGV").Value
            dblMtoExonerado = rstImpresion("MTO_EXONERADO").Value
            dblMtoImpuesto = rstImpresion("MTO_IMPUESTO").Value
            'dblMtoTotal = rstImpresion("MTO_NETO_SINIGV").Value + rstImpresion("MTO_EXONERADO").Value
            
            'modificado por jmelgar, en caso DES_AUX_CLI_NOMBRE es nulo busca el nombre del beneficiario en clientes
            strDesAuxCliNombre = IIf("" & rstImpresion("DES_AUX_CLI_NOMBRE").Value = "", UCase(objCliente.Nombre("" & rstImpresion("COD_CLIENTE").Value)), "" & rstImpresion("DES_AUX_CLI_NOMBRE").Value)
           
            Printer.Print
            Printer.FontName = "Draft 17cpi"
            While Not rstImpresion.EOF
            
                dblMtoBaseImp = dblMtoBaseImp + rstImpresion("MTO_SUBTOTAL_SINIGV").Value
                'Printer.Print right(Space(7) & Trim(rstImpresion("CANTIDAD").Value), 7) & Space(5) & left(rstImpresion("COD_PRODUCTO").Value & " " & rstImpresion("DES_PRODUCTO").Value & Space(40), 40) & _
                            Space(2) & right(Space(8) & Format(rstImpresion("PRC_UNIT_VTA").Value, "#####0.#0"), 8) & _
                           Space(6) & right(Space(8) & Format(rstImpresion("SUB_TOTAL").Value, "#####0.#0"), 8) & " " & right(Space(6) & Format(rstImpresion("PCT_DSCT_UNIT").Value, "#####0.#0"), 6) & _
                            " " & right(Space(8) & Format(rstImpresion("TOTAL").Value, "#####0.#0"), 8)
                Printer.Print right(Space(7) & Trim(rstImpresion("CANTIDAD").Value), 7) & Space(2) & left(rstImpresion("COD_PRODUCTO").Value & " " & rstImpresion("DES_PRODUCTO").Value & Space(40), 40) & _
                            Space(2) & right(Space(20) & IIf(IsNull(rstImpresion("NUM_LOTE").Value), "", rstImpresion("NUM_LOTE").Value), 20) & _
                            Space(2) & right(Space(10) & IIf(IsNull(rstImpresion("FCH_VENCIMIENTO").Value), "", rstImpresion("FCH_VENCIMIENTO").Value), 10) & Space(2) & right(Space(8) & Format(rstImpresion("PRC_UNIT_VTA_SINIGV").Value, "#####0.#0"), 8) & _
                            Space(2) & right(Space(10) & Format(rstImpresion("TOTAL_SIN_IGV").Value, "#####0.#000"), 10) & Space(2) & right(Space(6), 6) & _
                            Space(2) & right(Space(8) & Format(rstImpresion("MTO_SUBTOTAL_SINIGV").Value, "#####0.#0"), 8)

                If dblIGV = 0 Then
                   If Not IsNull(rstImpresion("NUM_PART_ARANCEL").Value) Then
                      Printer.Print "        P.A. : " & rstImpresion("NUM_PART_ARANCEL").Value
                      i = i + 1
                   End If
                End If
    
               rstImpresion.MoveNext
               i = i + 1

            Wend
            Printer.FontName = "Draft 12cpi"
           For j = 1 To intNumeroLineas - i
               Printer.Print
           Next j
           Printer.FontName = "Draft 17cpi"
           
           rstImpresion.MoveFirst
           dblCoaseguro = objDocumento.FnDev_BaseImpoGuia(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
           If dblCoaseguro = 0 Then
               'dblCoaseguro = objDocumento.FnDev_BaseImpoDoc(objUsuario.CodigoEmpresa, TipoDocumento, NumDocumento)
               dblCoaseguro = Val(dblMtoBaseImp) - Val(rstImpresion("MTO_BASE_IMP").Value)
           End If
           
           
           dblMtoSubtotal = Val(dblMtoBaseImp) - Val(dblCoaseguro)
           dblMtoTotal = Val(dblMtoSubtotal) + Val(dblMtoImpuesto)
            
           Printer.Print Space(97) & "Sub_Total" & Space(10) & Format(dblMtoBaseImp, "#####0.#0")
           
           If rstImpresion("FLG_DOC_COPAGO").Value = "0" Then
               Printer.Print Space(97) & "Coaseguro" & Space(10) & "(" & "" & (Format(dblCoaseguro, "#####0.#0") * -1) & "" & ")"
           End If
           
           dblMtoTotalRef = 0
           dblMtoTotalRef = DevImpDocCnv(objUsuario.CodigoEmpresa, numDocumento, TipoDocumento, objUsuario.CodigoLocal)
           
           Printer.FontName = "Draft 12cpi"
           If Not rstConvenio.EOF Then
                strGlosaConv = DevTextoDocCnv(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento, , objUsuario.CodigoLocal)
                If rstConvenio("COD_MODALIDAD_VENTA").Value = objUsuario.ModalidadVentaConvenio Then
                    If rstConvenio("FLG_DOC_COPAGO").Value = "1" Then
                        'Printer.Print "B E N E F I C I A R I O  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%)"
                        Printer.Print "B E N E F I C I A R I O  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO_ORIG").Value & "%)"
                    Else
                        Printer.Print "I N S T I T U C I O N  - " & strGlosaConv & " (" & rstConvenio("PCT_BENEFICIARIO").Value & "%)"
                    End If
    
                    Printer.Print "#REF " & rstConvenio("COD_TIPO_DOCUMENTO_REFER").Value & " " & rstConvenio("NUM_DOCUMENTO_REFER").Value & " (" & 100 - rstConvenio("PCT_BENEFICIARIO").Value & "%) - S/.: " & Format(IIf(dblMtoTotalRef < 0, 0, dblMtoTotalRef), "###,##0.00") & IIf(strDesAuxCliNombre = "", "", "   BENEFICIARIO: " & strDesAuxCliNombre)
    
                    Printer.Print DevNomRepartidor(objUsuario.CodigoEmpresa, _
                                          objUsuario.CodigoLocal, _
                                          TipoDocumento, _
                                          numDocumento)
                End If
           Else
                '' AQUI VA EL TEXTO DE LA PROMOCION
               Printer.Print
               Printer.Print
           End If
    
           Set rstDonacion = objDocumento.Impresion(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento, "1")
    
           If rstDonacion.EOF Then
                Printer.Print
           Else
                dblDonacion = 0
                While Not rstDonacion.EOF
                   dblDonacion = dblDonacion + rstDonacion("IMP_CON_REDONDEO").Value
                   strEntidadDonacion = strEntidadDonacion & left(rstDonacion("DES_HIJO").Value, 54) & "   " & Trim(objUsuario.SmbMoneda) & right("        " & Format(rstDonacion("IMP_CON_REDONDEO").Value, "####0.#0"), 8) & " "
                   rstDonacion.MoveNext
                Wend
    
                Printer.Print "     Usted está donando " & Trim(objUsuario.SmbMoneda) & Trim(Format(dblDonacion, "#####0.#0")); " a " & strEntidadDonacion
           End If
    
        '** Datos del Repartidor de tener **'
    
        '**
        Printer.Print "     SON : " & strLetras & " " & Trim(objUsuario.DesLgMoneda)
    
    
        Printer.Print
    '    Printer.Print
    '    Printer.Print
    '    Printer.Print
    '    Printer.Print Space(40) & "%  " & right(Space(11) & Format(dblIGV * 100, "##0.00"), 11) '** COMENTANDO SEGUN NUEVO FORMATO DE IMPRESION EN CONVENIOS 10/11/2009 CRUEDA
'        Printer.Print Space(8) & Trim(objUsuario.SmbMoneda) & right(Space(11) & _
'                        Format(IIf(dblMtoBaseImp > 0, dblMtoSubTotal, dblMtoExonerado), "#### ##0.#0"), 11) & Space(18) & _
'                        Trim(objUsuario.SmbMoneda) & right(Space(11) & Format(dblMtoImpuesto, "#### ##0.#0"), 11) & Space(25) & _
'                        Trim(objUsuario.SmbMoneda) & right(Space(11) & Format(dblMtoTotal, "#### ##0.#0"), 11)
            
        Printer.Print Space(8) & Trim(objUsuario.SmbMoneda) & right(Space(11) & _
                        Format(IIf(rstImpresion("MTO_BASE_IMP").Value > 0, rstImpresion("MTO_BASE_IMP").Value, rstImpresion("MTO_EXONERADO").Value), "#### ##0.#0"), 11) & Space(18) & _
                        Trim(objUsuario.SmbMoneda) & right(Space(11) & Format(rstImpresion("MTO_IMPUESTO").Value, "#### ##0.#0"), 11) & Space(25) & _
                        Trim(objUsuario.SmbMoneda) & right(Space(11) & Format(rstImpresion("MTO_TOTAL").Value, "#### ##0.#0"), 11)
            
        Printer.Print
        'Printer.Print
        strMensaje = rstImpresion("COD_MAQUINA").Value & "   " & rstImpresion("COD_USUARIO_DEPENDIENTE").Value & "   " & rstImpresion("COD_USUARIO").Value
        Printer.Print Space(AnchoFactura - (Len(strMensaje) + 1)) & strMensaje
        strMensaje = strCodMaquina + "   " + strCodUsuarioDependiente
        Printer.Print Space(AnchoFactura - (Len(strMensaje) + 1)) & strMensaje
        
        ''AQUI VA EL MENSAJE
    
        strMensaje = DevMensaje(objUsuario.CodigoEmpresa, numDocumento, TipoDocumento, objUsuario.CodigoLocal)
        Dim rsNewMess As oraDynaset
        Set rsNewMess = ListaMensajes
        While Not rsNewMess.EOF
            Printer.FontSize = 7
            Printer.Print "" & rsNewMess("DES_PARAMETRO").Value
            rsNewMess.MoveNext
        Wend
        'intTamañoEspacio = Int((AnchoFactura - Len(IIf(IsNull(strMensaje), "", strMensaje))) / 2)
        'Printer.Print Space(intTamañoEspacio) + IIf(IsNull(strMensaje), "", strMensaje)
        Printer.Print Space(AnchoFactura - (Len(strMensaje) + 1)) + IIf(IsNull(strMensaje), "", strMensaje)
        
'        ListaMensajes
        
        Printer.EndDoc
        
        Else
            MsgBox "No se encontraron datos en el documento " + Chr(13) + "F/ " & rstImpresion("NUM_DOCUMENTO").Value, vbYesNo + vbDefaultButton1
        End If
    Else
        ImprimeFacLarga objUsuario.CodigoEmpresa, numDocumento ' impresion de factura larga
    End If
    Set rstImpresion = Nothing
   Exit Sub
   
ErrorImpresora:
    Err.Raise Err.Number, "clsImpresiones.Boleta", Err.Description

End Sub

Function FnDev_BaseImpoGuia(ByVal vstrCia As String, _
                            ByVal vstrTipDocRef As String, _
                            ByVal vstrNumDocRef As String) As Double
 On Error GoTo CtrlErr
    FnDev_BaseImpoGuia = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO.FN_DEV_BASE_IMPONIBLE_GUIA", vstrCia, _
                                                                                                 vstrTipDocRef, _
                                                                                                 vstrNumDocRef)

  Exit Function
CtrlErr:
  Err.Raise Err.Number, "clsDocumento", Err.Description
End Function

Function FnDev_BaseImpoDoc(ByVal vstrCia As String, _
                           ByVal vstrTipDocRef As String, _
                           ByVal vstrNumDocRef As String) As Double
 On Error GoTo CtrlErr
    FnDev_BaseImpoDoc = gclsOracle.FN_Valor("BTLPROD.PKG_DOCUMENTO.FN_DEV_BASE_IMPONIBLE_DOC", vstrCia, _
                                                                                                vstrTipDocRef, _
                                                                                                vstrNumDocRef)

  Exit Function
CtrlErr:
  Err.Raise Err.Number, "clsDocumento", Err.Description
End Function

Public Function MensajeImpresionConvenio(ByVal CodigoConvenio As String) As String

On Error GoTo CtrlErr

    MensajeImpresionConvenio = "" & gclsOracle.FN_Valor("BTLPROD.PKG_CONVENIO.FN_MENSAJE_TICKET", CodigoConvenio)

    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.MensajeImpresionConvenio", Err.Description
End Function

Private Sub Class_Terminate()
    Set objMensaje = Nothing
End Sub
Public Function ListaMotivoNCDet(ByVal CodigoSubConcepto As String) As oraDynaset
    On Error GoTo CntrlErr
    
    Set ListaMotivoNCDet = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_LISTA_MOTIVO_NC_NEW", 0, CodigoSubConcepto)
    
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaMotivoNCDet", Err.Description
End Function

Public Function DevTipoFacturacion(ByVal vstrCia As String, _
                                   ByVal vstrLocal As String, _
                                   ByVal vstrTipDoc As String, _
                                   ByVal vstrDocumento As String) As String

On Error GoTo CtrlErr
    DevTipoFacturacion = gclsOracle.FN_Valor("BTLPROD.PKG_CARGA_ARCHIVO_CNV.FN_DEV_TIPO_FACTURACION", vstrCia, vstrLocal, vstrTipDoc, vstrDocumento)
    Exit Function
CtrlErr:
        Err.Raise Err.Number, "clsDocumento.DevTipoFacturacion", Err.Description

End Function


Private Function fnDevTipoCambioDoc(ByVal Cia As String, ByVal TipoDocumento As String, ByVal numDocumento As String) As Double
On Error GoTo handle
    Dim istrSQL As String
    Dim rs As oraDynaset
    Screen.MousePointer = vbHourglass
    istrSQL = "SELECT imp_tip_cambio from btlprod.cab_documento WHERE CIA='" & Cia & "' AND num_documento='" & numDocumento & "' AND cod_tipo_documento='" & TipoDocumento & "'"
     Set rs = gclsOracle.ODataBase.CreateDynaset(istrSQL, 0&)
     While Not rs.EOF
            fnDevTipoCambioDoc = Val(rs!IMP_TIP_CAMBIO)
        rs.MoveNext
     Wend
    Screen.MousePointer = vbDefault
    
    Exit Function
handle:
    Err.Raise Err.Number, "clsDocumento.fnDevTipoCambioDoc", Err.Description
    Screen.MousePointer = vbDefault
End Function

'Function fnImprimeCupon(ByVal CodigoDocDescto As String, ByVal NumDocDescto As String)
'    Dim rs As oraDynaset
'    Set rs = ListaCupon(CodigoDocDescto, NumDocDescto)
'    Set imagen.Picture = LoadPicture("c:\logobtl.bmp")
'    Printer.PaintPicture imagen, 1000, 20, 2150, 950
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Font.Size = 12
'    Printer.FontName = "IDAutomation.com Code39"  'LETRA DE CODIGO DE BARRA'
'    Printer.Print Space(5) & "*" & rs("NUM_DOCUMENTO_PAGO") & "*"
'    Printer.FontName = "Courier New"
'    Printer.Font.Size = 15
'    Printer.Font.Bold = True
'    Printer.Print Space(5) & rs("DES_TITULO")
'    Printer.Font.Size = 13
'    Printer.Print Space(5) & rs("DES_SUBTITULO")
'    Printer.Font.Bold = False
'    Printer.Font.Size = 10
'    Printer.Print Space(3) & "Para: WALDIR ZANABRIA"
'    Dim strMensaje As String
'    strMensaje = "" & rs("DES_CONTENIDO")
'    Printer.Font.Size = 8
'    While strMensaje <> ""
'        Printer.Print Space(2) & left(Mid(strMensaje, 1, 37) & Space(37), 37) & Space(2)
'        strMensaje = Mid(strMensaje, 38, Len(strMensaje) - Len(Mid(strMensaje, 1, 37)))
'    Wend
'    Printer.Print Space(3) & rs("NUM_DOCUMENTO_PAGO") & Space(10) & "Nro." & rs("SEC_VENTA_EMI")
'    Printer.Font.Size = 5
'    Printer.Print
'    Printer.Font.Bold = True
'    Printer.Font.Size = 7
'    Printer.Print Space(2) & Chr(34) & "En Boticas BTL nos preocupamos por tu Salud" & Chr(34)
'    Printer.Font.Bold = False
'    Printer.Font.Size = 5
'    Printer.Print Space(3) & "Vigencia del 24/05/11 al 31/12/11.Máximo descuento diario por cliente"
'    Printer.Print Space(3) & "s/.50 o 3 unidades por producto.No es acumulable con otros descuento"
'    Printer.Print Space(3) & "Valido solo en locales"
'    Printer.Print
'    Printer.Font.Size = 6
'    Printer.Print Space(3) & rs("COD_USUARIO") & Space(10) & rs("FCH_REGISTRA") & Space(10) & rs("COD_LOCAL_EMISOR")
'    Printer.Print
'    Printer.EndDoc
'End Function

Public Function ListaCupon(ByVal CodDocDscto As String, ByVal NumDocDscto As String) As oraDynaset
    On Error GoTo CntrlErr
    
    Set ListaCupon = gclsOracle.FN_Cursor("btlprod.pkg_documento_dscto.fn_lista_detalle_c", 0, CodDocDscto, NumDocDscto)
    
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaCupon", Err.Description
End Function

Public Function ListaDatosCupon(ByVal CodDocDscto As String) As oraDynaset
    On Error GoTo CntrlErr
    
    Set ListaDatosCupon = gclsOracle.FN_Cursor("btlprod.fn_datos_cupon", 0, CodDocDscto)
    
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaDatosCupon", Err.Description
End Function

Public Function ListaCupones(ByVal SecVenta As String) As oraDynaset
    On Error GoTo CntrlErr
    Set ListaCupones = gclsOracle.FN_Cursor("btlprod.pkg_documento_dscto.fn_lista_cupones", 0, SecVenta)
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaCupones", Err.Description
End Function

Public Function ListaMensajes() As oraDynaset
    On Error GoTo CntrlErr
    Set ListaMensajes = gclsOracle.FN_Cursor("btlprod.fn_dev_texto", 0)
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaMensajes", Err.Description
End Function

Public Function ListaImpresoraMaquina(ByVal Cia As String, _
                        ByVal TipoDocumento As String, _
                        ByVal codMaquina As String) As String
    On Error GoTo CtrlErr
    
    ListaImpresoraMaquina = "" & gclsOracle.FN_Valor("BTLPROD.PKG_IMPRESION.FN_IMPRESORA_MAQUINA", Cia, TipoDocumento, codMaquina)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaImpresoraMaquina", Err.Description
End Function

Public Function ListaTermico(ByVal strSecventa As String, ByVal strCodLocal As String) As oraDynaset
    On Error GoTo CntrlErr
    
    Set ListaTermico = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO_DSCTO.FN_LISTA_TERMICO", 0, strSecventa, strCodLocal)
    
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaTermico", Err.Description
End Function

Public Function ListaDocCredito(ByVal strSecventa As String) As String
    On Error GoTo CntrlErr
    
     ListaDocCredito = gclsOracle.FN_Valor("BTLPROD.FN_LISTA_DOC_CREDITO", strSecventa)
    
    Exit Function
CntrlErr:
    Err.Raise Err.Number, "clsDocumento.ListaDocCredito", Err.Description
End Function

Public Function FnValidaReimprimir(ByVal vstrCodUsuario As String) As Integer
On Error GoTo CtrlErr
    FnValidaReimprimir = gclsOracle.FN_Valor("BTLPROD.FN_VALIDA_REIMPRIMIR", vstrCodUsuario)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.FnValidaReimprimir", Err.Description
End Function
Public Function NombreComercial(ByVal CodigoLocal As String) As String
On Error GoTo CtrlErr
    NombreComercial = gclsOracle.FN_Valor("btlprod.fn_dev_nom_com", CodigoLocal)
   
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.NombreComercial", Err.Description
End Function

Public Function Mensaje1(ByVal CodigoLocal As String, ByVal TipoDocumento As String) As String
On Error GoTo CtrlErr
    Mensaje1 = gclsOracle.FN_Valor("BTLPROD.FN_IMPRIME_LINEA_1", CodigoLocal, TipoDocumento)
   
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.fn_imprime_linea_1", Err.Description
End Function


Public Function Mensaje2(ByVal CodigoLocal As String, ByVal TipoDocumento As String) As String
On Error GoTo CtrlErr
    Mensaje2 = gclsOracle.FN_Valor("BTLPROD.FN_IMPRIME_LINEA_2", CodigoLocal, TipoDocumento)
   
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.fn_imprime_linea_2", Err.Description
End Function


Public Function Mensaje3(ByVal CodigoLocal As String, ByVal TipoDocumento As String) As String
On Error GoTo CtrlErr
    Mensaje3 = gclsOracle.FN_Valor("BTLPROD.FN_IMPRIME_LINEA_3", CodigoLocal, TipoDocumento)
   
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.fn_imprime_linea_3", Err.Description
End Function

Public Function PORT(ByVal Maquina As String, ByVal TipoDocumento As String) As String
On Error GoTo CtrlErr
    PORT = gclsOracle.FN_Valor("BTLPROD.FN_DEV_PORT_IMP", Maquina, TipoDocumento)
   
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDocumento.PORT", Err.Description
End Function

Private Sub ImprimeCuponMonedero(ByVal TipoDocumento As String, ByVal numDocumento As String, Optional FlgAnulado As Boolean = False)
    'Dim oFPC As New clsFarmaPuntosConstante
    Dim oFP As New clsFarmaPuntos
    Dim oTB As New clsTarjetaBean
    'Dim strMensaje As String
    Dim intConstante As Integer, intX As Integer, dblPuntosDinero As Double
    Dim rst As oraDynaset
    Dim Linea(1 To 9) As String, lineaValor(1 To 6) As Double
    Dim dbl12mahorro As Double
    
    On Error GoTo Control
    
    Screen.MousePointer = vbHourglass
    
    Set rst = Impresion_Ticket(objUsuario.CodigoEmpresa, TipoDocumento, numDocumento)
    
    If IsNull(rst("NUM_TARJ_PUNTOS").Value) Or rst("EST_TRX_ORBIS").Value = "D" Then Exit Sub
    
    Linea(1) = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "INDFV471", objUsuario.CodigoEmpresa) ' AHORRASTE
    Linea(2) = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "INDFV470", objUsuario.CodigoEmpresa) ' HAS CANJEADO
    Linea(3) = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "INDFV498", objUsuario.CodigoEmpresa) ' HAS GANADO - CPN
    Linea(4) = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "INDFV469", objUsuario.CodigoEmpresa) ' HAS GANADO - PTO
    Linea(5) = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "INDFV485", objUsuario.CodigoEmpresa) ' ACUMULADOS
    Linea(6) = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "INDFV499", objUsuario.CodigoEmpresa) ' 12M AHORRO
    Linea(7) = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "INDFV500", objUsuario.CodigoEmpresa)
    Linea(8) = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "INDFV661", objUsuario.CodigoEmpresa)
    Linea(9) = "" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR_DESC", "INDFV458", objUsuario.CodigoEmpresa)
    dbl12mahorro = Val("" & gclsOracle.FN_Valor("NUEVO.PKG_PARAMETRO.FN_VALOR", "INDFV474"))
    
    intConstante = 184
    
    If SetearImpresoraCupon = False Then
        MsgBox "No tiene la impresora de cupones Instalada o tiene diferente nombre", vbCritical, App.ProductName
        Exit Sub
    End If
    For intX = 0 To Printer.FontCount - 1
        Debug.Print Printer.Fonts(intX)
    Next
    
    Set oTB = oFP.ConsultarSaldo(rst("NUM_TARJ_PUNTOS").Value, objUsuario.Codigo)
    
    While Not rst.EOF
        lineaValor(1) = Val(lineaValor(1)) + Val(rst("MTO_AHORRO").Value) ' AHORRASTE
        lineaValor(3) = Val(lineaValor(3)) + Val(rst("CTD_CUPON").Value) ' HAS GANADO - CPN
        rst.MoveNext
    Wend
    rst.MoveFirst
    
    lineaValor(2) = Val(rst("PT_REDIMIDO").Value) ' HAS CANJEADO
    lineaValor(4) = Val(rst("PT_ACUMULADO").Value) ' HAS GANADO - PTO
    lineaValor(5) = Val(rst("PT_TOTAL").Value) ' ACUMULADOS
    lineaValor(6) = Val(oTB.AhorroTotal) * Val(rst("FARMA_PUNTOS_REDIME").Value) ' 12M AHORRO
    
    'Logo
    mdiPrincipal.imgLogoBtl.Picture = mdiPrincipal.ImageList1.ListImages(5).Picture
    Printer.PaintPicture mdiPrincipal.imgLogoBtl, 1100, 20, 2150, 750
    Printer.CurrentY = 1154
    
    'Nombre cliente
    Printer.FontName = "Arial"
    Printer.FontSize = 8
    Printer.FontBold = True
    centra_printer oTB.NombreCompleto
    Printer.Print "Tarjeta: " & EncriptarNumeroTarjeta(oTB.NumeroTarjeta); Tab(35); "DNI/CE: " & oTB.DNI

    Printer.FontName = "Printer FontB 22.5cpi"
    Printer.FontSize = 7
    Printer.FontBold = False
    If Val(lineaValor(1)) > 0.5 Then _
        Printer.Print Linea(1); Tab(40); "    S/." & pfstr_Alineado(Format$(lineaValor(1), "#0.00"), 14, Derecha)
    If Val(lineaValor(2)) > 0 Then _
        Printer.Print Linea(2); Tab(40); pfstr_Alineado(Format$(lineaValor(2), "#0"), 7, Derecha) & " puntos dinero"
    If Val(lineaValor(3)) > 0 Then _
        Printer.Print Linea(3); Tab(40); pfstr_Alineado(Format$(lineaValor(3), "#00"), 7, Derecha) & " Cupon Dscto"
    If Val(lineaValor(4)) > 0 Then _
        Printer.Print Linea(4); Tab(40); pfstr_Alineado(Format$(lineaValor(4), "#0"), 7, Derecha) & " puntos dinero"
    If Val(lineaValor(5)) > 0 Then _
        Printer.Print Linea(5); Tab(40); pfstr_Alineado(Format$(lineaValor(5), "#0"), 7, Derecha) & " puntos dinero"
    If Val(lineaValor(6)) >= dbl12mahorro Then _
        Printer.Print Linea(6); Tab(40); "    S/." & pfstr_Alineado(Format$(lineaValor(6), "###,##0.00"), 14, Derecha)
    Printer.Print
    
    Printer.FontName = "Arial"
    Printer.FontSize = 7
    Printer.FontBold = True
    centra_printer Linea(7)
    Printer.Print
    Printer.FontBold = False
    Printer.Print Linea(8)
    justifica_printer 20, 4000, Printer.CurrentY, Linea(9)
    Printer.Print
    Printer.Print
    Printer.EndDoc
   
    Screen.MousePointer = vbDefault
    MsgBox "Recoger el voucher de la cuponera", vbInformation, App.ProductName
    Exit Sub
Control:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
End Sub

