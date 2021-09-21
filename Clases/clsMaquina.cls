VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaquina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim rsMaquina As oraDynaset
Function Valida() As Boolean
    Dim strIP As String
    strIP = Replace(Trim(Obtener(12)), Chr(0), "")
    Set rsMaquina = gclsOracle.FN_Cursor("BTLPROD.PKG_MAQUINA.FN_LISTA", 0, strIP)
    If rsMaquina.EOF Then MsgBox "El ip de esta maquina no esta registrado", vbCritical, App.ProductName: Valida = False: Exit Function
    Valida = True
End Function
Public Property Get UbigeoLocal() As String
    UbigeoLocal = "" & rsMaquina("UBIGEO").Value
End Property

Public Property Get Delivery() As String
    Delivery = "" & rsMaquina("FLG_DELIVERY").Value
End Property

Public Property Get CentroCosto() As String
    CentroCosto = "" & rsMaquina("DES_CNT_COSTO").Value
End Property

Public Property Get NumIp() As String
    NumIp = "" & rsMaquina("NUM_IP").Value
End Property
Public Property Get CodLocalCallCenter() As String
    CodLocalCallCenter = "" & rsMaquina("CEN_COS").Value
End Property

Public Property Get TipoMaquina() As String
    TipoMaquina = "" & rsMaquina("COD_TIPO_MAQUINA").Value
End Property

Public Property Get NombrePC() As String
    NombrePC = "" & rsMaquina("COD_MAQUINA").Value
End Property
Public Property Get Cia() As String
    Cia = "" & rsMaquina("CIA").Value
End Property
Public Property Get CodigoLocal() As String
    CodigoLocal = "" & rsMaquina("COD_LOCAL").Value
End Property
Public Property Get NombreLocal() As String
    NombreLocal = "" & rsMaquina("NOMBRE_LOCAL").Value
End Property
Public Property Get DireccionLocal() As String
    DireccionLocal = "" & rsMaquina("DES_DIRECCION").Value
End Property
'---- Para el registro de la maquina 07/12/2006 ----'

Public Property Get flgDeliveryProv() As String
    flgDeliveryProv = "" & rsMaquina("FLG_DELIVERY_PROV").Value
End Property


Public Function ListaTipoMaquina() As oraDynaset
    On Error GoTo CtrlErr
    Set ListaTipoMaquina = gclsOracle.FN_Cursor("BTLPROD.PKG_MAQUINA.FN_TIPO_MAQUINA", 0)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMaquina_LstTipoMaquina", Err.Description
End Function


Public Function Lista(vstrCia As String, _
                        Optional vstrCodLocal As String = "") As oraDynaset
    On Error GoTo CtrlErr
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_MAQUINA.FN_LISTA", 0, vstrCia, vstrCodLocal)

    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMaquina_fnLista", Err.Description
End Function

Public Function DocumentoxLocal(vstrCia As String, _
                                  vstrCodLocal As String, _
                                  vstrCodMaquina As String) As oraDynaset
    On Error GoTo CtrlErr
    Set DocumentoxLocal = gclsOracle.FN_Cursor("BTLPROD.PKG_MAQUINA.FN_DOCUMENTO_X_MAQUINA", 0, vstrCia, vstrCodLocal, vstrCodMaquina)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMaquina_fnLista", Err.Description
End Function


Public Function GrabaCorXDoc(ByVal vstrCia As String, _
                             ByVal vstrCodLocal As String, _
                             ByVal vstrCodMaquina As String, _
                             ByVal vstrCadCodTipoDoc As String, _
                             ByVal vstrCadNumActual As String, _
                             ByVal vstrCodUsuario As String, _
                             ByVal CodigoMaquinaRel As String, _
                             ByVal FlagColaImpresion As String, _
                             ByVal vstrCadTicketera As String, _
                             ByVal vstrCadCodFormato As String _
                             ) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant
    gvarValores = Array(vstrCia, _
                        vstrCodLocal, _
                        vstrCodMaquina, _
                        vstrCadCodTipoDoc, _
                        vstrCadNumActual, _
                        vstrCodUsuario, CodigoMaquinaRel, FlagColaImpresion, vstrCadTicketera, vstrCadCodFormato)
    
    gvarIO = Array(entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada)
                   
    GrabaCorXDoc = gclsOracle.SP("BTLPROD.PKG_MAQUINA.SP_GRABA_COR_X_DOC", _
                              gvarValores, _
                              gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsMaquina_GrabaCorXDoc", Err.Description
End Function

Public Function fn_Verifica_Sec_Correlativo(ByVal vstrCia As String, _
                                            ByVal vstrCodLocal As String, _
                                            ByVal vstrCadCodTipoDoc As String, _
                                            ByVal vstrCadNumActual As String)
On Error GoTo CtrlErr
    
    fn_Verifica_Sec_Correlativo = gclsOracle.FN_Valor("BTLPROD.PKG_MAQUINA.FN_VERIFICA_SEC_CORRELATIVO", _
                                                        vstrCia, _
                                                        vstrCodLocal, _
                                                        vstrCadCodTipoDoc, _
                                                        vstrCadNumActual)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsMaquina_fn_Verifica_Sec_Correlativo", Err.Description
End Function

'-- Optienes el IP de la maquina --'
'''Public Function fnIP_x_Maquina(ByVal vstrCia As String, _
'''                               ByVal vstrCodLocal As String, _
'''                               ByVal vstrCodMaquina As String, _
'''                               ByVal vstrCodTipoMaquina As String)
'''    On Error GoTo CtrlErr
'''    fnIP_x_Maquina = gclsOracle.FN_Valor("BTLPROD.PKG_MAQUINA.FN_IP_MAQUINA", vstrCia, _
'''                                                                              vstrCodLocal, _
'''                                                                              vstrCodMaquina, _
'''                                                                              vstrCodTipoMaquina)
'''
'''    Exit Function
'''CtrlErr:
'''    Err.Raise Err.Description, "clsMaquina_fnIP_x_Maquina", Err.Description
'''End Function

Public Function GrabaIP(ByVal vstrCia As String, _
                        ByVal vstrCodLocal As String, _
                        ByVal vstrCodMaquina As String, _
                        ByVal vstrCodTipoMaquina As String, _
                        ByVal vstrNumIP As String) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant
    gvarValores = Array(vstrCia, _
                        vstrCodLocal, _
                        vstrCodMaquina, _
                        vstrCodTipoMaquina, _
                        vstrNumIP)
    
    gvarIO = Array(entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada)
                   
    GrabaIP = gclsOracle.SP("BTLPROD.PKG_MAQUINA.SP_ACT_IP", _
                              gvarValores, _
                              gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsMaquina_GrabaIP", Err.Description
End Function

Public Function MaquinaLocal(ByVal vstrCia As String, _
                             ByVal vstrCodLocal As String, _
                             Optional ByVal vstrflgNinguno As String = "", _
                             Optional ByVal vstrCodMaqProp As String = "") As oraDynaset
On Error GoTo CtrlErr
    Set MaquinaLocal = gclsOracle.FN_Cursor("BTLPROD.PKG_MAQUINA.FN_LISTA_LOCAL", 0, vstrCia, vstrCodLocal, vstrflgNinguno, vstrCodMaqProp)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMaquina_MaquinaLocal", Err.Description
                             
End Function

Public Function ListaDocDefecto(ByVal Cia As String, ByVal NumIp As String) As String
On Error GoTo CtrlErr

    ListaDocDefecto = gclsOracle.FN_Valor("BTLPROD.PKG_MAQUINA.FN_LISTA_DOC_DEFECTO", Cia, NumIp)

Exit Function
CtrlErr:

    Err.Raise Err.Number, "clsMaquina.ListaDocDefecto", Err.Description
End Function




Public Function Graba(ByVal vstrCia As String, _
                             ByVal vstrCodLocal As String, _
                             ByVal vstrCodMaquina As String, _
                             ByVal vstrCodTipoMaquina As String, _
                             ByVal vstrFlgDelivery As String, _
                             ByVal vstrNumIP As String, _
                             ByVal vstrFlgActivo As String, _
                             ByVal vstrCodUsuario As String _
                             ) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant
    gvarValores = Array(vstrCia, _
                        vstrCodLocal, _
                        vstrCodMaquina, _
                        vstrCodTipoMaquina, _
                        vstrFlgDelivery, _
                        vstrNumIP, _
                        vstrFlgActivo, _
                        vstrCodUsuario)
    
    gvarIO = Array(entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada)
                   
    Graba = gclsOracle.SP("BTLPROD.PKG_MAQUINA.SP_GRABA", _
                              gvarValores, _
                              gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsMaquina.Graba", Err.Description
End Function

Public Function ListaTicketeraLocal(ByVal vstrCia As String, _
                                     ByVal vstrCodLocal As String) As oraDynaset
On Error GoTo CtrlErr
    Set ListaTicketeraLocal = gclsOracle.FN_Cursor("BTLPROD.PKG_MAQUINA.LISTA_TICKETERAS_LOCAL", 0, vstrCia, vstrCodLocal)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMaquina", Err.Description
End Function

Public Function Serie_Ticketera(ByVal vstrCia As String, _
                                 ByVal vstrCodLocal As String, _
                                 ByVal vstrTipoDoc As String, _
                                 ByVal vstrCodMaquina As String) As String
On Error GoTo CtrlErr
    Serie_Ticketera = gclsOracle.FN_Valor("BTLPROD.PKG_MAQUINA.FN_SERIE_TICKETERA", vstrCia, vstrCodLocal, vstrTipoDoc, vstrCodMaquina)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMaquina", Err.Description
End Function

Public Function fn_Existe_Rel_Tkt_Maq(ByVal Cia As String, _
                                      ByVal CodLocal As String, _
                                      ByVal TipoDoc As String, _
                                      ByVal Ticketera As String) As Integer
On Error GoTo CtrlErr
    fn_Existe_Rel_Tkt_Maq = gclsOracle.FN_Valor("BTLPROD.PKG_MAQUINA.FN_REL_EXISTE_TKT_MAQUINA", Cia, CodLocal, TipoDoc, Ticketera)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMaquina.fn_Existe_Rel_Tkt_Maq", Err.Description
End Function


Public Function CodFormato(ByVal vstrCia As String, _
                                      ByVal vstrTipoDoc As String, _
                                      ByVal vstrCodMaq As String) As String

    CodFormato = "" & gclsOracle.FN_Valor("BTLPROD.PKG_MAQUINA.FN_FORMATO", vstrCia, vstrCodMaq, vstrTipoDoc)
    
End Function



Public Function ListaSerieDocumentos(ByVal NombreMaquina As String) As oraDynaset
On Error GoTo CtrlErr
    Set ListaSerieDocumentos = gclsOracle.FN_Cursor("BTLPROD.PKG_CONTINGENCIA.FN_LISTA_SERIE_X_MAQUINA", 0, NombreMaquina)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsMaquina.ListaSerieDocumentos", Err.Description
End Function

