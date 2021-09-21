VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRemito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private byteNumElementosRemito As Integer

Public Function BancoxBtl(ByVal vstrCodBtl As String) As String
    On Error GoTo CtrlErr
        BancoxBtl = gclsOracle.FN_Valor("BTLPROD.PKG_REMITO.FN_BANCO_X_BTL", vstrCodBtl)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemito", Err.Description
End Function

Public Function CtaCtexBtl(ByVal vstrCodBtl As String, Optional vstrMoneda As String) As String
    On Error GoTo CtrlErr
        CtaCtexBtl = gclsOracle.FN_Valor("BTLPROD.PKG_REMITO.FN_CTATE_X_BTL", vstrCodBtl, vstrMoneda)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemito", Err.Description
End Function

Public Function Portavalor(ByVal vstrCodBtl As String) As String
    On Error GoTo CtrlErr
        Portavalor = gclsOracle.FN_Valor("BTLPROD.PKG_REMITO.FN_PORTAVALOR_X_BTL", vstrCodBtl)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemito", Err.Description
End Function

Public Function TipoPortavalor(ByVal vstrCodBtl As String) As String
    On Error GoTo CtrlErr
        TipoPortavalor = gclsOracle.FN_Valor("BTLPROD.PKG_REMITO.FN_TIPO_PORTAVALOR", vstrCodBtl)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemito", Err.Description
End Function

Public Function ListaRemesas(ByVal vstrCia As String, _
                             ByVal vstrCodBtl As String, _
                             ByVal vstrFglTipo As String, _
                             ByVal vstrTipoEnv As String, _
                             ByVal vstrFchIni As String, _
                             ByVal vstrFchFin As String) As oraDynaset
    On Error GoTo CtrlErr
       Set ListaRemesas = gclsOracle.FN_Cursor("BTLPROD.PKG_REMITO.FN_LISTA", 0, _
                                                  vstrCia, _
                                                  vstrCodBtl, _
                                                  vstrFglTipo, _
                                                  vstrTipoEnv, _
                                                  vstrFchIni, _
                                                  vstrFchFin)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemito", Err.Description
End Function

Public Function TotalSobres(ByVal vstrCia As String, _
                            ByVal vstrCodBtl As String, _
                            ByVal vstrFchIni As String, _
                            ByVal vstrFchFin As String) As String
    On Error GoTo CtrlErr
        TotalSobres = gclsOracle.FN_Valor("BTLPROD.PKG_REMITO.FN_TOTAL_SOBRES", _
                                                  vstrCia, _
                                                  vstrCodBtl, _
                                                  vstrFchIni, _
                                                  vstrFchFin)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemito", Err.Description
End Function

Public Function TotalSobresRemesa(ByVal vstrCia As String, _
                                  ByVal vstrCodBtl As String, _
                                  ByVal vstrMaquina As String, _
                                  ByVal vstrFchIni As String, _
                                  ByVal vstrFchFin As String, _
                                  Optional vstrMon As String, _
                                  Optional vstrRemesa As String) As String
    On Error GoTo CtrlErr
    TotalSobresRemesa = gclsOracle.FN_Valor("BTLPROD.PKG_REMITO.FN_TOTAL_SOBRES_X_REMESA", _
                                                                    vstrCia, _
                                                                    vstrCodBtl, _
                                                                    vstrMaquina, _
                                                                    vstrFchIni, _
                                                                    vstrFchFin, _
                                                                    vstrMon, _
                                                                    vstrRemesa)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemito", Err.Description

End Function

Public Function CtaBtl(ByVal vstrCodBtl As String, ByVal vstrMoneda As String) As String
    On Error GoTo CtrlErr
    CtaBtl = gclsOracle.FN_Valor("BTLPROD.PKG_REMITO.FN_CTA_BTL_MONEDA", vstrCodBtl, vstrMoneda)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemito", Err.Description
End Function

Public Function Graba(ByVal oradb As OraDatabase, _
                      ByVal vdblImpSoles As Double, _
                      ByVal vdblImpDolares As Double, _
                      ByVal vstrCodBanco As String, _
                      ByVal vstrNumPrecinto As String, _
                      ByVal vstrFchRemito As String, _
                      ByVal vdblFDS As Double, _
                      ByRef pCodRemitoGen As String, _
                      ByVal vstrNumRemitoPreimp As String) As String
                 
Dim i As Integer

On Error GoTo Control

byteNumElementosRemito = IIf(objLiquidacion.Remito.UpperBound(1) < 0, 1, objLiquidacion.Remito.UpperBound(1)) + 1

    For i = oradb.Parameters.Count - 1 To 0 Step -1
       oradb.Parameters.Remove i
    Next

oradb.Parameters.Add "A_COD_LOCAL", objUsuario.CodigoLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_IMP_SOLES", vdblImpSoles, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_IMP_DOLARES", vdblImpDolares, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_USUARIO", objUsuario.codigo, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_BANCO", vstrCodBanco, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_NUM_PRECINTO", vstrNumPrecinto, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_FCH_REMITO", vstrFchRemito, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_MTO_FDS", vdblFDS, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_NUM_REMITOPREIMP", vstrNumRemitoPreimp, ORAPARM_INPUT, ORATYPE_VARCHAR2

oradb.Parameters.Add "V_COD_REMITOGEN", "", ORAPARM_OUTPUT, ORATYPE_VARCHAR2

oradb.Parameters.AddTable "A_CAD_SEC_DET_REM", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosRemito, 200
oradb.Parameters.AddTable "A_CAD_IMP_SOLES", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosRemito, 200
oradb.Parameters.AddTable "A_CAD_IMP_DOLARES", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosRemito, 200
oradb.Parameters.AddTable "A_CAD_COD_DEPENDIENTE", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosRemito, 200



Dim SecDetRem As OracleInProcServer.OraParamArray
Dim ImpSoles As OracleInProcServer.OraParamArray
Dim ImpDolares As OracleInProcServer.OraParamArray
Dim CodDepen As OracleInProcServer.OraParamArray

Set SecDetRem = oradb.Parameters("A_CAD_SEC_DET_REM")
Set ImpSoles = oradb.Parameters("A_CAD_IMP_SOLES")
Set ImpDolares = oradb.Parameters("A_CAD_IMP_DOLARES")
Set CodDepen = oradb.Parameters("A_CAD_COD_DEPENDIENTE")

'------------------------REMITOS----------------------------
    
    
    For i = 0 To objLiquidacion.Remito.UpperBound(1)
      'SOLES'
      If objLiquidacion.Remito(i, 9) = "-1" Then
        If objLiquidacion.Remito(i, 8) = "1" Then
              SecDetRem(i) = objLiquidacion.Remito(i, 0)
              ImpSoles(i) = objLiquidacion.Remito(i, 5)
              ImpDolares(i) = "0"
              CodDepen(i) = objLiquidacion.Remito(i, 10)
        ElseIf objLiquidacion.Remito(i, 8) = "2" Then
        'DOLARES'
              SecDetRem(i) = objLiquidacion.Remito(i, 0)
              ImpSoles(i) = "0"
              ImpDolares(i) = objLiquidacion.Remito(i, 5)
              CodDepen(i) = objLiquidacion.Remito(i, 10)
        End If
      End If
    Next

    oradb.ExecuteSQL " BEGIN BTLPROD.PKG_REMITO.SP_GRABA(:A_COD_LOCAL, :A_IMP_SOLES, :A_IMP_DOLARES, :A_COD_USUARIO, :A_COD_BANCO, :A_NUM_PRECINTO, :A_CAD_SEC_DET_REM, :A_CAD_IMP_SOLES, :A_CAD_IMP_DOLARES, :A_CAD_COD_DEPENDIENTE, :A_FCH_REMITO, :A_MTO_FDS, :V_COD_REMITOGEN,:A_NUM_REMITOPREIMP) ; END;"
    
    pCodRemitoGen = Trim(oradb.Parameters("V_COD_REMITOGEN").Value)
Exit Function

Control:

    Err.Raise Err.Number, "clsRemito.Grabar", Err.Description

End Function

'Public Property Get RetCodRemito() As String
'    RetCodRemito = pCodRemitoGen
'End Property

Public Function ListaGenerados(ByVal vstrCodLocal As String, _
                               ByVal vstrFchIni As String, _
                               ByVal vstrFchFin As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaGenerados = gclsOracle.FN_Cursor("BTLPROD.PKG_REMITO.FN_LISTA_GEN", 0, vstrCodLocal, vstrFchIni, vstrFchFin)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemito_ListaGenreados", Err.Description
End Function

Public Function Anula(ByVal vstrCodLocal As String, _
                      ByVal vstrCodRemito As String, Optional ByVal FlgCreditoCobranza As String = "0") As String
                      
On Error GoTo CtrlErr
    Dim gvarValores As Variant
    Dim gvarIO  As Variant
    
    gvarValores = Array(vstrCodLocal, _
                        vstrCodRemito, FlgCreditoCobranza)
    
    gvarIO = Array(entrada, entrada, _
                   entrada)
                   
    Anula = gclsOracle.SP("BTLPROD.PKG_REMITO.SP_ANULA", _
                              gvarValores, _
                              gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsRemito_Anula", Err.Description
End Function

' ---- @@@@@@@@@@@@@@@@@@@@@@@ ------'
' ---- RE IMPRESION DE REMITOS ------'
' ---- @@@@@@@@@@@@@@@@@@@@@@@ ------'

Public Function Dev_CtdReg_Remeda(ByVal vstrCodLocal As String, _
                                  ByVal vstrCodRemito As String) As String

On Error GoTo CtrlErr
    Dev_CtdReg_Remeda = gclsOracle.FN_Valor("BTLPROD.PKG_REMITO.FN_DEV_CTD_REG_REMESAS", vstrCodLocal, vstrCodRemito)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemito_Dev_CtdReg_Remeda", Err.Description

End Function

Public Function Re_Imp_Remito(ByVal vstrCodLocal As String, _
                              ByVal vstrCodRemito As String) As oraDynaset

On Error GoTo CtrlErr
    Set Re_Imp_Remito = gclsOracle.FN_Cursor("BTLPROD.PKG_REMITO.FN_REIMP_REMITO", 0, vstrCodLocal, vstrCodRemito)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemito_Re_Imp_Remito", Err.Description

End Function

Public Function Re_Imp_Remito_Hermes(ByVal vstrCodLocal As String, _
                                     ByVal vstrCodRemito As String, _
                                     ByVal vstrMoneda As String) As oraDynaset

On Error GoTo CtrlErr
    Set Re_Imp_Remito_Hermes = gclsOracle.FN_Cursor("BTLPROD.PKG_REMITO.FN_REIMP_REMITO_HERMES", 0, vstrCodLocal, vstrCodRemito, vstrMoneda)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemito_Re_Imp_Remito_Hermes", Err.Description

End Function

Public Function DireccLocal(ByVal vstrCodLocal As String) As String
    On Error GoTo CtrlErr
    DireccLocal = gclsOracle.FN_Valor("BTLPROD.PKG_REMITO.FN_DEV_DIRECC_LOCAL", vstrCodLocal)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemito_DireccLocal", Err.Description

End Function

Public Function ExisteRemito(ByVal vstrCodRemito As String) As String
    On Error GoTo CtrlErr
    ExisteRemito = gclsOracle.FN_Valor("BTLPROD.PKG_REMITO.FN_EXISTE_REMITO", vstrCodRemito)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemito.ExisteRemito", Err.Description

End Function
