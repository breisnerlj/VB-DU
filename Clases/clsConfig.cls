VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Graba(ByVal vintCodigo As Integer, _
                ByVal vstrParametro As String, _
                ByVal vstrValor As String, _
                ByVal vstrCommit As String)

    On Error GoTo Control
    Dim lstrSql As String



    gclsOracle.ODataBase.Parameters.Add "A_COD_CONFIG", vintCodigo, ORAPARM_INPUT, ORATYPE_NUMBER
    gclsOracle.ODataBase.Parameters.Add "A_COD_PARAMETRO", vstrParametro, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_VAL_PARAMETRO", vstrValor, ORAPARM_INPUT, ORATYPE_VARCHAR2
    gclsOracle.ODataBase.Parameters.Add "A_FLG_COMMIT", vstrCommit, ORAPARM_INPUT, ORATYPE_CHAR

    lstrSql = "BEGIN " & _
              "CMR.PKG_CONFIG.SP_GRABA_VALOR(A_COD_CONFIG     =>  :A_COD_CONFIG, " & _
              "                              A_COD_PARAMETRO  =>  :A_COD_PARAMETRO, " & _
              "                              A_VAL_PARAMETRO  =>  :A_VAL_PARAMETRO, " & _
              "                              A_FLG_COMMIT     =>  :A_FLG_COMMIT); " & _
              "END;"

    gclsOracle.ODataBase.ExecuteSQL (lstrSql)

   Exit Sub
Control:
    Err.Raise Err.Number, "clsOrdenCompra.GrabaConfig", Err.Description
End Sub

Function Valor(ByVal vintCodigo As Integer, _
               ByVal vstrParametro) As String
    On Error GoTo CtrlErr
        Valor = gclsOracle.FN_Valor("CMR.PKG_CONFIG.FN_VALOR", vintCodigo, vstrParametro)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsOrdenCompra.Lista_LeadTime", Err.Description
End Function


