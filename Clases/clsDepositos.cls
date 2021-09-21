VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDepositos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private xDepositos As New XArrayDB
Private strControl As String

Public Property Get Control() As String
    Control = strControl
End Property

Public Property Let Control(ByVal vNewValue As String)
    strControl = vNewValue
End Property

Public Function ListaCtaCte(ByVal vstrBanco As String, _
                            Optional ByVal vstrCia As String = "") As oraDynaset

    On Error GoTo CtrlErr
        Set ListaCtaCte = gclsOracle.FN_Cursor("BTLPROD.PKG_DEPOSITOS.FN_LISTA_CTACTE", 0, vstrBanco, vstrCia)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDespositos.ListaCtaCte", Err.Description
End Function

Public Function ListaBancos() As oraDynaset

    On Error GoTo CtrlErr
        Set ListaBancos = gclsOracle.FN_Cursor("BTLPROD.PKG_DEPOSITOS.FN_LISTA_BANCOS", 0)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDespositos.ListaBancos", Err.Description
End Function

Public Function ListaRemitos(ByVal vstrCia As String, ByVal vstrLocal As String, _
                             ByVal vstrMoneda As String, ByVal vstrFecha As String) As oraDynaset

    On Error GoTo CtrlErr
        Set ListaRemitos = gclsOracle.FN_Cursor("BTLPROD.PKG_DEPOSITOS.FN_LISTA_REMITOS", 0, _
                                                vstrCia, vstrLocal, vstrMoneda, vstrFecha)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDespositos.ListaRemitos", Err.Description
End Function

Public Function LocalPortaValor(ByVal vstrLocal As String) As Boolean
    
    On Error GoTo CtrlErr
        LocalPortaValor = IIf(gclsOracle.FN_Valor("BTLPROD.PKG_DEPOSITOS.FN_PORTAVALOR_BTL", vstrLocal) = 1, True, False)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDespositos.LocalPortaValor", Err.Description
End Function

Public Function Graba(ByVal vstrCia As String, _
                      ByVal vstrLocal As String, _
                      ByVal vstrCueban As String, _
                      ByVal vstrNroOperacion As String, _
                      ByVal vstrImporte As Double, _
                      ByVal vstrDiferencia As Double, _
                      ByVal vstrFchDeposito As String, _
                      ByVal vstrObservaciones As String, _
                      ByVal vstrUsuario As String, _
                      ByVal vstrMoneda As String, _
                      ByVal varrRemito As Variant, _
                      ByVal varrImpRemito As Variant, _
                      ByVal vstrTipo As String, _
                      Optional ByVal varrRelCodDiferencia As Variant, _
                      Optional ByVal varrRelMonto As Variant) As String

On Error GoTo CtrlErr
    
    Dim lstrSql As String
    Dim i, j As Integer
    Dim lintFilas, lintFilas2 As Integer
    
    Dim lparRemito As OracleInProcServer.OraParamArray
    Dim lparImpRemito As OracleInProcServer.OraParamArray

    Dim lparRelCodDiferencia As OracleInProcServer.OraParamArray
    Dim lparRelMonto As OracleInProcServer.OraParamArray
    
    lintFilas = UBound(varrRemito) - LBound(varrRemito) + 1
    lintFilas2 = UBound(varrRelCodDiferencia) - LBound(varrRelCodDiferencia) + 1
    gclsOracle.LimpiaParametros
    
    With gclsOracle.ODataBase.Parameters
         .AddTable "A_CAD_COD_REMITO", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas, 100
         .AddTable "A_CAD_IMP_REMITO", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas, 100
'
         .AddTable "A_CAD_COD_DIF", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas2, 100
         .AddTable "A_CAD_MONTO_DIF", ORAPARM_INPUT, ORATYPE_VARCHAR2, lintFilas2, 100
'

         .Add "A_CIA", vstrCia, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_COD_LOCAL", vstrLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_CUEBAN", vstrCueban, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_NRO_OPERACION", vstrNroOperacion, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_IMP_DEPOSITO", vstrImporte, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_IMP_DIFERENCIA", vstrDiferencia, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_FCH_DESPOSITO", vstrFchDeposito, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_DES_OBS", vstrObservaciones, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_COD_USUARIO", vstrUsuario, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_MONEDA", vstrMoneda, ORAPARM_INPUT, ORATYPE_VARCHAR2
         .Add "A_TIP_MAQUINA", vstrTipo, ORAPARM_INPUT, ORATYPE_VARCHAR2
    End With
    
    Set lparRemito = gclsOracle.ODataBase.Parameters("A_CAD_COD_REMITO")
    Set lparImpRemito = gclsOracle.ODataBase.Parameters("A_CAD_IMP_REMITO")

    Set lparRelCodDiferencia = gclsOracle.ODataBase.Parameters("A_CAD_COD_DIF")
    Set lparRelMonto = gclsOracle.ODataBase.Parameters("A_CAD_MONTO_DIF")

    For i = LBound(varrRemito) To UBound(varrRemito)
        lparRemito(i) = varrRemito(i)
        lparImpRemito(i) = varrImpRemito(i)
    Next i
    
    For j = LBound(varrRelCodDiferencia) To UBound(varrRelCodDiferencia)
        lparRelCodDiferencia(j) = varrRelCodDiferencia(j)
        lparRelMonto(j) = varrRelMonto(j)
    Next j
    
    
    lstrSql = "BEGIN " & _
              "BTLPROD.PKG_DEPOSITOS.SP_GRABA(A_CIA             => :A_CIA," & _
              "                               A_COD_LOCAL       => :A_COD_LOCAL," & _
              "                               A_CUEBAN          => :A_CUEBAN," & _
              "                               A_NRO_OPERACION   => :A_NRO_OPERACION," & _
              "                               A_IMP_DEPOSITO    => :A_IMP_DEPOSITO," & _
              "                               A_IMP_DIFERENCIA  => :A_IMP_DIFERENCIA," & _
              "                               A_FCH_DESPOSITO   => :A_FCH_DESPOSITO," & _
              "                               A_DES_OBS         => :A_DES_OBS," & _
              "                               A_COD_USUARIO     => :A_COD_USUARIO," & _
              "                               A_MONEDA          => :A_MONEDA," & _
              "                               A_CAD_COD_REMITO  => :A_CAD_COD_REMITO," & _
              "                               A_CAD_IMP_REMITO  => :A_CAD_IMP_REMITO," & _
              "                               A_TIP_MAQUINA     => :A_TIP_MAQUINA," & _
              "                               A_CAD_COD_DIF     => :A_CAD_COD_DIF," & _
              "                               A_CAD_MONTO_DIF     => :A_CAD_MONTO_DIF); " & _
              "END;"
    
    gclsOracle.Execute lstrSql

    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDespositos.Graba", Err.Description
End Function

Public Function ListaDepositos(ByVal vstrCia As String, ByVal vstrLocal As String, _
                               ByVal vstrFchIni As String, ByVal vstrFchFin As String) As oraDynaset
    
    On Error GoTo CtrlErr
        Set ListaDepositos = gclsOracle.FN_Cursor("BTLPROD.PKG_DEPOSITOS.FN_LISTA_DEPOSITOS", 0, vstrCia, vstrLocal, vstrFchIni, vstrFchFin)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsDespositos.ListaDepositos", Err.Description
End Function

Public Function Anula(ByVal vstrCia As String, _
                      ByVal vstrLocal As String, _
                      ByVal vstrNroDeposito As String, _
                      ByVal vstrFchDesposito As String) As String

On Error GoTo CtrlErr
    Dim gvarValores As Variant
    Dim gvarIO  As Variant
    
    gvarValores = Array(vstrCia, vstrLocal, vstrNroDeposito, vstrFchDesposito)
    
    gvarIO = Array(entrada, entrada, entrada, entrada)
                   
    Anula = gclsOracle.SP("BTLPROD.PKG_DEPOSITOS.SP_ANULA", gvarValores, gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsDespositos.Graba", Err.Description
End Function
