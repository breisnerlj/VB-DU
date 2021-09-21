VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRemesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Dim obLiquidacion As New clsLiquidacion
Private byteNumElementosRemesa As Byte

Public Function DevSecSol(ByVal vstrMaquina As String, ByVal vstrCia As String) As String
    On Error GoTo CtrlErr
        DevSecSol = gclsOracle.FN_Valor("BTLPROD.PKG_REMESA.FN_DEV_SEC_SOLES", vstrMaquina, vstrCia)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemesa", Err.Description
End Function

Public Function DevSecDol(ByVal vstrMaquina As String, ByVal vstrCia As String) As String
    On Error GoTo CtrlErr
        DevSecDol = gclsOracle.FN_Valor("BTLPROD.PKG_REMESA.FN_DEV_SEC_DOLARES", vstrMaquina, vstrCia)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemesa", Err.Description
End Function

Public Function DevSecDep(ByVal vstrMaquina As String, ByVal vstrCia As String) As String
    On Error GoTo CtrlErr
        DevSecDep = gclsOracle.FN_Valor("BTLPROD.PKG_REMESA.FN_DEV_SEC_DEPOSITO", vstrMaquina, vstrCia)
        
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemesa", Err.Description
End Function

Public Static Property Get NumElementosRemesa() As Byte
NumElementosRemesa = byteNumElementosRemesa

End Property

Public Static Property Let NumElementosRemesa(ByVal lbyteNumElementosRemesa As Byte)
byteNumElementosRemesa = lbyteNumElementosRemesa
End Property

Public Function Graba(ByVal oradb As OraDatabase, _
                      ByVal vstrUsuOrig As String, _
                      ByVal vstrUsuDest As String, _
                      ByVal vstrCodMaquina As String, _
                      ByVal vstrCodLiquidacion As String) As String
Dim i As Integer

On Error GoTo Control

byteNumElementosRemesa = IIf(objLiquidacion.Remesa.UpperBound(1) < 0, 1, objLiquidacion.Remesa.UpperBound(1)) + 1

    For i = oradb.Parameters.Count - 1 To 0 Step -1
       oradb.Parameters.Remove i
    Next

oradb.Parameters.Add "A_COD_MAQUINA", vstrCodMaquina, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_USU_ORG", vstrUsuOrig, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_USU_DEST", vstrUsuDest, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_USUARIO", objUsuario.Codigo, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_CIA", objUsuario.CodigoEmpresa, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_BTL", objUsuario.CodigoLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2
oradb.Parameters.Add "A_COD_LIQUIDACION", vstrCodLiquidacion, ORAPARM_INPUT, ORATYPE_VARCHAR2

oradb.Parameters.AddTable "A_CAD_MONEDA", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosRemesa, 200
oradb.Parameters.AddTable "A_CAD_IMP_SOBRE", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosRemesa, 200
oradb.Parameters.AddTable "A_CAD_CONCEPTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosRemesa, 200
oradb.Parameters.AddTable "A_CAD_FPAGO_P", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosRemesa, 200
oradb.Parameters.AddTable "A_CAD_FPAGO_H", ORAPARM_INPUT, ORATYPE_VARCHAR2, byteNumElementosRemesa, 200

Dim CodMomeda As OracleInProcServer.OraParamArray
Dim ImpSobre As OracleInProcServer.OraParamArray
Dim Concepto As OracleInProcServer.OraParamArray
Dim CodFP_P As OracleInProcServer.OraParamArray
Dim CodFP_H As OracleInProcServer.OraParamArray

Set CodMomeda = oradb.Parameters("A_CAD_MONEDA")
Set ImpSobre = oradb.Parameters("A_CAD_IMP_SOBRE")
Set Concepto = oradb.Parameters("A_CAD_CONCEPTO")
Set CodFP_P = oradb.Parameters("A_CAD_FPAGO_P")
Set CodFP_H = oradb.Parameters("A_CAD_FPAGO_H")

'------------------------REMESAS----------------------------
    
    For i = 0 To objLiquidacion.Remesa.UpperBound(1)
      If objLiquidacion.Remesa(i, 7) = "1" Then
            CodMomeda(i) = objLiquidacion.Remesa(i, 7)
            ImpSobre(i) = objLiquidacion.Remesa(i, 5)
            Concepto(i) = objLiquidacion.Remesa(i, 0)
            CodFP_P(i) = objLiquidacion.Remesa(i, 8)
            CodFP_H(i) = objLiquidacion.Remesa(i, 9)
      ElseIf objLiquidacion.Remesa(i, 7) = "2" Then
            CodMomeda(i) = objLiquidacion.Remesa(i, 7)
            ImpSobre(i) = objLiquidacion.Remesa(i, 6)
            Concepto(i) = objLiquidacion.Remesa(i, 0)
            CodFP_P(i) = objLiquidacion.Remesa(i, 8)
            CodFP_H(i) = objLiquidacion.Remesa(i, 9)
      End If
    Next

    oradb.ExecuteSQL " BEGIN BTLPROD.PKG_REMESA.SP_GRABA(:A_COD_MAQUINA, :A_COD_USU_ORG, :A_COD_USU_DEST, :A_COD_USUARIO, :A_CIA, :A_COD_BTL, :A_COD_LIQUIDACION, :A_CAD_MONEDA, :A_CAD_IMP_SOBRE, :A_CAD_CONCEPTO, :A_CAD_FPAGO_P, :A_CAD_FPAGO_H) ; END;"

Exit Function

Control:

    Err.Raise Err.Number, "clsRemesa.Grabar", Err.Description

End Function

Public Function Lista(ByVal vstrCia As String, _
                      ByVal vstrCodLocal As String, _
                      ByVal vstrFchIni As String, _
                      ByVal vstrFchFin As String) As oraDynaset
    On Error GoTo CtrlErr
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_REMESA.FN_LISTA", 0, vstrCia, vstrCodLocal, vstrFchIni, vstrFchFin)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemesa_Lista", Err.Description
End Function
                      
Public Function Anula(ByVal vstrCia As String, _
                      ByVal vstrCodLocal As String, _
                      ByVal vstrCodMaquina As String, _
                      ByVal vstrCodLiquidacion As String, _
                      ByVal vstrCodRemesa As String) As String
On Error GoTo CtrlErr
    Dim gvarValores As Variant
    Dim gvarIO  As Variant
    
    gvarValores = Array(vstrCia, _
                        vstrCodLocal, _
                        vstrCodMaquina, _
                        vstrCodLiquidacion, _
                        vstrCodRemesa)
    
    gvarIO = Array(entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada)
                   
    Anula = gclsOracle.SP("BTLPROD.PKG_REMESA.SP_ANULA", _
                              gvarValores, _
                              gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsRemesa_Anula", Err.Description
End Function

Public Function Impresion(ByVal vstrCia As String, _
                          ByVal vstrCodLocal As String, _
                          ByVal vstrCodLiquidacion As String)
Dim objPrinter As New clsImpresiones
Dim Imprimir As oraDynaset
    On Error GoTo CtrlErr
    Set Imprimir = gclsOracle.FN_Cursor("BTLPROD.PKG_REMESA.FN_IMP_REMESA", 0, vstrCia, vstrCodLocal, vstrCodLiquidacion)
    
    If Imprimir("IMP_SOLES").Value > 0 Then
        If MsgBox("Se va a imprimir los Montos en Soles - Colocar el formato", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            objPrinter.Imprime_Remesas_Soles objUsuario.CodigoEmpresa, _
                                             Imprimir("COD_LOCAL").Value, _
                                             Imprimir("COD_LIQUIDACION").Value, _
                                             Imprimir("COD_REMESA").Value, _
                                             Format(Imprimir("IMP_SOLES").Value, "###,###0.00"), _
                                             Imprimir("FCH_DEPOSITO").Value, _
                                             Imprimir("HORA_DEPOSITO").Value, _
                                             Imprimir("COD_USUARIO_ORIGEN").Value, _
                                             Imprimir("NOMB_ORIGEN").Value, _
                                             "" & Imprimir("MAQUINA").Value, _
                                             Imprimir("COD_USUARIO_DESTINO").Value, _
                                             Imprimir("NOMB_DESTINO").Value
        End If
    End If

    If Imprimir("IMP_DOLARES").Value > 0 Then
        If MsgBox("Se va a imprimir los Montos en Dolares - Colocar el formato", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            objPrinter.Imprime_Remesas_Dolares objUsuario.CodigoEmpresa, _
                                             Imprimir("COD_LOCAL").Value, _
                                             Imprimir("COD_LIQUIDACION").Value, _
                                             Imprimir("COD_REMESA").Value, _
                                             Format(Imprimir("IMP_DOLARES").Value, "###,###0.00"), _
                                             Imprimir("FCH_DEPOSITO").Value, _
                                             Imprimir("HORA_DEPOSITO").Value, _
                                             Imprimir("COD_USUARIO_ORIGEN").Value, _
                                             Imprimir("NOMB_ORIGEN").Value, _
                                             "" & Imprimir("MAQUINA").Value, _
                                             Imprimir("COD_USUARIO_DESTINO").Value, _
                                             Imprimir("NOMB_DESTINO").Value
        End If
    End If
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsRemesa_Impresion", Err.Description
End Function
