VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCajeroExpress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strTipoOperacion As String
Private strNumeroTarjeta As String
Private strNumeroTransaccion As String
Private dblImporte As Double
Private intMoneda As Integer
Private dbltipocambio As Double
Private objInterfase As Object

Private byteNumElementosFormaPago As Byte
Private byteNumElementosDatoAdic As Byte
Dim strSecCaja As String
Dim strCodLiquidacion As String



Public Property Get TipoOperacion() As String
    TipoOperacion = strTipoOperacion
End Property

Public Property Let TipoOperacion(ByVal lstrTipoOperacion As String)
    strTipoOperacion = lstrTipoOperacion
End Property


Public Property Get NumeroTarjeta() As String
    NumeroTarjeta = strNumeroTarjeta
End Property

Public Property Let NumeroTarjeta(ByVal lstrNumeroTarjeta As String)
    strNumeroTarjeta = lstrNumeroTarjeta
End Property

Public Property Get NumeroTransaccion() As String
    NumeroTransaccion = strNumeroTransaccion
End Property

Public Property Let NumeroTransaccion(ByVal lstrNumeroTransaccion As String)
    strNumeroTransaccion = lstrNumeroTransaccion
End Property

Public Property Get Importe() As Double
    Importe = dblImporte
End Property

Public Property Let Importe(ByVal ldblImporte As Double)
    dblImporte = ldblImporte
End Property

Public Property Get Moneda() As Integer
    Moneda = intMoneda
End Property

Public Property Let Moneda(ByVal lintMoneda As Integer)
    intMoneda = lintMoneda
End Property

Public Property Get TipoCambio() As Double
    TipoCambio = dbltipocambio
End Property

Public Property Let TipoCambio(ByVal ldblTipoCambio As Double)
    dbltipocambio = ldblTipoCambio
End Property



Public Function b_ActivaPOS(ByVal pinpParm As String, ByRef poutParm As String) As Integer

    On Error GoTo CtrlError

    Set objInterfase = CreateObject("HCC.clsHCC")
    b_ActivaPOS = objInterfase.b_ActivaPOS(pinpParm, poutParm)
    
    ''poutParm = "operacion=ABONO,moneda=SOLES,amount=10200,card=9715000900932820,approval_code=P96037"
    ''b_ActivaPOS = 0
    
    
    GrabaLogCajeroExpress gclsOracle.ODataBase, objUsuario.CodigoEmpresa, objUsuario.CodigoLocal, objUsuario.Codigo, poutParm, b_ActivaPOS
    
    
    If b_ActivaPOS = 0 Then Call CargaValores(poutParm)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsCajeroExpress.b_ActivaPOS", Err.Description

End Function

Public Function b_CierraTurnoPOS(ByVal pinpParm As String) As Integer
    On Error GoTo CtrlError
    Set objInterfase = CreateObject("HCC.clsHCC")
    
    b_CierraTurnoPOS = objInterfase.b_CierraTurnoPOS(pinpParm)
    
    Exit Function
CtrlError:
    Err.Raise Err.Number, "clsCajeroExpress.b_CierraTurnoPOS", Err.Description
    
End Function

Private Sub CargaValores(ByVal lsTrama As String)
Dim strTexto1 As String
Dim strTexto2 As String
Dim strTexto3 As String
Dim strTexto4 As String
Dim strTexto5 As String

        strTexto1 = Replace(lsTrama, "operacion=", "")
        TipoOperacion = Mid(strTexto1, 1, InStr(strTexto1, ",") - 1)
        If TipoOperacion = "OTROS" Then
            strTexto2 = Replace(Mid(strTexto1, InStr(strTexto1, ",") + 1), "moneda=", "")
            Moneda = IIf(Mid(strTexto2, 1, InStr(strTexto2, ",") - 1) = "SOLES", 0, 1)
            strTexto3 = Replace(Mid(strTexto2, InStr(strTexto2, ",") + 1), "card=", "")
            If strTexto3 <> "0000000000000000" Then
            NumeroTarjeta = Mid(strTexto3, 1, InStr(strTexto3, ",") - 1)
                strTexto4 = Replace(Mid(strTexto3, InStr(strTexto3, ",") + 1), "approval_code=", "")
                NumeroTransaccion = Mid(strTexto4, 1)
            End If
            Importe = 0
        Else
            strTexto2 = Replace(Mid(strTexto1, InStr(strTexto1, ",") + 1), "moneda=", "")
            Moneda = IIf(Mid(strTexto2, 1, InStr(strTexto2, ",") - 1) = "SOLES", 0, 1)
            strTexto3 = Replace(Mid(strTexto2, InStr(strTexto2, ",") + 1), "amount=", "")
            Importe = Val(Mid(strTexto3, 1, InStr(strTexto3, ",") - 1)) / 100
            strTexto4 = Replace(Mid(strTexto3, InStr(strTexto3, ",") + 1), "card=", "")
            NumeroTarjeta = Mid(strTexto4, 1, InStr(strTexto4, ",") - 1)
            strTexto5 = Replace(Mid(strTexto4, InStr(strTexto4, ",") + 1), "approval_code=", "")
            NumeroTransaccion = Mid(strTexto5, 1)
        End If
End Sub
Public Function ListaTipoOperacion(ByVal CodEntidad As String, Optional ByVal CodTipoOperacion As String = "") As oraDynaset
    Set ListaTipoOperacion = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO.FN_LISTA_TIPO_OPERACION", 0, CodEntidad, CodTipoOperacion)
End Function

Public Function intValida(ByVal strNumTarjeta As String) As Integer
    On Error GoTo CtrlErr
        intValida = gclsOracle.FN_Valor("BTLCADENA.PKG_PAGO_tARJETA.FN_VALIDA_TARJETA ", strNumTarjeta)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, Err.Description
End Function
Public Function strNombre(ByVal strNumTarjeta As String, ByVal strTipo As String) As String
    On Error GoTo CtrlErr
        strNombre = gclsOracle.FN_Valor("BTLCADENA.PKG_PAGO_tARJETA.FN_RET_NOMBRE ", strNumTarjeta, strTipo)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, Err.Description
End Function
'Public Function MGR_sGetINI(sINIFile As String, sSection As String, sKey _
'        As String, sDefault As String) As String
'
'    Dim strTemp As String * 256
'    Dim intLength As Integer
'
'    strTemp = Space$(256)
'    intLength = GetPrivateProfileString(sSection, sKey, sDefault, strTemp, 255, sINIFile)
'    MGR_sGetINI = Left$(strTemp, intLength)
'
'End Function

Public Sub OldActCajeroExpress(ByVal oradb As OraDatabase, _
                                ByVal CodTipoOperacion As String, _
                                ByVal CodEntidad As String, _
                                ByVal NumTarjeta As String, _
                                ByVal CodAutorizacion As String, _
                                ByVal Importe As Double, _
                                ByVal CodMoneda As String, _
                                ByVal TipoCambio As Double, _
                                ByVal CodBtl As String, _
                                ByVal CodUsuario As String, _
                                ByVal CodMaquina As String, _
                                ByVal CodArqueo As String, _
                                ByVal CodTarjeta As String)




On Error GoTo Error


oradb.Parameters.Remove "A_COD_TIPO_OPERACION"
oradb.Parameters.Remove "A_COD_ENTIDAD"
oradb.Parameters.Remove "A_NUM_TARJETA"
oradb.Parameters.Remove "A_COD_AUTORIZACION"
oradb.Parameters.Remove "A_IMPORTE"
oradb.Parameters.Remove "A_COD_MONEDA"
oradb.Parameters.Remove "A_TIPO_CAMBIO"
oradb.Parameters.Remove "A_COD_BTL"
oradb.Parameters.Remove "A_COD_USUARIO"
oradb.Parameters.Remove "A_COD_MAQUINA"
oradb.Parameters.Remove "A_COD_ARQUEO"
oradb.Parameters.Remove "A_COD_TARJETA"
   
    
oradb.Parameters.Add "A_COD_TIPO_OPERACION", "", ORAPARM_INPUT
oradb.Parameters.Add "A_COD_ENTIDAD", "", ORAPARM_INPUT
oradb.Parameters.Add "A_NUM_TARJETA", "", ORAPARM_INPUT
oradb.Parameters.Add "A_COD_AUTORIZACION", "", ORAPARM_INPUT
oradb.Parameters.Add "A_IMPORTE", "", ORAPARM_INPUT
oradb.Parameters.Add "A_COD_MONEDA", "", ORAPARM_INPUT
oradb.Parameters.Add "A_TIPO_CAMBIO", "", ORAPARM_INPUT
oradb.Parameters.Add "A_COD_BTL", "", ORAPARM_INPUT
oradb.Parameters.Add "A_COD_USUARIO", "", ORAPARM_INPUT
oradb.Parameters.Add "A_COD_MAQUINA", "", ORAPARM_INPUT
oradb.Parameters.Add "A_COD_ARQUEO", "", ORAPARM_INPUT
oradb.Parameters.Add "A_COD_TARJETA", "", ORAPARM_INPUT
                    
oradb.Parameters("A_COD_TIPO_OPERACION").Value = CodTipoOperacion
oradb.Parameters("A_COD_ENTIDAD").Value = CodEntidad
oradb.Parameters("A_NUM_TARJETA").Value = NumTarjeta
oradb.Parameters("A_COD_AUTORIZACION").Value = CodAutorizacion
oradb.Parameters("A_IMPORTE").Value = Importe
oradb.Parameters("A_COD_MONEDA").Value = CodMoneda
oradb.Parameters("A_TIPO_CAMBIO").Value = TipoCambio
oradb.Parameters("A_COD_BTL").Value = CodBtl
oradb.Parameters("A_COD_USUARIO").Value = CodUsuario
oradb.Parameters("A_COD_MAQUINA").Value = CodMaquina
oradb.Parameters("A_COD_ARQUEO").Value = CodArqueo
oradb.Parameters("A_COD_TARJETA").Value = CodTarjeta


                    
   oradb.ExecuteSQL "BEGIN BTLCADENA.PKG_QUERYS_SERVICIOS.SP_ACT_CAJERO_EXPRESS(" & _
    ":A_COD_TIPO_OPERACION,:A_COD_ENTIDAD,:A_NUM_TARJETA,:A_COD_AUTORIZACION,:A_IMPORTE,:A_COD_MONEDA," & _
    ":A_TIPO_CAMBIO,:A_COD_BTL,:A_COD_USUARIO,:A_COD_MAQUINA, :A_COD_ARQUEO, :A_COD_TARJETA); END;"
                    
                    
   Exit Sub
                    
   
Error:
    Err.Raise Err.Number, "clsServicios.ActCajeroExpress", Err.Description


End Sub


Public Function GrabaCajeroCorresponsal(ByRef oTipDoc As String, _
                    ByRef oNumDoc As String) As String
    
On Error GoTo CtrlErr




    Dim i As Integer
    
    On Error GoTo CtrlErr
    
    Dim gvarValores  As Variant
    Dim gvarIO  As Variant
    
    
'''''           A_COD_CIA                 MEDCO.CIA.CIA%TYPE                                            ,
'''''           A_COD_LOCAL                  BTLPROD.CAB_DOCUMENTO.COD_LOCAL%TYPE                ,
'''''           A_COD_MAQUINA                BTLPROD.REL_DOCUMENTO_MAQUINA.COD_MAQUINA%TYPE      ,
'''''           A_COD_MODALIDAD_VENTA        BTLPROD.CAB_DOCUMENTO.COD_MODALIDAD_VENTA%TYPE      ,
'''''           A_COD_TIPO_VENTA             BTLPROD.CAB_DOCUMENTO.COD_TIPO_VENTA%TYPE         ,--Delivery/Local/Internet
'''''           A_COD_TIP_DOC_VTA            BTLPROD.CAB_DOCUMENTO.COD_TIPO_DOCUMENTO%TYPE       ,
'''''           A_IMP_TIP_CAMBIO             BTLPROD.CAB_DOCUMENTO.IMP_TIP_CAMBIO%TYPE   ,
'''''           A_COD_USUARIO                BTLPROD.CAB_DOCUMENTO.COD_USUARIO%TYPE              ,
'''''           A_COD_TIPO_OPERACION         CMR.MAE_TIPO_OPERACION.COD_TIPO_OPERACION%TYPE              ,
'''''           A_COD_FORM_PAGO              BTLPROD.DET_FORPAG_DOC.COD_FORMA_PAGO%TYPE ,
'''''           A_COD_FPAG_HIJO              BTLPROD.DET_FORPAG_DOC.COD_HIJO%TYPE ,
'''''           A_COD_MONEDA                 BTLPROD.DET_FORPAG_DOC.COD_MONEDA%TYPE ,
'''''           A_COD_TIPO_TARJETA             BTLPROD.DET_FORPAG_DOC.COD_TIPO_TARJETA%TYPE ,
'''''           A_NUM_TARJETA                BTLPROD.DET_FORPAG_DOC.NUM_TARJETA%TYPE ,
'''''           A_NUM_AUTORIZACION           BTLPROD.DET_FORPAG_DOC.NUM_AUTORIZACION%TYPE,
'''''           A_COD_BANCO                  BTLPROD.DET_FORPAG_DOC.COD_BANCO%TYPE,
'''''           A_IMPORTE                    BTLPROD.DET_FORPAG_DOC.IMP_MONEDA_NAC%TYPE,
'''''           V_TIPDOC_GENERAD             OUT CHAR,
'''''           V_NUMDOC_GENERAD             OUT CHAR
    
    
    
    gvarValores = Array(objUsuario.CodigoEmpresa, _
                        objUsuario.CodigoLocal, _
                        objUsuario.NombrePC, _
                        Format(Cajero_Corresponsal, "000"), _
                        objUsuario.CodTipoVenta, _
                        gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_TIP_DOC_VOUCHER"), _
                        objUsuario.TipoCambio, _
                        objUsuario.Codigo, _
                        TipoOperacion, _
                        gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_COD_FORM_PAGO_CAJERO"), _
                        gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_COD_FPAGO_HIJO_CAJERO"), _
                        Moneda, _
                        gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_COD_TIP_TARJE_CAJERO"), _
                        NumeroTarjeta, _
                        NumeroTransaccion, _
                        gclsOracle.Const_Val("BTLPROD.PKG_CONSTANTES.CONS_COD_BANCO_CAJERO_CORRES"), _
                        Importe, _
                        oTipDoc, _
                        oNumDoc)
    
    gvarIO = Array(Entrada, _
                   Entrada, _
                   Entrada, _
                   Entrada, _
                   Entrada, _
                   Entrada, _
                   Entrada, _
                   Entrada, _
                   Entrada, _
                   Entrada, _
                   Entrada, _
                   Entrada, _
                   Entrada, _
                   Entrada, _
                   Entrada, _
                   Entrada, _
                   Entrada, _
                   Salida, _
                   Salida)
                   
    GrabaCajeroCorresponsal = gclsOracle.SP("BTLPROD.PKG_DOCUMENTO.SP_GRABA_CAJERO_CORRESPONSAL", _
                              gvarValores, _
                              gvarIO)
    
    
    oTipDoc = gvarValores(17)
    oNumDoc = gvarValores(18)
    




Exit Function

CtrlErr:
    Err.Raise Err.Number, "clsCajeroExpress.GrabaCajeroCorresponsal", Err.Description


End Function



Public Sub GrabaLogCajeroExpress(ByVal oradb As OraDatabase, _
                                ByVal CodCia As String, _
                                ByVal CodLocal As String, _
                                ByVal CodUsuario As String, _
                                ByVal ObsTrama As String, _
                                ByVal NumActivaPOS As Integer)




On Error GoTo Error


oradb.Parameters.Remove "A_CIA"
oradb.Parameters.Remove "A_COD_LOCAL"
oradb.Parameters.Remove "A_COD_USUARIO"
oradb.Parameters.Remove "A_OBS_TRAMA"
oradb.Parameters.Remove "A_NUM_ACTIVA_POS"
   
    
oradb.Parameters.Add "A_CIA", "", ORAPARM_INPUT
oradb.Parameters.Add "A_COD_LOCAL", "", ORAPARM_INPUT
oradb.Parameters.Add "A_COD_USUARIO", "", ORAPARM_INPUT
oradb.Parameters.Add "A_OBS_TRAMA", "", ORAPARM_INPUT
oradb.Parameters.Add "A_NUM_ACTIVA_POS", "", ORAPARM_INPUT
                    
oradb.Parameters("A_CIA").Value = CodCia
oradb.Parameters("A_COD_LOCAL").Value = CodLocal
oradb.Parameters("A_COD_USUARIO").Value = CodUsuario
oradb.Parameters("A_OBS_TRAMA").Value = ObsTrama
oradb.Parameters("A_NUM_ACTIVA_POS").Value = NumActivaPOS


                    
   oradb.ExecuteSQL "BEGIN BTLPROD.PKG_DOCUMENTO.SP_GRABA_LOG_CAJERO_EXPRESS(" & _
                ":A_CIA,:A_COD_LOCAL,:A_COD_USUARIO,:A_OBS_TRAMA,:A_NUM_ACTIVA_POS); END;"
                    
                    
   Exit Sub
                    
   
Error:
    Err.Raise Err.Number, "clsCajeroExpress.GrabaLogCajeroExpress", Err.Description


End Sub


Public Property Get SecCaja() As String
    SecCaja = strSecCaja
    
End Property

Public Property Let SecCaja(ByVal lstrSecCaja As String)
    strSecCaja = lstrSecCaja
End Property





Public Property Get CodLiquidacion() As String
    CodLiquidacion = strCodLiquidacion
End Property

Public Property Let CodLiquidacion(ByVal lstrCodLiquidacion As String)
    strCodLiquidacion = CodLiquidacion
End Property
