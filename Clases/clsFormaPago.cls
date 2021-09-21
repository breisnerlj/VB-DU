VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private strSmbMoneda As String
Private strDesLgMoneda As String
Private xTarjVerif As New XArrayDB '** Variable arreglo para la verificaion de tarjeta

Private Sub Class_Initialize()
    xTarjVerif.ReDim 0, -1, 0, 11 'Setea el array de las tarjetas
End Sub

Public Property Get TarjetaVerif() As XArrayDB
    Set TarjetaVerif = xTarjVerif 'Devuelve la variable de la clase en forma de XArray para las tarjetas
End Property

Public Function Lista(Optional Codigo As String = "") As oraDynaset
    On Error GoTo Pase
    
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_FORMA_PAGO.FN_LISTA", 0, Codigo)
    
    Exit Function
Pase:
    Err.Raise Err.Number, "clsFormaPago", Err.Description
End Function

Public Function ListaFormaPDocumento(TipoDocumento As String, ByVal NumeroDocumento As String) As oraDynaset
    On Error GoTo Pase
    Set ListaFormaPDocumento = gclsOracle.FN_Cursor("BTLPROD.PKG_FORMA_PAGO.fn_lista_fpago_documento", 0, TipoDocumento, NumeroDocumento)
    Exit Function
Pase:
    Err.Raise Err.Number, "clsFormaPago", Err.Description
End Function

'-- Efectivo --'
Public Function ListaHijo(ByVal vstrCodFP As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaHijo = gclsOracle.FN_Cursor("BTLPROD.PKG_FORMA_PAGO.FN_LISTA_HIJO", 0, vstrCodFP)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago", Err.Description
End Function

Public Function ListaHijo2(ByVal vstrCodFP As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaHijo2 = gclsOracle.FN_Cursor("BTLPROD.PKG_FORMA_PAGO.FN_LISTA_HIJO2", 0, vstrCodFP)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago", Err.Description
End Function

'-- Valida las Tarjetas de Credito --'
Public Function ValidaTarjeta(ByVal vstrNumTarj As String, _
                                  ByVal vstrTipo As String) As String
    On Error GoTo CtrlErr
'''    Dim strRespuesta As String
'''    strRespuesta = Me.TarjetaBloqueada(vstrNumTarj)
'''    If strRespuesta = "" Then
    'ValidaTarjeta = gclsOracle.FN_Valor("BTLCADENA.PKG_PAGO_TARJETA.FN_RET_NOMBRE", vstrNumTarj, vstrTipo)
    ValidaTarjeta = gclsOracle.FN_Valor("BTLPROD.PKG_TARJETA.FN_LISTA_TJT_NUMERO", vstrNumTarj, vstrTipo)
''''    Else
''''        ValidaTarjeta = strRespuesta
''''    End If
    
    Exit Function
CtrlErr:
    If ValidaTarjeta = "" Then Exit Function
    Err.Raise Err.Number, "clsFormaPago", Err.Description
End Function

'-- Tipo de Cuota --'
Public Function TipoCuota() As oraDynaset
    On Error GoTo CtrlErr
    Set TipoCuota = gclsOracle.FN_Cursor("BTLPROD.PKG_FORMA_PAGO.FN_TIPO_CUOTA_TARJ", 0)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago", Err.Description
End Function

'-- Lista Moneda --'
Public Function ListaMoneda(Optional CodMoneda As String = "") As oraDynaset
    On Error GoTo CtrlErr
    Set ListaMoneda = gclsOracle.FN_Cursor("BTLPROD.PKG_FORMA_PAGO.FN_LISTA_MONEDA", 0, CodMoneda)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago", Err.Description
End Function

'-- lista redondeo -- '
Public Function ListaRedondeo() As oraDynaset
    On Error GoTo CtrlErr
    Set ListaRedondeo = gclsOracle.FN_Cursor("BTLPROD.PKG_FORMA_PAGO.FN_LISTA_REDONDEO", 0)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago", Err.Description
End Function


'-- lista donacion -- '
Public Function ListaDonacion()
    On Error GoTo CtrlErr
    Set ListaDonacion = gclsOracle.FN_Cursor("BTLPROD.PKG_FORMA_PAGO.FN_LISTA_DONACION", 0)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago", Err.Description
End Function

'-- Verfificaciond e Tarjeta --'
Public Function ListaTarjxVerif(ByVal vstrCia As String, _
                                ByVal vstrCodLocal As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaTarjxVerif = gclsOracle.FN_Cursor("BTLPROD.PKG_TARJETA.FN_LISTA_TARJ_X_VERIF", 0, vstrCia, vstrCodLocal)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago", Err.Description
End Function

'-- Lista tarjetas de cliente pendiente de verificacion --'
Public Function LstDetTarxVerif(ByVal vstrCia As String, _
                                ByVal vstrCodLocal As String, _
                                ByVal vstrNumProforma As String, _
                                ByVal vstrCodCliente As String) As oraDynaset
    On Error GoTo CtrlErr
     Set LstDetTarxVerif = gclsOracle.FN_Cursor("BTLPROD.PKG_TARJETA.FN_DET_TARJ_X_VERIF", 0, _
                                                                                         vstrCia, _
                                                                                         vstrCodLocal, _
                                                                                         vstrNumProforma, _
                                                                                         vstrCodCliente)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago", Err.Description
End Function

Public Function GrabaNumAut(ByVal vstrCia As String, _
                            ByVal vstrCodLocal As String, _
                            ByVal vstrNumProforma As String, _
                            ByVal CodTarjeta As String, _
                            ByVal vstrNumTarj As String, _
                            ByVal vstrNumAutorizacion As String) As String

    Dim varValores As Variant
    Dim varIO As Variant
    
    varValores = Array(vstrCia, _
                       vstrCodLocal, _
                       vstrNumProforma, _
                       CodTarjeta, _
                       vstrNumTarj, _
                       vstrNumAutorizacion)
    
    varIO = Array(entrada, _
                  entrada, _
                  entrada, _
                  entrada, _
                  entrada, _
                  entrada)
    
    GrabaNumAut = gclsOracle.SP("BTLPROD.PKG_TARJETA.SP_GRABA_NUM_AUT_TARJ", _
                                                           varValores, varIO)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago", Err.Description

End Function


Public Function VerfirificaTarjeta(ByVal vstrCia As String, _
                                   ByVal vstrCodLocal As String, _
                                   ByVal vstrNumProforma As String, _
                                   ByVal Usuario As String, _
                                   ByVal vstrCadNumAut As String, _
                                   ByVal vstrCadCodHijoFP As String, _
                                   ByVal vstrCodMotRech As String, _
                                   ByVal vstrCadFlgVerifPos As String _
                                  ) As String

    Dim varValores As Variant
    Dim varIO As Variant
    
    varValores = Array(vstrCia, _
                       vstrCodLocal, _
                       vstrNumProforma, _
                       Usuario, _
                       vstrCadNumAut, _
                       vstrCadCodHijoFP, _
                       vstrCodMotRech, _
                       vstrCadFlgVerifPos _
                       )
    
    varIO = Array(entrada, _
                  entrada, _
                  entrada, _
                  entrada, _
                  entrada, _
                  entrada, _
                  entrada, _
                  entrada _
                  )
    
    VerfirificaTarjeta = gclsOracle.SP("BTLPROD.PKG_TARJETA.SP_VERIFICA_TARJ", _
                                                           varValores, varIO)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago", Err.Description

End Function

'*************************************************************************************
' Metodo para agregar un detalle de tarjetas verificacion al arreglo
'*************************************************************************************
Public Function AddDetTarj(ByVal vstrCodTarjeta$, _
                           ByVal vstrDesTarjeta$, _
                           ByVal vstrNumTarjeta$, _
                           ByVal vstrNumAutorizacion$, _
                           ByVal vstrFchVencimiento$, _
                           ByVal vstrNumCuotas$, _
                           ByVal vstrPagoDif$, _
                           ByVal vdblImporte#, _
                           ByVal vstrEmpPOS$, _
                           Optional vstrFlfVerifPos$) As XArrayDB
    Dim ultimo As Integer ' declara variable contador
    Dim aux As Integer
    If xTarjVerif.Count(1) < 0 Then Exit Function
    
        xTarjVerif.AppendRows
        xTarjVerif(xTarjVerif.UpperBound(1), 0) = vstrCodTarjeta
        xTarjVerif(xTarjVerif.UpperBound(1), 1) = vstrDesTarjeta
        xTarjVerif(xTarjVerif.UpperBound(1), 2) = vstrNumTarjeta
        xTarjVerif(xTarjVerif.UpperBound(1), 3) = Empty
        xTarjVerif(xTarjVerif.UpperBound(1), 4) = vstrFchVencimiento
        xTarjVerif(xTarjVerif.UpperBound(1), 5) = vstrNumCuotas
        xTarjVerif(xTarjVerif.UpperBound(1), 6) = vstrPagoDif
        xTarjVerif(xTarjVerif.UpperBound(1), 7) = vdblImporte
        xTarjVerif(xTarjVerif.UpperBound(1), 8) = vstrEmpPOS
        xTarjVerif(xTarjVerif.UpperBound(1), 10) = vstrFlfVerifPos
    Set AddDetTarj = xTarjVerif

End Function

'****************************************************'
'- Metodo Para Cargar la remesas por asignar remito -'
'****************************************************'
Public Sub LoadDetTarVerif(ByVal vstrNumProforma As String, _
                           ByVal vstrCodCliente As String)

    Dim odynTarjVerif As oraDynaset
    Dim objFormaPago As New clsFormaPago
 
    On Error GoTo CtrlErr
    Set odynTarjVerif = objFormaPago.LstDetTarxVerif(objUsuario.CodigoEmpresa, _
                                                     objUsuario.CodigoLocal, _
                                                     vstrNumProforma, _
                                                     vstrCodCliente)
                                                
    
    xTarjVerif.ReDim 0, -1, 0, 10 'Setea el array de las tarjetas
    
    odynTarjVerif.MoveFirst
    While Not odynTarjVerif.EOF
    
       AddDetTarj odynTarjVerif("COD_TIPO_TARJETA").Value, _
                  odynTarjVerif("DES_TIPO_TARJETA").Value, _
                  odynTarjVerif("NUM_TARJETA").Value, _
                  "" & odynTarjVerif("NUM_AUTORIZACION").Value, _
                  odynTarjVerif("FCH_VENCIMIENTO").Value, _
                  odynTarjVerif("NUM_CUOTAS").Value, _
                  odynTarjVerif("DES_CUOTA").Value, _
                  odynTarjVerif("IMP_MONEDA_NAC").Value, _
                  odynTarjVerif("EMP_POS").Value, _
                  "" & odynTarjVerif("VERIF_POS").Value
       odynTarjVerif.MoveNext
    Wend
    frm_DLV_Verificacion_Tarjeta.grdTarjetas.Rebind
    
    'Set objFormaPago = Nothing
    Exit Sub
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago", Err.Description
    
End Sub

Public Function ListaFPagoTipMaquina(Optional CodTipoMaquina As String = "", Optional flgModifica As String = "") As oraDynaset
    On Error GoTo Pase
    
    Set ListaFPagoTipMaquina = gclsOracle.FN_Cursor("BTLPROD.PKG_FORMA_PAGO.FN_LISTA_FPAGO_TIPMAQUINA", 0, CodTipoMaquina, flgModifica)
    
    Exit Function
Pase:
    Err.Raise Err.Number, "clsFormaPago.ListaFPagoTipMaquina", Err.Description
End Function

Public Function Graba(ByVal TipoDocumento As String, ByVal NumeroDocumento As String, ByVal NumeroSecuencia As String, ByVal MontoTotal As String, ByVal CodigoUsuario As String) As String
                      
    Dim NumElementosFormaPago As Integer
    Dim i As Integer
    NumElementosFormaPago = IIf(objVenta.FormaPago.UpperBound(1) < 0, 0, objVenta.FormaPago.UpperBound(1)) + 1
    
    
    For i = gclsOracle.ODataBase.Parameters.Count - 1 To 0 Step -1
        gclsOracle.ODataBase.Parameters.Remove i
    Next
    
    gclsOracle.ODataBase.Parameters.Add "A_COD_CIA", objUsuario.CodigoEmpresa, ORAPARM_INPUT, ORATYPE_VARCHAR2 '1
    gclsOracle.ODataBase.Parameters.Add "A_COD_LOCAL", objUsuario.CodigoLocal, ORAPARM_INPUT, ORATYPE_VARCHAR2 '2
    gclsOracle.ODataBase.Parameters.Add "A_COD_CLIENTE", "", ORAPARM_INPUT, ORATYPE_VARCHAR2 '3
    gclsOracle.ODataBase.Parameters.AddTable "V_TIPDOC_GENERAD", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '4
    gclsOracle.ODataBase.Parameters.AddTable "V_NUMDOC_GENERAD", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '5
    gclsOracle.ODataBase.Parameters.AddTable "V_TIPDOC_GENERAD_CO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '6
    gclsOracle.ODataBase.Parameters.AddTable "V_NUMDOC_GENERAD_CO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '7
    gclsOracle.ODataBase.Parameters.Add "V_SUBTOT_TOTAL", Val(MontoTotal), ORAPARM_INPUT, ORATYPE_VARCHAR2 '8
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_COD_FORM_PAGO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '9
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_COD_FPAG_HIJO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '10
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_COD_MONEDA_PAGO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '11
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_TIP_CAMBIO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '12
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_COD_TIP_TARJETA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '13
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_NUM_TARJETA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '14
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_NUM_AUTORIZACION", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '15
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_FLG_TIP_CUOTA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '16
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_NUM_CUOTA", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '17
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_FCH_VENCIMIENTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '18
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_MTO_RETIRO_EFEC", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '19
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_COD_DOC_PAGO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '20
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_NUM_DOC_PAGO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '21
    gclsOracle.ODataBase.Parameters.AddTable "A_NUM_DOC_NOTA_CRED", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '22
    gclsOracle.ODataBase.Parameters.AddTable "A_FCH_DOC_NOTA_CRED", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '23
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_COD_DOC_IDENT", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '24
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_NUM_DOC_IDENT", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '25
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_COD_BANCO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '26
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_NUM_MOVIMIENTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '27
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_FCH_MOVIMIENTO", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '28
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_CTA_CTE_BTL", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '29
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_COD_DONACION", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '30
    gclsOracle.ODataBase.Parameters.AddTable "A_CAD_IMPORTE", ORAPARM_INPUT, ORATYPE_VARCHAR2, NumElementosFormaPago, 200 '31
    gclsOracle.ODataBase.Parameters.Add "A_FLG_ES_CREDITO", "", ORAPARM_OUTPUT, ORATYPE_VARCHAR2 '32
    gclsOracle.ODataBase.Parameters.Add "A_NUM_SECUENCIA", NumeroSecuencia, ORAPARM_INPUT, ORATYPE_VARCHAR2 '33
    gclsOracle.ODataBase.Parameters.Add "A_COD_USUARIO", CodigoUsuario, ORAPARM_INPUT, ORATYPE_VARCHAR2 '1
    
    '1
    '2
    '3
    Dim TIPDOC_GENERAD As OracleInProcServer.OraParamArray '4
    Dim NUMDOC_GENERAD As OracleInProcServer.OraParamArray '5
    Dim TIPDOC_GENERAD_CO As OracleInProcServer.OraParamArray '6
    Dim NUMDOC_GENERAD_CO As OracleInProcServer.OraParamArray '7
    '8
    Dim CodFormPago As OracleInProcServer.OraParamArray '9
    Dim CodFPagHijo As OracleInProcServer.OraParamArray '10
    Dim CodMonedaPago As OracleInProcServer.OraParamArray '11
    Dim TipCambio As OracleInProcServer.OraParamArray '12
    Dim CodTipTarjeta As OracleInProcServer.OraParamArray '13
    Dim NumTarjeta As OracleInProcServer.OraParamArray '14
    Dim NumAutorizacion As OracleInProcServer.OraParamArray '15
    Dim FlgTipCuota As OracleInProcServer.OraParamArray '16
    Dim NumCuota As OracleInProcServer.OraParamArray '17
    Dim FchVencimiento As OracleInProcServer.OraParamArray '18
    Dim MtoRetiroEfec As OracleInProcServer.OraParamArray '19
    Dim CodDocPago As OracleInProcServer.OraParamArray '20
    Dim NumDocPago As OracleInProcServer.OraParamArray '21
    Dim NumDocNotaCred As OracleInProcServer.OraParamArray '22
    Dim FchDocNotaCred As OracleInProcServer.OraParamArray '23
    Dim CodDocIdent As OracleInProcServer.OraParamArray '24
    Dim NumDocIdent As OracleInProcServer.OraParamArray '25
    Dim CodBanco As OracleInProcServer.OraParamArray '26
    Dim NumMovimiento As OracleInProcServer.OraParamArray '27
    Dim FchMovimiento As OracleInProcServer.OraParamArray '28
    Dim CtaCteBtl As OracleInProcServer.OraParamArray '29
    Dim CodDonacion As OracleInProcServer.OraParamArray '30
    Dim Importe As OracleInProcServer.OraParamArray '31
    '32
    '33
    'Dim NombTitular As OracleInProcServer.OraParamArray '32
    
    '1
    '2
    '3
    Set TIPDOC_GENERAD = gclsOracle.ODataBase.Parameters("V_TIPDOC_GENERAD") '4
    Set NUMDOC_GENERAD = gclsOracle.ODataBase.Parameters("V_NUMDOC_GENERAD") '5
    Set TIPDOC_GENERAD_CO = gclsOracle.ODataBase.Parameters("V_TIPDOC_GENERAD_CO") '6
    Set NUMDOC_GENERAD_CO = gclsOracle.ODataBase.Parameters("V_NUMDOC_GENERAD_CO") '7
    '8
    Set CodFormPago = gclsOracle.ODataBase.Parameters("A_CAD_COD_FORM_PAGO") '9
    Set CodFPagHijo = gclsOracle.ODataBase.Parameters("A_CAD_COD_FPAG_HIJO") '10
    Set CodMonedaPago = gclsOracle.ODataBase.Parameters("A_CAD_COD_MONEDA_PAGO") '11
    Set TipCambio = gclsOracle.ODataBase.Parameters("A_CAD_TIP_CAMBIO") '12
    Set CodTipTarjeta = gclsOracle.ODataBase.Parameters("A_CAD_COD_TIP_TARJETA") '13
    Set NumTarjeta = gclsOracle.ODataBase.Parameters("A_CAD_NUM_TARJETA") '14
    Set NumAutorizacion = gclsOracle.ODataBase.Parameters("A_CAD_NUM_AUTORIZACION")  '15
    Set FlgTipCuota = gclsOracle.ODataBase.Parameters("A_CAD_FLG_TIP_CUOTA") '16
    Set NumCuota = gclsOracle.ODataBase.Parameters("A_CAD_NUM_CUOTA") '17
    Set FchVencimiento = gclsOracle.ODataBase.Parameters("A_CAD_FCH_VENCIMIENTO") '18
    Set MtoRetiroEfec = gclsOracle.ODataBase.Parameters("A_CAD_MTO_RETIRO_EFEC") '19
    Set CodDocPago = gclsOracle.ODataBase.Parameters("A_CAD_COD_DOC_PAGO") '20
    Set NumDocPago = gclsOracle.ODataBase.Parameters("A_CAD_NUM_DOC_PAGO") '21
    Set NumDocNotaCred = gclsOracle.ODataBase.Parameters("A_NUM_DOC_NOTA_CRED") '22
    Set FchDocNotaCred = gclsOracle.ODataBase.Parameters("A_FCH_DOC_NOTA_CRED") '23
    Set CodDocIdent = gclsOracle.ODataBase.Parameters("A_CAD_COD_DOC_IDENT") '24
    Set NumDocIdent = gclsOracle.ODataBase.Parameters("A_CAD_NUM_DOC_IDENT") '25
    Set CodBanco = gclsOracle.ODataBase.Parameters("A_CAD_COD_BANCO") '26
    Set NumMovimiento = gclsOracle.ODataBase.Parameters("A_CAD_NUM_MOVIMIENTO") '27
    Set FchMovimiento = gclsOracle.ODataBase.Parameters("A_CAD_FCH_MOVIMIENTO") '28
    Set CtaCteBtl = gclsOracle.ODataBase.Parameters("A_CAD_CTA_CTE_BTL") '29
    Set CodDonacion = gclsOracle.ODataBase.Parameters("A_CAD_COD_DONACION") '30
    Set Importe = gclsOracle.ODataBase.Parameters("A_CAD_IMPORTE") '31
   ' Set NombTitular = gclsOracle.ODataBase.Parameters("A_NOM_TITULAR") '32
    '32
    '33

For i = 0 To objVenta.FormaPago.UpperBound(1)

        '1
        '2
        '3
        TIPDOC_GENERAD(i) = TipoDocumento '4
        NUMDOC_GENERAD(i) = NumeroDocumento '5
        TIPDOC_GENERAD_CO(i) = "" '6
        NUMDOC_GENERAD_CO(i) = "" '7
        '8
        CodFormPago(i) = objVenta.FormaPago(i, 0) '9
        CodFPagHijo(i) = objVenta.FormaPago(i, 2) '10
        CodMonedaPago(i) = objVenta.FormaPago(i, 6) '11
        TipCambio(i) = objVenta.FormaPago(i, 11) '12
        CodTipTarjeta(i) = objVenta.FormaPago(i, 5) '13
        NumTarjeta(i) = objVenta.FormaPago(i, 12) '14
        NumAutorizacion(i) = objVenta.FormaPago(i, 20) '15
        FlgTipCuota(i) = objVenta.FormaPago(i, 15) '16
        NumCuota(i) = objVenta.FormaPago(i, 13) '17
        FchVencimiento(i) = objVenta.FormaPago(i, 14) '18
        MtoRetiroEfec(i) = objVenta.FormaPago(i, 29) '19
        CodDocPago(i) = objVenta.FormaPago(i, 7) '20
        NumDocPago(i) = objVenta.FormaPago(i, 16) '21
        NumDocNotaCred(i) = objVenta.FormaPago(i, 22) '22
        FchDocNotaCred(i) = objVenta.FormaPago(i, 21) '23
        CodDocIdent(i) = "" '24
        NumDocIdent(i) = objVenta.FormaPago(i, 25) '25
        CodBanco(i) = objVenta.FormaPago(i, 8) '26
        NumMovimiento(i) = objVenta.FormaPago(i, 17) '27
        FchMovimiento(i) = objVenta.FormaPago(i, 19) '28
        CtaCteBtl(i) = objVenta.FormaPago(i, 10) '29
        CodDonacion(i) = objVenta.FormaPago(i, 9) '30
        Importe(i) = objVenta.FormaPago(i, 4) '31
        '32
        '33
    Next
    
    
gclsOracle.ODataBase.ExecuteSQL " BEGIN BTLPROD.PKG_DOCUMENTO.SP_GRABA_DET_FPAGO(:A_COD_CIA, :A_COD_LOCAL, :A_COD_CLIENTE, :V_TIPDOC_GENERAD, :V_NUMDOC_GENERAD, " _
           & ":V_TIPDOC_GENERAD_CO, :V_NUMDOC_GENERAD_CO, :V_SUBTOT_TOTAL, :A_CAD_COD_FORM_PAGO, :A_CAD_COD_FPAG_HIJO," _
           & ":A_CAD_COD_MONEDA_PAGO, :A_CAD_TIP_CAMBIO, :A_CAD_COD_TIP_TARJETA, :A_CAD_NUM_TARJETA, :A_CAD_FLG_TIP_CUOTA," _
           & ":A_CAD_NUM_CUOTA, :A_CAD_FCH_VENCIMIENTO, :A_CAD_MTO_RETIRO_EFEC, :A_CAD_NUM_AUTORIZACION, :A_CAD_COD_DOC_PAGO," _
           & ":A_CAD_NUM_DOC_PAGO, :A_NUM_DOC_NOTA_CRED, :A_FCH_DOC_NOTA_CRED, :A_CAD_COD_DOC_IDENT, :A_CAD_NUM_DOC_IDENT," _
           & ":A_CAD_COD_BANCO, :A_CAD_NUM_MOVIMIENTO, :A_CAD_FCH_MOVIMIENTO, :A_CAD_CTA_CTE_BTL, :A_CAD_COD_DONACION," _
           & ":A_CAD_IMPORTE, :A_FLG_ES_CREDITO, :A_NUM_SECUENCIA, :A_COD_USUARIO) ; END;"
    Graba = ""
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago", Err.Description

End Function



Public Function ListaTotalSec(Optional Secuencia As String = "") As oraDynaset
    On Error GoTo Pase
    
    Set ListaTotalSec = gclsOracle.FN_Cursor("BTLPROD.PKG_FORMA_PAGO.FN_LISTA_TOTAL_DOC", 0, Secuencia)
    
    Exit Function
Pase:
    Err.Raise Err.Number, "clsFormaPago", Err.Description
End Function

Public Function ListaTJTxdlv(ByVal vstrCliente As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaTJTxdlv = gclsOracle.FN_Cursor("BTLPROD.PKG_FORMA_PAGO.FN_LISTA_TARJETAS_X_DLV", 0, vstrCliente)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago.ListaTJTxdlv", Err.Description
End Function


Public Function GrabaTarjetasXDLV(ByVal strCodCliente As String, _
                                   ByVal StrNumTarjeta As String, _
                                   ByVal strCodTipoTarjeta As String, _
                                   ByVal strFchVencimiento As String, _
                                   ByVal strflgActivo As String) As String

    Dim varValores As Variant
    Dim varIO As Variant
    
    varValores = Array(strCodCliente, _
                       StrNumTarjeta, _
                       strCodTipoTarjeta, _
                       strFchVencimiento, _
                       strflgActivo)
    
    varIO = Array(entrada, _
                  entrada, _
                  entrada, _
                  entrada, _
                  entrada)
    
    GrabaTarjetasXDLV = gclsOracle.SP("BTLPROD.PKG_FORMA_PAGO.SP_GRABA_TARJETAS_X_DLV", _
                                                           varValores, varIO)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago.GrabaTarjetasXDLV", Err.Description
End Function

Public Function ListaTarjetasVenta() As oraDynaset
    On Error GoTo CtrlErr
    Set ListaTarjetasVenta = gclsOracle.FN_Cursor("BTLPROD.PKG_FORMA_PAGO.LISTA_TARJETAS_VENTA", 0)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago.ListaTarjetasVenta", Err.Description
End Function



Public Function GrabarBin(ByVal CodigoBines As String, _
                          ByVal CodigoTarjeta As String, _
                          ByVal CodigoUsuario As String, _
                          ByVal NuevoBin As String) As String

    Dim varValores As Variant
    Dim varIO As Variant
    
    varValores = Array(CodigoBines, _
                       CodigoTarjeta, _
                       CodigoUsuario, _
                       NuevoBin)
    
    varIO = Array(entrada, _
                  entrada, _
                  entrada, _
                  entrada)
    
    GrabarBin = gclsOracle.SP("BTLPROD.PKG_FORMA_PAGO.SP_ACTUALIZA_BINES", _
                                                           varValores, varIO)
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago.GrabarBin", Err.Description
End Function

Public Function TarjetaBloqueada(ByVal vstrNumTarj As String) As String
Dim strRepuesta As String
    On Error GoTo CtrlErr
    'ValidaTarjeta = gclsOracle.FN_Valor("BTLCADENA.PKG_PAGO_TARJETA.FN_RET_NOMBRE", vstrNumTarj, vstrTipo)
    strRepuesta = "" & gclsOracle.FN_Valor("BTLPROD.PKG_TARJETA.fn_tarjeta_BLOQUEADA", vstrNumTarj)
    TarjetaBloqueada = strRepuesta
    Exit Function
CtrlErr:
    If TarjetaBloqueada = "" Then Exit Function
    Err.Raise Err.Number, "clsFormaPago", Err.Description
End Function

Public Function ValidaMOD10(ByVal vstrNumTarj As String, ByVal flgValida As String) As String
    On Error GoTo CtrlErr
    ValidaMOD10 = gclsOracle.FN_Valor("BTLPROD.PKG_TARJETA.fn_lista_tjt_numerom10", vstrNumTarj, "", flgValida)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago", Err.Description
End Function

Public Function ValidaVoucher(ByVal vstrCia As String, vstrCodLocal As String, ByVal vstrCodMaquina As String, ByVal vstrCodUsuario As String, ByVal vstrNumTarj As String, ByVal strNumAutorizacion As String) As Integer
    On Error GoTo CtrlErr
    ValidaVoucher = gclsOracle.FN_Valor("BTLPROD.PKG_TARJETA.F_VALIDA_VOUCHER", vstrCia, vstrCodMaquina, vstrCodLocal, vstrCodUsuario, vstrNumTarj, strNumAutorizacion)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago.ValidaVoucher", Err.Description
End Function


Public Function ListaCupon(ByVal strBusqueda As String) As oraDynaset
    On Error GoTo CtrlErr
    Set ListaCupon = gclsOracle.FN_Cursor("BTLPROD.PKG_DOCUMENTO_DSCTO.fn_dev_cupon", 0, strBusqueda)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago.ListaCupon", Err.Description
End Function

Public Function ListaTarjetasDlv() As oraDynaset
    On Error GoTo CtrlErr
    Set ListaTarjetasDlv = gclsOracle.FN_Cursor("BTLPROD.PKG_FORMA_PAGO.FN_LISTA_TARJ_DLV", 0)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsFormaPago.ListaCupon", Err.Description
End Function


