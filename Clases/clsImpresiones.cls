VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImpresiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim objDocumento As New clsDocumento
Dim objLetra As New clsLetras
Dim objRemito As New clsRemito
Dim i%

Dim odynReImp As oraDynaset
Dim strDesEfecSol As String
Dim strDesEfecDol As String
Dim strBanco As String
Dim strPortavalor As String
Dim strCtaCteSol As String
Dim strCtaCteDol As String
'''Autor     : Juan Arturo Escate
'''Fecha     : 25/11/2009
'''Proposito : Se comento porque se implemento un nuevo metodo que no consultase de la pantalla sino lo grabado
'''Public Function Impresion_Prosegur(ByVal vstrCliente As String, _
'''                                   ByVal vstrRecibido As String, _
'''                                   ByVal vstrPara As String, _
'''                                   ByVal vstrLugar As String, _
'''                                   ByVal vstrDesEfecSol As String, _
'''                                   ByVal vstrDesEfecDol As String, _
'''                                   ByVal vstrTotSol As String, _
'''                                   ByVal vstrTotDol As String, _
'''                                   ByVal vstrDireccLocal As String, _
'''                                   ByVal vstrLocalidad As String, _
'''                                   ByVal vstrNumPrecinto As String, _
'''                                   Optional vstrCtdSobre As String)
'''
'''    'Set objPrint = New clsImpresora
'''    'Dim intFormato As Integer
'''    'Dim lngAncho As Double, lngLargo As Double
'''
'''    'lngAncho = 6.9 * 25.4
'''    'lngLargo = 5.9 * 25.4
'''    'intFormato = objPrint.CrearFormato("PP", lngAncho, lngLargo)
'''    'intFormato = lngAncho * lngLargo
'''    'Set objPrint = Nothing
'''    'Printer.PaperSize = intFormato
'''    'Printer.Height = 6.9 * 25.4
'''    'Printer.Width = 5.9 * 25.4
'''
'''    i = 1
'''    For i = 1 To 13
'''        Printer.Print
'''    Next i
'''    Printer.FontName = "Draft 17cpi"
'''    Printer.FontSize = 11
'''    Printer.Print Space(8) & Trim(vstrCtdSobre) & Space(30) & " " & Trim(vstrCliente)
'''    Printer.Print
'''    Printer.Print Space(42) & " " & Trim(vstrRecibido)
'''    Printer.Print
'''    Printer.Print Space(42) & " " & Trim(vstrPara)
'''    Printer.Print
'''    Printer.Print Space(42) & " " & Trim(vstrDireccLocal) & Space(20) & Trim(vstrLocalidad)
'''    Printer.Print
'''    For i = 1 To 4
'''        Printer.Print
'''    Next i
'''    Printer.Print Space(3) & " " & left((vstrDesEfecSol) & Space(80), 80) & Space(1) & vstrTotSol
'''    Printer.Print
'''    Printer.Print Space(3) & " " & left((vstrDesEfecDol) & Space(80), 80) & Space(1) & vstrTotDol
'''    Printer.Print
'''    Printer.Print
'''    'Printer.FontName = "Draft 15cpi"
'''    Printer.Font.Size = 17
'''    Printer.Print Space(3) & " " & Trim(vstrNumPrecinto)
'''    Printer.Font.Size = 11
'''    Printer.EndDoc
'''
'''    Exit Function
'''Error:
'''    Err.Raise Err.Number, "clsImpresiones", Err.Description
'''
'''End Function



'''
'''Public Function Impresion_Hermes(ByVal vstrLugar As String, _
'''                                 ByVal vstrFecha As String, _
'''                                 ByVal vstrEmpresa As String, _
'''                                 ByVal vstrDesEfectivo As String, _
'''                                 ByVal vstrBanco As String, _
'''                                 ByVal vstrCtaCte As String, _
'''                                 ByVal vstrMonto As String, _
'''                                 ByVal vstrNumPrecinto As String, _
'''                                 Optional vstrDirEntrega As String, _
'''                                 Optional vstrEntrega As String, _
'''                                 Optional vstrEnvase As String)
'''
'''    'Set objPrint = New clsImpresora
'''    'Dim intFormato As Integer
'''    'Dim lngAncho As Double, lngLargo As Double
'''
'''    'lngAncho = 7.9 * 25.4
'''    'lngLargo = 5.9 * 25.4
'''    'intFormato = objPrint.CrearFormato("PH", lngAncho, lngLargo)
'''    'Set objPrint = Nothing
'''    'Printer.PaperSize = intFormato
'''
'''    'Printer.Height = 7.9 * 25.4
'''    'Printer.Width = 5.9 * 25.4
'''    i = 1
'''    For i = 1 To 4
'''        Printer.Print
'''    Next i
'''    Printer.FontName = "Draft 17cpi"
'''    Printer.FontSize = 11
'''    Printer.Print " " & left((vstrLugar) & Space(30), 30) & Space(2) & left((Format(vstrFecha, "dd/mm/yyyy")) & Space(12), 12)
'''    Printer.Print
'''    Printer.Print " " & left((vstrEmpresa) & Space(60), 60) & Space(30) & Trim(vstrEntrega) & Space(10) & Trim(vstrEnvase)
'''    Printer.Print
'''    Printer.Print " " & left((vstrDesEfectivo) & Space(60), 60) & Space(26) & right(Space(15) & vstrMonto, 15)
'''    Printer.Print
'''    Printer.Font.Size = 18
'''    Printer.Font.Bold = True
'''    If Mid(vstrMonto, 1, 1) = "S" Then
'''       Printer.Print " " & "      " & "X" & "          " & "X"
'''     Else
'''       Printer.Print " " & "      " & "X" & "          "
'''    End If
'''    Printer.Font.Size = 11
'''    Printer.Font.Bold = False
'''    Printer.Print
'''    Printer.Print Space(23) & " " & left((vstrBanco) & Space(60), 60)
'''    Printer.Print
'''    Printer.Print Space(12) & " " & left((vstrCtaCte) & Space(70), 70)
'''    Printer.Print
'''    Printer.Print
'''    Printer.Print
'''    Printer.Print
'''    Printer.Print Space(12) & " " & Trim(vstrDirEntrega)
'''    For i = 1 To 5
'''        Printer.Print
'''    Next i
'''    Printer.FontSize = 18
'''    Printer.Print Space(10) & " " & Trim(vstrNumPrecinto)
'''    Printer.FontSize = 11
'''    Printer.EndDoc
'''
'''    Exit Function
'''Error:
'''    Err.Raise Err.Number, "clsImpresiones", Err.Description
'''
'''End Function

Public Function NotaCred(ByVal vstrcia As String, _
                         ByVal vstrLocal As String, _
                         ByVal vstrDocumento As String)

Dim odynCabNC As oraDynaset
Dim odynDetNC As oraDynaset
Dim intRespuesta As Integer
Dim l%, k%
Dim strSubTot As String
Dim strIgv As String
Dim strTotal As String

   On Error GoTo ErrorImpresora

   strTotal = "": strIgv = "": strTotal = ""

   Set odynCabNC = objDocumento.fnCabNotaCred(vstrcia, vstrLocal, vstrDocumento)
   Set odynDetNC = objDocumento.fnDetNotaCred(vstrDocumento)

   Printer.Width = 220 * 56.7
   Printer.Height = 170 * 56.7
   Printer.Print
   Printer.FontName = "Draft 17cpi"
   Printer.Print "."
   Printer.FontName = "Draft 12cpi"
   Printer.Print
   Printer.Print
   Printer.Print
   Printer.Print Space$(59) & "N.I.T. " & odynCabNC("NUM_RUC_EMPRESA").Value
   Printer.Print
   Printer.Print
   Printer.Print Space$(62) & odynCabNC("NUM_DOCUMENTO").Value
   Printer.Print
   Printer.Print
   Printer.Print Space$(11) & Format(day(odynCabNC("FCH_REGISTRA").Value), "0#") & Space$(6) & left(odynCabNC("MES").Value & Space(20), 20) & Space$(6) & Format(year(odynCabNC("FCH_REGISTRA").Value), "####")
   Printer.Print
   Printer.Print
   Printer.Print Space$(2) & left(odynCabNC("NOMBRE_CLI").Value & Space(58), 58) & Space$(10) & Format(odynCabNC("FCH_FACTURA").Value, "dd/MM/yyyy")
   Printer.Print
   Printer.Print Space$(2) & left(odynCabNC("RUC").Value & Space(62), 62) & Space$(6) & odynCabNC("NUM_DOCUMENTO_REFER").Value
   Printer.Print
   Printer.Print Space$(2) & odynCabNC("DES_AUX_CLI_DIRECC").Value
   Printer.Print
   Printer.Print
   Printer.Print Space$(7) & odynCabNC("DES_SUBCONCEPTO").Value
   Printer.Print
   Printer.Print
   Printer.Print

   l = 0

   Dim Cadena$
   odynDetNC.MoveFirst
   While Not odynDetNC.EOF
      If (odynDetNC("FRA").Value > 0) And (odynDetNC("UND").Value > 0) Then
          Cadena = right(Space(5) & odynDetNC("UND").Value, 5) & " " & right(Space(5) & "F" & odynDetNC("FRA").Value, 5) & Space(1)
                        'Left(odynDetNC("DES_PRODUCTO").Value & Space(40), 40) & _
                        'Right(Space(5) & Format(odynDetNC("PRC_UNIT_VTA").Value, "###,###0.00"), 5)
      ElseIf (odynDetNC("UND").Value > 0) And (odynDetNC("FRA").Value <= 0) Then
          Cadena = right(Space(5) & odynDetNC("UND").Value, 5) & Space(1)
                        'Left(odynDetNC("DES_PRODUCTO").Value & Space(40), 40) & _
                        'Right(Space(5) & Format(odynDetNC("PRC_UNIT_VTA").Value, "###,###0.00"), 5)
      ElseIf (odynDetNC("UND").Value <= 0) And (odynDetNC("FRA").Value > 0) Then
          Cadena = right(Space(5) & "F" & odynDetNC("FRA").Value, 5) & Space(1)
                     '   Left(odynDetNC("DES_PRODUCTO").Value & Space(40), 40) & _
                     '   Right(Space(5) & Format(odynDetNC("PRC_UNIT_VTA").Value, "###,###0.00"), 5)
      Else

      End If

      Cadena = IIf(Cadena = "", right(Space(5), 5), Cadena) & Space(1) & _
               left(odynDetNC("DES_PRODUCTO").Value & Space(40), 40) & _
               right(Space(5) & Format(odynDetNC("PRC_UNIT_VTA").Value, "###,###0.00"), 5)

      Printer.Print Cadena

      odynDetNC.MoveNext
      l = l + 1
   Wend
   strSubTot = odynCabNC("MTO_BASE_IMP").Value
   strIgv = odynCabNC("MTO_IMPUESTO").Value
   strTotal = odynCabNC("MTO_TOTAL").Value

   For k = 0 To (11 - l) '''(11 - grdDetalle.Rows)
       Printer.Print
   Next k

   Printer.Print "SON : " & UCase(objLetra.NumAletra(strTotal)) & " " & "NUEVOS SOLES"
   Printer.Print
   Printer.Print Space(69) & "S/. " & Format(Val(strSubTot), "###,###0.00")
   Printer.Print Space(69) & "S/. " & Val(strIgv)
   Printer.Print
   Printer.Print Space(69) & "S/. " & Format(Val(strTotal), "###,###0.00")
   Printer.EndDoc

   On Error GoTo 0

   Exit Function

ErrorImpresora:

   intRespuesta = MsgBox("Existe un problema con la impresora" + Chr(13) + "Desea esperar a ser resuelto?", vbYesNo)
   If intRespuesta = vbYes Then
      Resume
   Else
      Exit Function
   End If

End Function

Public Function Imprime_Liquidacion(ByVal vstrcia As String, _
                                    ByVal vstrCodLocal As String, _
                                    ByVal vstrCodMaquina As String, _
                                    ByVal vstrCodLiquidacion As String)
    Dim odynCabLiq As oraDynaset
    Dim odynDetLiq As oraDynaset
    Dim odynCabRemesa As oraDynaset
    Dim odynDif As oraDynaset
    Dim intRespuesta%
    Dim strDifCaja As String
    Dim strCodFormaPago As String
    Dim strDesFormaPago As String
    Dim dblTotFormaPago As Double
    Dim dblTotRetiroEfe As Double
    Dim dblTotalEfectivo As Double
    Dim dblTotalRedondeo As Double
    Dim dblTotalDiferencia As Double

    Dim i%

    On Error GoTo ErrorImpresora

    Set odynCabLiq = objLiquidacion.ImpLiquidacionCabecera(vstrcia, vstrCodLocal, _
                                                           vstrCodMaquina, vstrCodLiquidacion)

    Set odynDetLiq = objLiquidacion.ImpLiquidacionDetalle(vstrCodLiquidacion)

    Set odynCabRemesa = objLiquidacion.ImpLiquidacionRemesa(vstrcia, vstrCodLocal, _
                                                            vstrCodMaquina, vstrCodLiquidacion)

    strDifCaja = objLiquidacion.DifCajaImp(vstrCodLiquidacion)

    Printer.Width = 220 * 56.7
    Printer.Height = 170 * 56.7
    Printer.Print
    Printer.FontName = "Draft 17cpi"
    Printer.Print Space(1) & left(odynCabLiq("DESCRIP").Value & Space(30), 30) & Space(30) & "BTL" & "-" & vstrCodLocal
    Printer.FontName = "Draft 12cpi"
    Printer.Print Space(1) & left(odynCabLiq("DIRECCION").Value & Space(30), 30)
    Printer.Print
    Printer.Print
    Printer.Print Space(35) & "LIQUIDACION DE CAJA"
    Printer.Print
    Printer.Print Space(1) & "FCH.INI.:" & Space(1) & left(odynCabLiq("FCH_INICIO").Value & Space(12), 12) & Space(1) & "FCH.LIQ.:" & Space(1) & left(odynCabLiq("FCH_LIQUIDA").Value & Space(12), 12)
    Printer.Print Space(1) & "QUIMICO :" & Space(1) & left(odynCabLiq("QUIMICO").Value & Space(30), 30) & Space(20) '& odynCabLiq("FCH_LIQUIDACION").Value
    Printer.Print Space(1) & "CAJERO  :" & Space(1) & left(odynCabLiq("TECNICO").Value & Space(30), 30) & Space(20) & odynCabLiq("COD_MAQUINA").Value
    Printer.Print
    Printer.Print

    '----------------'
    'Detalle de Venta'
    '----------------'
    Dim dblTotSist As Double
    dblTotSist = 0

    Printer.Print Space(1) & "1.- DOCUMENTOS EMITIDOS (SISTEMA)"
    Printer.Print Space(1) & "    ---------------------------- "
    Printer.Print
    Printer.Print Space(1) & left("FPago N1" & Space(12), 12) & Space(2), _
                             left("FPago N2" & Space(12), 12) & Space(8), _
                             left("Importe" & Space(8), 8) & Space(2), _
                             left("Ret Efect" & Space(7), 7)

    Printer.Print String(80, "-")
    odynDetLiq.MoveFirst
    dblTotRetiroEfe = 0
    dblTotalEfectivo = 0
    dblTotalRedondeo = 0
    While Not odynDetLiq.EOF
        strCodFormaPago = odynDetLiq("COD_FORMA_PAGO").Value
        strDesFormaPago = odynDetLiq("DES_FORMA_PAGO").Value
        dblTotFormaPago = 0
        While Not odynDetLiq.EOF And strCodFormaPago = odynDetLiq("COD_FORMA_PAGO").Value
            Printer.Print left(Trim(odynDetLiq("DES_FORMA_PAGO").Value) & Space(12), 12) & Space(2), _
                      left(odynDetLiq("DES_HIJO").Value & Space(12), 12) & Space(8), _
                      right(Space(8) & Format(odynDetLiq("IMP_LIQ_MONED_NAC").Value, "####0.00"), 8) & Space(2), _
                      left(odynDetLiq("FLG_RETIRO_EFEC").Value & Space(7), 7)
                      If odynDetLiq("FLG_RETIRO_EFEC").Value <> "1" Then
                        dblTotSist = dblTotSist + odynDetLiq("IMP_LIQ_MONED_NAC").Value
                        dblTotFormaPago = dblTotFormaPago + odynDetLiq("IMP_LIQ_MONED_NAC").Value
                        If strCodFormaPago = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_FORMA_PAGO_EFE") Then
                            dblTotalEfectivo = dblTotalEfectivo + odynDetLiq("IMP_LIQ_MONED_NAC").Value
                        End If
                      Else
                        dblTotRetiroEfe = dblTotRetiroEfe + odynDetLiq("IMP_LIQ_MONED_NAC").Value
                      End If
                      
                      If strCodFormaPago = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_FORMA_PAGO_RED") Then
                         dblTotalRedondeo = dblTotalRedondeo + odynDetLiq("IMP_LIQ_MONED_NAC").Value
                      End If
            odynDetLiq.MoveNext
        Wend
        
        
        
        Printer.Print left(Trim("") & Space(12), 12) & Space(2), _
                      left("" & Space(12), 12) & Space(8), _
                      right(Space(8) & String(8, "-"), 8)
        Printer.Print left(Trim("") & Space(12), 12) & Space(2), _
                      left("Total " & strDesFormaPago & Space(12), 12) & Space(8), _
                      right(Space(8) & Format(dblTotFormaPago, "####0.00"), 8)
                      
        Printer.Print
    Wend
    Printer.Print String(80, "-")
    
    Printer.Print left(Trim("") & Space(12), 12) & Space(2), _
                      left("Efectivo" & Space(12), 12) & Space(8), _
                      right(Space(8) & Format(dblTotalEfectivo, "####0.00"), 8)
    Printer.Print left(Trim("") & Space(12), 12) & Space(2), _
                      left("Retiro" & Space(12), 12) & Space(8), _
                      right(Space(8) & Format(dblTotRetiroEfe, "####0.00"), 8)
    Printer.Print String(80, "-")
    
    Printer.Print left(Trim("") & Space(12), 12) & Space(2), _
                      left("Total Efectivo" & Space(12), 12) & Space(8), _
                      right(Space(8) & Format(dblTotalEfectivo - dblTotRetiroEfe, "####0.00"), 8)
    
    Printer.Print String(80, "=")
    Printer.Print left(Trim("") & Space(12), 12) & Space(2), _
                      left("Total Cobrado" & Space(12), 12) & Space(8), _
                      right(Space(8) & Format(dblTotSist, "####0.00"), 8)
    
    'Printer.Print Space(1) & "Total Cobrado " & Space(20) & Format(dblTotSist, "#,###,##0.00")

    For i = 0 To 4
        Printer.Print
    Next i

    '-----------------------'
    'Detalle de lo Declarado'
    '-----------------------'
    Dim dblTotDecl As Double
    dblTotDecl = 0

    Printer.Print Space(1) & "2.- LIQUIDACION DE COBRANZA (DECLARADO)"
    Printer.Print Space(1) & "    ---------------------------------- "
    Printer.Print
    Printer.Print Space(1) & left("Moneda" & Space(10), 10) & Space(8), _
                             left("TC" & Space(9), 9) & Space(2), _
                             left("SubTotal" & Space(8), 8)

    Printer.Print String(50, "-")
    odynCabRemesa.MoveFirst
    While Not odynCabRemesa.EOF
        Printer.Print left(Trim(odynCabRemesa("CONCEPTO").Value) & Space(10), 10) & Space(8), _
                      right(Space(9) & Format(odynCabRemesa("IMP_TIPO_CAMBIO").Value, "####0.00"), 9) & Space(2), _
                      right(Space(8) & Format(odynCabRemesa("SUBTOTAL").Value, "####0.00"), 8)

                      dblTotDecl = dblTotDecl + odynCabRemesa("SUBTOTAL").Value

        odynCabRemesa.MoveNext
    Wend
    Printer.Print String(50, "-")
    Printer.Print
    odynCabRemesa.MoveFirst
    
    dblTotalDiferencia = Val(strDifCaja) + dblTotalRedondeo
    
    Printer.Print left(Trim("") & Space(10), 10) & Space(8), _
                      right(Space(9) & Trim("Total Declarado"), 9) & Space(2), _
                      right(Space(8) & Format(dblTotDecl, "####0.00"), 8)
    
    ''Printer.Print Space(1) & "Total Declarado " & Space(20) & Format(dblTotDecl, "#,###,##0.00")
    Printer.Print
    Printer.Print String(50, "-")
    Printer.Print
    Printer.Print Space(1) & "Diferencia de Caja =>" & Space(20) & Format(dblTotalDiferencia, "####0.00")
    Printer.Print Space(1) & "Reposición de Caja =>" & Space(20) & Format(odynCabLiq("MTO_RETIRO_CAJA").Value, "####0.00")
    Printer.Font.Bold = True
    'Printer.FontName = "Draft 15cpi"
    Printer.Font.Size = 20
    Printer.Print Space(1) & "***************************************************************** "
    Printer.Print Space(1) & "CODIGO LIQUIDACION =>" & Space(20) & odynCabLiq("COD_LIQUIDACION").Value
    Printer.Print Space(1) & "*****************************************************************"
    Printer.Font.Bold = False
    Printer.Font.Size = 12
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print Space(10) & left("---------------------" & Space(30), 30) & Space(20) & left("---------------------" & Space(30), 30)
    Printer.Print Space(10) & left("Firma del Quimico" & Space(30), 30) & Space(20) & left("Firma del Tècnico" & Space(30), 30)
    Printer.EndDoc

    Exit Function
ErrorImpresora:
   intRespuesta = MsgBox("Existe un problema con la impresora" + Chr(13) + "Desea esperar a ser resuelto?", vbYesNo)
   If intRespuesta = vbYes Then
      Resume
   Else
      Exit Function
   End If
End Function

Public Function CrearFormato(ByVal vstrNombre As String, _
                             ByVal vlngWidth As Long, _
                             ByVal vlngHeigh As Long) As Integer
    Dim FormSize As SIZEL
    Dim PrinterHandle As Long
    Dim PrinterName As String
    Dim temp As String
    temp = vstrNombre
    PrinterName = Printer.Devicename 'Impresora Actual
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        With FormSize
        .cx = vlngWidth * 1000
        .cy = vlngHeigh * 1000
        End With
        If GetFormName(PrinterHandle, FormSize, temp) = 0 Then
            AddNewForm PrinterHandle, FormSize, vstrNombre
        End If
        CrearFormato = GetFormName(PrinterHandle, FormSize, "")
        ClosePrinter (PrinterHandle)
    End If
End Function

'***** @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ ********'
'***** REIMPRESION DE REMITOS PROSEGUR/HERMES ********'
'***** @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ ********'
'''Autor     : Juan Arturo Escate
'''Fecha     : 25/11/2009
'''Proposito : Se comento porque se implemento un nuevo metodo que no consultase de la pantalla sino lo grabado
'''Public Function Re_Impresion_Prosegur(ByVal vstrCodLocal As String, _
'''                                      ByVal vstrCodRemito As String)
'''
'''    Dim strCtdRem As String
'''
'''    strCtdRem = objRemito.Dev_CtdReg_Remeda(vstrCodLocal, vstrCodRemito)
'''    strBanco = objRemito.BancoxBtl(vstrCodLocal)
'''    strPortavalor = objRemito.Portavalor(vstrCodLocal)
'''    strCtaCteSol = objRemito.CtaCtexBtl(vstrCodLocal, "S")
'''
'''    Set odynReImp = objRemito.Re_Imp_Remito(vstrCodLocal, vstrCodRemito)
'''
'''    strDesEfecSol = objLetra.NumAletra(odynReImp("TOT_SOLES").Value)
'''    strDesEfecDol = objLetra.NumAletra(odynReImp("TOT_DOLARES").Value)
'''
'''    i = 1
'''    For i = 1 To 13
'''        Printer.Print
'''    Next i
'''    Printer.FontName = "Draft 17cpi"
'''    Printer.FontSize = 11
'''    Printer.Print Space(8) & "Cantidad Sobres =>" & " " & Trim(strCtdRem) & Space(30) & " " & Trim(strBanco)
'''    Printer.Print
'''    Printer.Print Space(42) & " " & Trim(odynReImp("COD_LOCAL").Value)
'''    Printer.Print
'''    Printer.Print Space(42) & " " & Trim(strPortavalor)
'''    Printer.Print
'''    Printer.Print Space(42) & " " & Trim(odynReImp("DIRECCION").Value) & Space(20) & Trim(odynReImp("DISTRITO").Value)
'''    Printer.Print
'''    For i = 1 To 4
'''        Printer.Print
'''    Next i
'''    Printer.Print Space(3) & " " & left((strDesEfecSol) & Space(80), 80) & Space(1) & Format(odynReImp("TOT_SOLES").Value, "#,###,##0.00")
'''    Printer.Print
'''    Printer.Print Space(3) & " " & left((strDesEfecDol) & Space(80), 80) & Space(1) & Format(odynReImp("TOT_DOLARES").Value, "#,###,##0.00")
'''    Printer.Print
'''    Printer.Print
'''    'Printer.Print
'''    Printer.Print Space(3) & " " & Trim(odynReImp("PRECINTO").Value)
'''    Printer.EndDoc
'''
'''    Exit Function
'''Error:
'''    Err.Raise Err.Number, "clsImpresiones", Err.Description
'''
'''End Function

'''Public Function Re_Impresion_Hermes(ByVal vstrCodLocal As String, _
'''                                    ByVal vstrCodRemito As String, _
'''                                    ByVal vstrMoneda As String)
'''
'''    If vstrMoneda = "S" Then
'''        Set odynReImp = objRemito.Re_Imp_Remito_Hermes(vstrCodLocal, vstrCodRemito, vstrMoneda)
'''        strDesEfecSol = objLetra.NumAletra(odynReImp("TOT_SOLES").Value)
'''        strCtaCteSol = objRemito.CtaCtexBtl(vstrCodLocal, "S")
'''      Else
'''        Set odynReImp = objRemito.Re_Imp_Remito_Hermes(vstrCodLocal, vstrCodRemito, vstrMoneda)
'''        strDesEfecDol = objLetra.NumAletra(odynReImp("TOT_DOLARES").Value)
'''        strCtaCteDol = objRemito.CtaCtexBtl(vstrCodLocal, "D")
'''    End If
'''
'''    strBanco = objRemito.BancoxBtl(objUsuario.CodigoLocal)
'''
'''    i = 1
'''    For i = 1 To 4
'''        Printer.Print
'''    Next i
'''    Printer.FontName = "Draft 17cpi"
'''    Printer.FontSize = 11
'''    Printer.Print " " & left((odynReImp("DISTRITO").Value & " " & odynReImp("COD_LOCAL").Value) & Space(30), 30) & Space(2) & left((Format(odynReImp("FCH_REGISTRA").Value, "dd/mm/yyyy")) & Space(12), 12)
'''    Printer.Print
'''    Printer.Print " " & left((odynReImp("DIRECCION").Value) & Space(60), 60) & Space(30) & Trim("") & Space(10) & Trim("01")
'''    Printer.Print
'''    If Trim(vstrMoneda) = "S" Then
'''        Printer.Print " " & left((strDesEfecSol) & Space(60), 60) & Space(26) & right(Space(15) & Format(odynReImp("TOT_SOLES").Value, "#,###,##0.00"), 15)
'''    Else
'''        Printer.Print " " & left((strDesEfecDol) & Space(60), 60) & Space(26) & right(Space(15) & Format(odynReImp("TOT_DOLARES").Value, "#,###,##0.00"), 15)
'''    End If
'''    Printer.Print
'''    Printer.Font.Size = 18
'''    Printer.Font.Bold = True
'''    If Trim(vstrMoneda) = "S" Then
'''       Printer.Print " " & "      " & "X" & "          " & "X"
'''     Else
'''       Printer.Print " " & "      " & "X" & "          "
'''    End If
'''    Printer.Font.Size = 11
'''    Printer.Font.Bold = False
'''    Printer.Print
'''    Printer.Print Space(23) & " " & left((Trim(strBanco)) & Space(60), 60)
'''    Printer.Print
'''    If Trim(vstrMoneda) = "S" Then
'''        Printer.Print Space(12) & " " & left((Trim(strCtaCteSol)) & Space(70), 70)
'''    Else
'''        Printer.Print Space(12) & " " & left((Trim(strCtaCteDol)) & Space(70), 70)
'''    End If
'''    Printer.Print
'''    Printer.Print
'''    Printer.Print
'''    Printer.Print
'''    Printer.Print Space(12) & " " & Trim("HERMES PROCESAMIENTO")
'''    For i = 1 To 5
'''        Printer.Print
'''    Next i
'''    Printer.FontSize = 18
'''    Printer.Print Space(10) & " " & Trim(odynReImp("PRECINTO").Value)
'''    Printer.FontSize = 11
'''    Printer.EndDoc
'''
'''    Exit Function
'''Error:
'''    Err.Raise Err.Number, "clsImpresiones", Err.Description
'''
'''End Function

Public Function Imprime_Detalle_Liquidacion(ByVal vstrcia As String, _
                                            ByVal vstrCodLocal As String, _
                                            ByVal vstrCodMaquina As String, _
                                            ByVal vstrCodLiquidacion As String)
    Dim odynCabLiq As oraDynaset
    Dim odynDocEmitidos As oraDynaset
    Dim odynConvenio As oraDynaset
    Dim odynTarjetas As oraDynaset
    Dim odynDocAnulados As oraDynaset
    
    Dim intRespuesta%
    
    Dim lintCnt As Integer
    Dim ldblLuz As Double
    Dim ldblAgua As Double
    Dim ldblTelefoniaFija As Double
    Dim ldblTelefoniaMovil As Double
    Dim ldblTelefoniaTotal As Double
    Dim ldblScotiabankCargo As Double
    Dim ldblScotiabankAbono As Double
    Dim ldblScotiabankTotal As Double
    Dim ldblSoat As Double
    Dim ldblVisaCashCargo As Double
    Dim ldblVisaCashAbono As Double
    Dim ldblVisaCashTotal As Double
    Dim ldblDonaciones As Double
    Dim ldblCobranzaSoles As Double
    Dim ldblCobranzaDolares As Double
    Dim ldblCobranzaTotal As Double
    Dim ldblCupones As Double
    Dim ldblCheques As Double
    Dim lstrConcepto As String
    Dim ldblImporte As String
    Dim ldblTotDocEmitidos As Double
    Dim ldblTotConvenios As Double
    Dim ldblTotServicios As Double
    Dim ldblTotCobranza As Double
    Dim ldblTotTarjetas As Double
    Dim ldblTotCobrado1 As Double
    Dim ldblTotCobrado2 As Double
    Dim ldblTotDiferencias As Double

    Dim i%
    Dim p As Printer

'    On Error GoTo ErrorImpresora

    Set odynCabLiq = objLiquidacion.ImpLiquidacionCabecera(vstrcia, vstrCodLocal, vstrCodMaquina, vstrCodLiquidacion)
    Set odynDocEmitidos = objLiquidacion.ListaDocsEmitidos(vstrcia, vstrCodLiquidacion)
    Set odynConvenio = objLiquidacion.CalcConvenio(vstrcia, vstrCodLocal, vstrCodMaquina, vstrCodLiquidacion)
    Set odynTarjetas = objLiquidacion.ListaTarjetas(vstrcia, vstrCodLocal, vstrCodLiquidacion)
    Set odynDocAnulados = objLiquidacion.ListaDocsEmitidos(vstrcia, vstrCodLiquidacion, "ANU")
    
    ldblLuz = objLiquidacion.CalcServicio(vstrcia, vstrCodLocal, vstrCodLiquidacion, "004")
    ldblAgua = objLiquidacion.CalcServicio(vstrcia, vstrCodLocal, vstrCodLiquidacion, "003")
    ldblTelefoniaFija = objLiquidacion.CalcServicio(vstrcia, vstrCodLocal, vstrCodLiquidacion, "001")
    ldblTelefoniaMovil = objLiquidacion.CalcServicio(vstrcia, vstrCodLocal, vstrCodLiquidacion, "002")
    ldblTelefoniaTotal = ldblTelefoniaFija + ldblTelefoniaMovil
    ldblScotiabankCargo = objLiquidacion.CalcCajero(vstrcia, vstrCodLocal, vstrCodLiquidacion, "C")
    ldblScotiabankAbono = objLiquidacion.CalcCajero(vstrcia, vstrCodLocal, vstrCodLiquidacion, "A")
    ldblScotiabankTotal = ldblScotiabankCargo + ldblScotiabankAbono
    ldblSoat = objLiquidacion.CalcSoat(vstrcia, vstrCodLocal, vstrCodLiquidacion)
    ldblVisaCashCargo = objLiquidacion.CalcVisaCash(vstrcia, vstrCodLocal, vstrCodLiquidacion, "C")
    ldblVisaCashAbono = objLiquidacion.CalcVisaCash(vstrcia, vstrCodLocal, vstrCodLiquidacion, "A")
    ldblVisaCashTotal = ldblVisaCashCargo + ldblVisaCashAbono
    ldblDonaciones = objLiquidacion.CalcDonacion(vstrcia, vstrCodLocal, vstrCodLiquidacion)
    ldblCobranzaSoles = objLiquidacion.CalcEfect(vstrcia, vstrCodLocal, vstrCodLiquidacion, "001")
    ldblCobranzaDolares = objLiquidacion.CalcEfect(vstrcia, vstrCodLocal, vstrCodLiquidacion, "002")
    ldblCobranzaTotal = ldblCobranzaSoles + ldblCobranzaDolares
    ldblCupones = objLiquidacion.CalcCupones(vstrcia, vstrCodLocal, vstrCodLiquidacion)
    ldblCheques = objLiquidacion.CalcCheques(vstrcia, vstrCodLocal, vstrCodLiquidacion)
    
'    Printer.Width = 220 * 56.7
'    Printer.Height = 170 * 56.7
    
    For Each p In Printers
       If UCase(p.PORT) = "LPT1:" Then
          Set Printer = p
          Exit For
       End If
    Next p
    
    If Ver.dwPlatformId = VER_PLATFORM_WIN32_NT And Ver.dwMajorVersion = VER_PRINCIPAL_WINXP Then
        Dim objImpresion As New clsImpresion
     
        If Not objImpresion.cargaParametros("003") Then
            Exit Function
        End If
    Else
        Printer.Width = 1440 * 8.5
        Printer.Height = 1440 * 8
    End If
    
    Printer.Print
    Printer.FontName = "Draft 17cpi"
    Printer.Print Space(1) & left(odynCabLiq("DESCRIP").Value & Space(30), 30) & Space(30) & "BTL" & "-" & vstrCodLocal
    Printer.FontName = "Draft 12cpi"
    Printer.Print Space(1) & left(odynCabLiq("DIRECCION").Value & Space(30), 30)
    Printer.Print
    Printer.Print
    Printer.Print Space(35) & "LIQUIDACION DE CAJA"
    Printer.Print
    Printer.Print Space(1) & "FCH.INI.:" & Space(1) & left(odynCabLiq("FCH_INICIO").Value & Space(12), 12) & Space(1) & "FCH.LIQ.:" & Space(1) & left(odynCabLiq("FCH_LIQUIDA").Value & Space(12), 12)
    Printer.Print Space(1) & "QUIMICO :" & Space(1) & left(odynCabLiq("QUIMICO").Value & Space(30), 30) & Space(20) '& odynCabLiq("FCH_LIQUIDACION").Value
    Printer.Print Space(1) & "CAJERO  :" & Space(1) & left(odynCabLiq("TECNICO").Value & Space(30), 30) & Space(20) & odynCabLiq("COD_MAQUINA").Value
    Printer.Print
    Printer.Print
    
    Printer.Print "DETALLE DE LIQUIDACION " & vstrCodLiquidacion
    Printer.Print
    Printer.Print "Documentos Emitidos"
    Printer.Print String(69, "=")
    Printer.Print "Tipo                 Inicial         Final           Total          "
    Printer.Print String(69, "=")
    
    lintCnt = 0
    ldblTotDocEmitidos = 0
    
    odynDocEmitidos.MoveFirst
    While Not odynDocEmitidos.EOF
        Printer.Print left(odynDocEmitidos("DES_TIPODOC").Value & Space(20), 20) & " " & _
                      left(odynDocEmitidos("DOCINI").Value & Space(15), 15) & " " & _
                      left(odynDocEmitidos("DOCFIN").Value & Space(15), 15) & " " & _
                      right(Space(15) & Trim(Format(odynDocEmitidos("DOCTOT").Value, "#,###,##0.00")), 15)

        lintCnt = lintCnt + odynDocEmitidos("DOCCNT").Value
        ldblTotDocEmitidos = ldblTotDocEmitidos + odynDocEmitidos("DOCTOT").Value
        
        odynDocEmitidos.MoveNext
    Wend
    If lintCnt = 0 Then
        Printer.Print "NINGUNO"
    Else
        Printer.Print String(69, "-")
        Printer.Print "Total Documentos Emitidos " & right(Space(15) & Trim(Format(ldblTotDocEmitidos, "#,###,##0.00")), 15)
    End If
     
    Printer.Print
    Printer.Print "Servicios"
    Printer.Print String(69, "=")
    
    ldblTotServicios = 0
    For i = 1 To 6
        If i = 1 Then
            lstrConcepto = "Total cobrado de Luz"
            ldblImporte = ldblLuz
        End If
        If i = 2 Then
            lstrConcepto = "Total cobrado de Agua"
            ldblImporte = ldblAgua
        End If
        If i = 3 Then
            lstrConcepto = "Total cobrado de Telefónica"
            ldblImporte = ldblTelefoniaTotal
        End If
        If i = 4 Then
            lstrConcepto = "Total Scotiabank"
            ldblImporte = ldblScotiabankTotal
        End If
        If i = 5 Then
            lstrConcepto = "Total Soat"
            ldblImporte = ldblSoat
        End If
        If i = 6 Then
            lstrConcepto = "Total VisaCash"
            ldblImporte = ldblVisaCashTotal
        End If
        
        Printer.Print left(lstrConcepto & Space(30), 30) & " " & right(Space(15) & Trim(Format(ldblImporte, "#,###,##0.00")), 15)
        
        ldblTotServicios = ldblTotServicios + ldblImporte
    Next i
    
    Printer.Print String(69, "-")
    Printer.Print "Total Servicios " & right(Space(15) & Trim(Format(ldblTotServicios, "#,###,##0.00")), 15)
    
    Printer.Print
    Printer.Print "Convenios"
    Printer.Print String(69, "=")
    
    lintCnt = 0
    ldblTotConvenios = 0
    
    odynConvenio.MoveFirst
    While Not odynConvenio.EOF
        Printer.Print left(odynConvenio("DES_CONVENIO").Value & Space(20), 20) & " " & _
                      right(Space(15) & Trim(Format(odynConvenio("MTO_TOTAL").Value, "#,###,##0.00")), 15)
        
        lintCnt = lintCnt + 1
        ldblTotConvenios = ldblTotConvenios + odynConvenio("MTO_TOTAL").Value
        
        odynConvenio.MoveNext
    Wend
    
    If lintCnt = 0 Then
        Printer.Print "NINGUNO"
    Else
        Printer.Print String(69, "-")
        Printer.Print "Total Convenios " & right(Space(15) & Trim(Format(ldblTotConvenios, "#,###,##0.00")), 15)
    End If
    
    Printer.Print
    Printer.Print "Donaciones"
    Printer.Print String(69, "=")
    Printer.Print "Total Donaciones " & right(Space(15) & Trim(Format(ldblDonaciones, "#,###,##0.00")), 15)

    Printer.Print
    Printer.Print "Cobranza"
    Printer.Print String(69, "=")
    Printer.Print left("Soles" & Space(20), 20) & right(Space(15) & Trim(Format(ldblCobranzaSoles, "#,###,##0.00")), 15)
    Printer.Print left("Dolares" & Space(20), 20) & right(Space(15) & Trim(Format(ldblCobranzaDolares, "#,###,##0.00")), 15)
    Printer.Print left("Cupón" & Space(20), 20) & right(Space(15) & Trim(Format(ldblCupones, "#,###,##0.00")), 15)
    ldblTotCobranza = ldblCobranzaSoles + ldblCobranzaDolares + ldblCupones
    Printer.Print String(69, "-")
    Printer.Print "Total Cobranza " & right(Space(15) & Trim(Format(ldblTotCobranza, "#,###,##0.00")), 15)

    Printer.Print
    Printer.Print "Tarjetas de Crédito"
    Printer.Print String(69, "=")

    lintCnt = 0
    ldblTotTarjetas = 0

    odynTarjetas.MoveFirst
    While Not odynTarjetas.EOF
        Printer.Print left(odynTarjetas("DES").Value & Space(20), 20) & " " & _
                      right(Space(15) & Trim(Format(odynTarjetas("MONTO").Value, "#,###,##0.00")), 15)

        lintCnt = lintCnt + 1
        ldblTotTarjetas = ldblTotTarjetas + odynTarjetas("MONTO").Value

        odynTarjetas.MoveNext
    Wend
    
    If lintCnt = 0 Then
        Printer.Print "NINGUNO"
    Else
        Printer.Print String(69, "-")
        Printer.Print "Total Tarjetas " & right(Space(15) & Trim(Format(ldblTotTarjetas, "#,###,##0.00")), 15)
    End If
    
    Printer.Print
    Printer.Print "Cheques"
    Printer.Print String(69, "=")
    Printer.Print "Total Cheques " & right(Space(15) & Trim(Format(ldblCheques, "#,###,##0.00")), 15)
    
    Printer.Print
    Printer.Print "Documentos Anulados"
    Printer.Print String(69, "=")
    Printer.Print "Tipo                Cantidad"
    Printer.Print String(69, "=")
    
    lintCnt = 0
    
    odynDocAnulados.MoveFirst
    While Not odynDocAnulados.EOF
        Printer.Print left(odynDocAnulados("DES_TIPODOC").Value & Space(20), 20) & " " & _
                      right(Space(15) & Trim(Format(odynDocAnulados("DOCCNT").Value, "#,###,##0")), 15)

        lintCnt = lintCnt + 1

        odynDocAnulados.MoveNext
    Wend
    If lintCnt = 0 Then
        Printer.Print "NINGUNO"
    End If
    
    ldblTotCobrado1 = ldblTotDocEmitidos + ldblTotServicios + ldblTotConvenios + ldblDonaciones
    ldblTotCobrado2 = ldblTotCobranza + ldblTotTarjetas + ldblCheques
    ldblTotDiferencias = ldblTotCobrado1 - ldblTotCobrado2
    
    Printer.Print
'    Printer.Print String(69, "=")
'    Printer.Print "TOTAL DIFERENCIAS " & right(Space(15) & Trim(Format(ldblTotDiferencias, "#,###,##0.00")), 15)
'    Printer.Print String(69, "=")
    
    Printer.Print
    Printer.Font.Bold = False
    Printer.Font.Size = 12
    Printer.EndDoc

    Exit Function
ErrorImpresora:
   intRespuesta = MsgBox("Existe un problema con la impresora" + Chr(13) + "Desea esperar a ser resuelto?", vbYesNo)
   If intRespuesta = vbYes Then
      Resume
   Else
      Exit Function
   End If
End Function

Public Function Lista_DocsxProf(ByVal vstrcia As String, _
                                ByVal vstrCodLocal As String, _
                                ByVal vstrNumProforma As String) As oraDynaset
    On Error GoTo CtrlErr
    Set Lista_DocsxProf = gclsOracle.FN_Cursor("BTLPROD.PKG_PROFORMA_V3.FN_LISTA_DOCS_X_PROF", 0, vstrcia, vstrCodLocal, vstrNumProforma)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsProforma.Lista_DocsxProf", vbCritical, Err.Description
End Function

Public Function Imprime_Remesas_Soles(ByVal vstrcia As String, _
                                      ByVal vstrCodLocal As String, _
                                      ByVal vstrCodLiqudiacion As String, _
                                      ByVal vstrCodRemesa As String, _
                                      ByVal vstrImpSoles As String, _
                                      ByVal vstrFchDesposito As String, _
                                      ByVal vstrHoraDeposito As String, _
                                      ByVal vstrCodUsuarioOrigen As String, _
                                      ByVal vstrUsuarioOrigen As String, _
                                      ByVal vstrMaquina As String, _
                                      ByVal vstrCodUsuarioDestino As String, _
                                      ByVal vstrNombreDestino As String)

    Printer.FontName = "Draft 17cpi"
    Printer.FontSize = 11
    For i = 1 To 3
        Printer.Print
    Next i
    Printer.FontName = "Draft 14cpi"
    Printer.Print Space(33) & "Nº Remesa : " & Trim(vstrCodRemesa)
    Printer.Print Space(33) & String(39, "-")
    Printer.FontName = "Draft 17cpi"
    For i = 1 To 6
        Printer.Print
    Next i
    'Printer.FontName = "Draft 8cpi"    Draft 6 es el mas grande
    Printer.FontName = "Draft 6cpi"
    Printer.Print Space(9) & pfstr_Alineado("SOLES", 10, centro)
    Printer.Print
    Printer.Print Space(9) & pfstr_Alineado(vstrImpSoles, 10, centro)
    Printer.FontName = "Draft 17cpi"
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print Space(5) & "Local : " & Trim(vstrCodLocal)
    Printer.Print Space(5) & "Nº Liquidación : " & Trim(vstrCodLiqudiacion)
    Printer.Print
    Printer.Print Space(34) & ("QF") & Space(1) & vstrCodUsuarioDestino & Space(1) & vstrNombreDestino
    Printer.Print Space(14) & Trim(vstrFchDesposito) & Space(10) & String(60, "-")
    Printer.Print Space(34) & (vstrMaquina) & Space(1) & vstrCodUsuarioOrigen & Space(1) & vstrUsuarioOrigen
    Printer.Print Space(14) & Trim(vstrHoraDeposito)
    Printer.EndDoc
    Exit Function
Error:
    Err.Raise Err.Number, "clsImpresiones", Err.Description
End Function

Public Function Imprime_Remesas_Dolares(ByVal vstrcia As String, _
                                        ByVal vstrCodLocal As String, _
                                        ByVal vstrCodLiqudiacion As String, _
                                        ByVal vstrCodRemesa As String, _
                                        ByVal vstrImpDolares As String, _
                                        ByVal vstrFchDesposito As String, _
                                        ByVal vstrHoraDeposito As String, _
                                        ByVal vstrCodUsuarioOrigen As String, _
                                        ByVal vstrUsuarioOrigen As String, _
                                        ByVal vstrMaquina As String, _
                                        ByVal vstrCodUsuarioDestino As String, _
                                        ByVal vstrNombreDestino As String)

    Printer.FontName = "Draft 17cpi"
    Printer.FontSize = 11
    For i = 1 To 3
        Printer.Print
    Next i
    Printer.FontName = "Draft 14cpi"
    Printer.Print Space(33) & "Nº Remesa : " & Trim(vstrCodRemesa)
    Printer.Print Space(33) & String(39, "-")
    Printer.FontName = "Draft 17cpi"
    For i = 1 To 6
        Printer.Print
    Next i
    'Printer.FontName = "Draft 8cpi"    Draft 6 es el mas grande
    Printer.FontName = "Draft 6cpi"
    Printer.Print Space(9) & pfstr_Alineado("DOLARES", 10, centro)
    Printer.Print
    Printer.Print Space(9) & pfstr_Alineado(vstrImpDolares, 10, centro)
    Printer.FontName = "Draft 17cpi"
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print Space(5) & "Local : " & Trim(vstrCodLocal)
    Printer.Print Space(5) & "Nº Liquidación : " & Trim(vstrCodLiqudiacion)
    Printer.Print
    Printer.Print Space(34) & ("QF") & Space(1) & vstrCodUsuarioDestino & Space(1) & vstrNombreDestino
    Printer.Print Space(14) & Trim(vstrFchDesposito) & Space(10) & String(60, "-")
    Printer.Print Space(34) & (vstrMaquina) & Space(1) & vstrCodUsuarioOrigen & Space(1) & vstrUsuarioOrigen
    Printer.Print Space(14) & Trim(vstrHoraDeposito)
    Printer.EndDoc
    Exit Function
Error:
    Err.Raise Err.Number, "clsImpresiones", Err.Description
End Function

Public Function Imprime_Detalle_Remito(ByVal vstrcia As String, _
                                       ByVal vstrCodLocal As String, _
                                       ByVal vstrCodRemito As String)

    Dim odynDetalle As oraDynaset
    Dim i%
    Dim antUsuario, antCaja, antLiquidacion As String
    Dim nvoUsuario, nvoCaja, nvoLiquidacion As String
    Dim ldblTotSoles As Double
    Dim ldblTotDolares As Double
    
'    Dim p As Printer
'
'    For Each p In Printers
'       If UCase(p.Port) = "LPT1:" Then
'          Set Printer = p
'          Exit For
'       End If
'    Next p

    On Error GoTo CtrlErr
    
    Set odynDetalle = gclsOracle.FN_Cursor("BTLPROD.PKG_REMITO.FN_IMP_DETALLE_REMITO", 0, _
                                                                                       vstrcia, _
                                                                                       vstrCodLocal, _
                                                                                       vstrCodRemito)
    Printer.FontName = "Draft 16cpi"
    Printer.Font.Size = 11
'    Alinea la Hoja en forma Vertical
'    Printer.Orientation = 2

    If Len(odynDetalle("QF").Value) > 40 Then

        Printer.Print "QF:" & Space(1) & pfstr_Alineado(Mid(odynDetalle("QF").Value, 1, 8) & _
                                         pfstr_Espacio(Mid(odynDetalle("QF").Value, 9), 1), 40, Izquierda) & Space(4) & _
                      "Fecha:" & Space(1) & odynDetalle("FCH_REMITO").Value & Space(4) & _
                      "Remito:" & Space(1) & odynDetalle("REMITO").Value & Space(4) & _
                      "Fondo:" & Space(1) & Format(odynDetalle("FONDO_SENCILLO").Value, "###,###0.00")
        Printer.Print Space(5) & pfstr_Alineado(pfstr_Espacio(Mid(odynDetalle("QF").Value, 9), 2), 39, Izquierda)
    Else

        Printer.Print "QF:" & Space(1) & odynDetalle("QF").Value & Space(4) & _
                      "Fecha:" & Space(1) & odynDetalle("FCH_REMITO").Value & Space(4) & _
                      "Remito:" & Space(1) & odynDetalle("REMITO").Value & Space(4) & _
                      "Fondo:" & Space(1) & Format(odynDetalle("FONDO_SENCILLO").Value, "###,###0.00")
    End If
    
    Printer.Print

    Printer.FontName = "Draft 17cpi"
    Printer.Font.Size = 11
    
    Printer.Print String(135, "-")
    Printer.Print "|" & pfstr_Alineado("Cajero", 40, centro) & "|" & _
                        pfstr_Alineado("Caja", 10, centro) & "|" & _
                        pfstr_Alineado("Liquidacion", 17, centro) & "|" & _
                        pfstr_Alineado("Fecha Remesa", 21, centro) & "|" & _
                        pfstr_Alineado("Remesa", 17, centro) & "|" & _
                        pfstr_Alineado("Soles", 11, centro) & "|" & _
                        pfstr_Alineado("Dolares", 11, centro) & _
                  "|"
    Printer.Print String(135, "-")

        antUsuario = ""
        antCaja = ""
        antLiquidacion = ""
        nvoUsuario = ""
        nvoCaja = ""
        nvoLiquidacion = ""

    Do While Not odynDetalle.EOF
        
        nvoUsuario = Mid(odynDetalle("USUARIO").Value, 1, 8)
        nvoCaja = odynDetalle("COD_MAQUINA").Value
        nvoLiquidacion = odynDetalle("COD_LIQUIDACION").Value
        
        If antUsuario = nvoUsuario And antCaja = nvoCaja And antLiquidacion = nvoLiquidacion Then
                Printer.Print "|" & String(40, " ") & "|" & _
                                    String(10, " ") & "|" & _
                                    String(17, " ") & "|" & _
                                    pfstr_Alineado(odynDetalle("FCH_REMESA").Value, 21, centro) & "|" & _
                                    pfstr_Alineado(odynDetalle("COD_REMESA").Value, 17, centro) & "|" & _
                                    pfstr_Alineado(Format(odynDetalle("IMP_SOLES").Value, "###,###0.00"), 11, centro) & "|" & _
                                    pfstr_Alineado(Format(odynDetalle("IMP_DOLARES").Value, "###,###0.00"), 11, centro) & "|"
                Printer.Print "|" & String(40, " ") & "|" & _
                                    String(10, " ") & "|" & _
                                    String(17, " ") & "|" & _
                                    String(64, "-")
                ldblTotSoles = ldblTotSoles + odynDetalle("IMP_SOLES").Value
                ldblTotDolares = ldblTotDolares + odynDetalle("IMP_DOLARES").Value
        Else
            If antUsuario <> nvoUsuario And antCaja <> nvoCaja And antLiquidacion <> nvoLiquidacion Then
                Printer.Print String(135, "-")
            End If
            
            If Len(odynDetalle("USUARIO").Value) > 40 Then
                Printer.Print "|" & Space(1) & pfstr_Alineado(Mid(odynDetalle("USUARIO").Value, 1, 8) & _
                                               pfstr_Espacio(Mid(odynDetalle("USUARIO").Value, 9), 1), 39, centro) & "|" & _
                                               pfstr_Alineado(odynDetalle("COD_MAQUINA").Value, 10, centro) & "|" & _
                                               pfstr_Alineado(odynDetalle("COD_LIQUIDACION").Value, 17, centro) & "|" & _
                                               pfstr_Alineado(odynDetalle("FCH_REMESA").Value, 21, centro) & "|" & _
                                               pfstr_Alineado(odynDetalle("COD_REMESA").Value, 17, centro) & "|" & _
                                               pfstr_Alineado(Format(odynDetalle("IMP_SOLES").Value, "###,###0.00"), 11, centro) & "|" & _
                                               pfstr_Alineado(Format(odynDetalle("IMP_DOLARES").Value, "###,###0.00"), 11, centro) & "|"
                Printer.Print "|" & Space(1) & pfstr_Alineado(pfstr_Espacio(Mid(odynDetalle("USUARIO").Value, 9), 2), 39, centro) & "|" & _
                                               String(10, " ") & "|" & _
                                               String(17, " ") & "|" & _
                                               String(64, "-")
                
                ldblTotSoles = ldblTotSoles + odynDetalle("IMP_SOLES").Value
                ldblTotDolares = ldblTotDolares + odynDetalle("IMP_DOLARES").Value
            Else
                Printer.Print "|" & Space(1) & pfstr_Alineado(Mid(odynDetalle("USUARIO").Value, 1, 8) & _
                                               pfstr_Espacio(Mid(odynDetalle("USUARIO").Value, 9), 1) & _
                                               pfstr_Espacio(Mid(odynDetalle("USUARIO").Value, 9), 2), 39, centro) & "|" & _
                                               pfstr_Alineado(odynDetalle("COD_MAQUINA").Value, 10, centro) & "|" & _
                                               pfstr_Alineado(odynDetalle("COD_LIQUIDACION").Value, 17, centro) & "|" & _
                                               pfstr_Alineado(odynDetalle("FCH_REMESA").Value, 21, centro) & "|" & _
                                               pfstr_Alineado(odynDetalle("COD_REMESA").Value, 17, centro) & "|" & _
                                               pfstr_Alineado(Format(odynDetalle("IMP_SOLES").Value, "###,###0.00"), 11, centro) & "|" & _
                                               pfstr_Alineado(Format(odynDetalle("IMP_DOLARES").Value, "###,###0.00"), 11, centro) & "|"
                Printer.Print "|" & String(40, " ") & "|" & _
                                    String(10, " ") & "|" & _
                                    String(17, " ") & "|" & _
                                    String(64, "-")

                ldblTotSoles = ldblTotSoles + odynDetalle("IMP_SOLES").Value
                ldblTotDolares = ldblTotDolares + odynDetalle("IMP_DOLARES").Value
                
            End If
                antUsuario = nvoUsuario
                antCaja = nvoCaja
                antLiquidacion = nvoLiquidacion
        End If
        odynDetalle.MoveNext
        i = i + 1
    Loop
    
    
    Printer.Print String(135, "-")
    Printer.Print
    Printer.Print
    Printer.Print String(71, " ") & pfstr_Alineado("Total Soles", 21, centro) & String(20, " ") & _
                                    pfstr_Alineado("S/ " & Format(ldblTotSoles, "###,###0.00"), 23, centro)
    Printer.Print String(71, " ") & pfstr_Alineado("Total Dolares", 21, centro) & String(20, " ") & _
                                    pfstr_Alineado("$  " & Format(ldblTotDolares, "###,###0.00"), 23, centro)
    Printer.EndDoc

    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsImpresiones", Err.Description

End Function

Private Function pfstr_Espacio(ByVal vstrNombres As String, ByVal Opcion As Integer) As String
    Dim strCadena, strCad0, strcad1, strCad2 As String
    Dim i, intEspacio As Integer
    
    strCad0 = vstrNombres
    
    strcad1 = ""
    strCad2 = ""
    i = 1
    intEspacio = 0
    
    For i = 1 To Len(strCad0)
        If Mid(strCad0, i, 1) <> " " Then
            If intEspacio < 2 Then
               strcad1 = strcad1 & Mid(strCad0, i, 1)
               Debug.Print strcad1
            Else
                strCad2 = strCad2 & Mid(strCad0, i, 1)
                Debug.Print strCad2
            End If
        Else
            intEspacio = intEspacio + 1
            If intEspacio = 1 Then
                If Mid(strCad0, i, 1) = " " Then
                    strcad1 = strcad1 & Mid(strCad0, i, 1)
                    Debug.Print strcad1
                End If
           ElseIf intEspacio >= 2 Then
                If Mid(strCad0, i, 1) = " " Then
                    strCad2 = strCad2 & Mid(strCad0, i, 1)
                    Debug.Print strCad2
                End If
           End If
        End If
    Next i
    
    If Opcion = 1 Then
        strCadena = strcad1
    Else
        strCadena = strCad2
    End If
    pfstr_Espacio = strCadena
End Function
''''''''''
Public Function ProsegurNew(ByVal vstrCodLocal As String, _
                                      ByVal vstrCodRemito As String)

    Dim strCtdRem As String
    
    strCtdRem = objRemito.Dev_CtdReg_Remeda(vstrCodLocal, vstrCodRemito)
    strBanco = objRemito.BancoxBtl(vstrCodLocal)
    strPortavalor = objRemito.Portavalor(vstrCodLocal)
    strCtaCteSol = objRemito.CtaCtexBtl(vstrCodLocal, "S")
    
    Set odynReImp = objRemito.Re_Imp_Remito(vstrCodLocal, vstrCodRemito)
                
    strDesEfecSol = objLetra.NumAletra(odynReImp("TOT_SOLES").Value)
    strDesEfecDol = objLetra.NumAletra(odynReImp("TOT_DOLARES").Value)
                
    i = 1
    For i = 1 To 13
        Printer.Print
    Next i
    Printer.FontName = "Draft 17cpi"
    Printer.FontSize = 11
    Printer.Print Space(8) & "Cantidad Sobres =>" & " " & Trim(strCtdRem) & Space(30) & " " & Trim(strBanco)
    Printer.Print
    Printer.Print Space(42) & " " & Trim(odynReImp("COD_LOCAL").Value)
    Printer.Print
    Printer.Print Space(42) & " " & Trim(strPortavalor)
    Printer.Print
    Printer.Print Space(42) & " " & Trim(odynReImp("DIRECCION").Value) & Space(20) & Trim(odynReImp("DISTRITO").Value)
    Printer.Print
    For i = 1 To 4
        Printer.Print
    Next i
    If Not Val(odynReImp("TOT_SOLES").Value) = 0 Then
        Printer.Print Space(3) & " " & left(UCase((strDesEfecSol)) & " " & "NUEVOS SOLES" & Space(80), 80) & Space(1) & Format(odynReImp("TOT_SOLES").Value, "#,###,##0.00")
    Else
        Printer.Print ""
    End If
    Printer.Print
    If Not Val(odynReImp("TOT_DOLARES").Value) = 0 Then
        Printer.Print Space(3) & " " & left(UCase((strDesEfecDol)) & " " & "DOLARES AMERICANOS" & Space(80), 80) & Space(1) & Format(odynReImp("TOT_DOLARES").Value, "#,###,##0.00")
    Else
        Printer.Print ""
    End If
    Printer.Print
    Printer.Print
    'Printer.Print
    Printer.Print Space(3) & " " & Trim(odynReImp("PRECINTO").Value)
    Printer.EndDoc

    Exit Function
Error:
    Err.Raise Err.Number, "clsImpresiones", Err.Description

End Function


Public Function HermesNew(ByVal vstrCodLocal As String, _
                                    ByVal vstrCodRemito As String, _
                                    ByVal vstrMoneda As String)
    
    If vstrMoneda = "S" Then
        Set odynReImp = objRemito.Re_Imp_Remito_Hermes(vstrCodLocal, vstrCodRemito, vstrMoneda)
        strDesEfecSol = objLetra.NumAletra(odynReImp("TOT_SOLES").Value)
        strCtaCteSol = objRemito.CtaCtexBtl(vstrCodLocal, "S")
      Else
        Set odynReImp = objRemito.Re_Imp_Remito_Hermes(vstrCodLocal, vstrCodRemito, vstrMoneda)
        strDesEfecDol = objLetra.NumAletra(odynReImp("TOT_DOLARES").Value)
        strCtaCteDol = objRemito.CtaCtexBtl(vstrCodLocal, "D")
    End If
    
    strBanco = objRemito.BancoxBtl(objUsuario.CodigoLocal)

    i = 1
    For i = 1 To 4
        Printer.Print
    Next i
    Printer.FontName = "Draft 17cpi"
    Printer.FontSize = 11
    Printer.Print "   " & left((odynReImp("DIRECCION").Value) & "-" & odynReImp("DISTRITO").Value & Space(60), 60) & Space(2) & left((Format(odynReImp("FCH_REGISTRA").Value, "dd/mm/yyyy")) & Space(12), 12)
    Printer.Print
    Printer.Print "   " & left((" " & odynReImp("COD_LOCAL").Value) & " " & objUsuario.NombreLocal & Space(30), 30) & Space(50) & Trim("") & Space(10) & Trim("01")
    Printer.Print
    If Trim(vstrMoneda) = "S" Then
        Printer.Print " " & left((UCase(strDesEfecSol)) & " NUEVOS SOLES" & Space(60), 60) & Space(26) & right(Space(15) & "S/. " & Format(odynReImp("TOT_SOLES").Value, "#,###,##0.00"), 15)
    Else
        Printer.Print " " & left((UCase(strDesEfecDol)) & " DOLARES AMERICANOS" & Space(60), 60) & Space(26) & right(Space(15) & "$ " & Format(odynReImp("TOT_DOLARES").Value, "#,###,##0.00"), 15)
    End If
    Printer.Print ""
    Printer.Font.Size = 18
    Printer.Font.Bold = True
    Printer.Print ""
    If Trim(vstrMoneda) = "S" Then
       Printer.Print " " & " X" & "          " & "X"
     Else
       Printer.Print " " & " X" & "          "
    End If
    Printer.Font.Size = 11
    Printer.Font.Bold = False
    Printer.Print " "
    Printer.Print " "
    Printer.Print " "
    Printer.Print Space(60) & " " & left((Trim(strBanco)) & Space(60), 60)
    Printer.Print
    Printer.Print
    If Trim(vstrMoneda) = "S" Then
        Printer.Print Space(3) & " " & left((Trim(strCtaCteSol)) & Space(70), 70)
    Else
        Printer.Print Space(3) & " " & left((Trim(strCtaCteDol)) & Space(70), 70)
    End If
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
   
    Printer.Print Space(35) & " " & Trim("HERMES PROCESAMIENTO")
    For i = 1 To 4
        Printer.Print
    Next i
    Printer.FontSize = 18
    
    Printer.Print Space(10) & " " & IIf(vstrMoneda = "S", fncGuion(Trim(odynReImp("PRECINTO").Value), , ","), fncGuion(Trim(odynReImp("PRECINTO").Value), 1, ","))
    
    Printer.FontSize = 11
    Printer.EndDoc

    Exit Function
Error:
    Err.Raise Err.Number, "clsImpresiones", Err.Description

End Function


