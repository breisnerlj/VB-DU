VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private strNombre As String
Private strDireccion As String
Private strDistrito As String
Private strCodigoRespuesta As String
Private strSuministro As String
Private strRecibos As String
Dim xdbVPos As New XArrayDB
Dim iFila As Integer
Dim lblnDato As Boolean

Public Function Lista(Optional Codigo As String = "", Optional Tipo As String = "") As oraDynaset
    Set Lista = gclsOracle.FN_Cursor("BTLPROD.PKG_SERVICIO.FN_LISTA", 0, Codigo, Tipo)
End Function
Public Function ListaCobranza(ByVal Cia As String, ByVal CodigoLocal As String, ByVal TipoDocumento As String, ByVal NumeroDocumento As String) As oraDynaset
    Set ListaCobranza = gclsOracle.FN_Cursor("BTLPROD.PKG_SERVICIO.FN_LISTA_MOV", 0, Cia, CodigoLocal, TipoDocumento, NumeroDocumento)
End Function


Public Function ListaTipo(Optional Codigo As String = "") As oraDynaset
    Set ListaTipo = gclsOracle.FN_Cursor("BTLPROD.PKG_SERVICIO.FN_LISTA_TIPO", 0, Codigo)
End Function


Public Function ListaTipoCons(Optional Codigo As String = "") As oraDynaset
    Set ListaTipoCons = gclsOracle.FN_Cursor("BTLPROD.PKG_SERVICIO.FN_LISTA_TIPO_CONS", 0, Codigo)
End Function



Private Function DevuelveSuministro(ByVal Tipo As String, ByRef Criterio As String) As String
    Dim lenCriterio As Integer
    Criterio = Replace(Criterio, "-", "")
    lenCriterio = Len(Criterio)
    Select Case Tipo
        Case 0 'Codigo Barras
              If lenCriterio = LenCodBarraSedapal Then
                DevuelveSuministro = left(Criterio, LenNumSuminSedapal)
              Else
                MsgBox "El Codigo de barra no esta completo", vbCritical, App.ProductName
                Exit Function
              End If
              GoTo h
        Case 1 ' Suminstro
              If lenCriterio = LenNumSuminSedapal Or lenCriterio = LenNumSuminSedapal + 1 Then
                DevuelveSuministro = left(Criterio, LenNumSuminSedapal)
              Else
                MsgBox "El Número de Suminstro no esta completo", vbCritical, App.ProductName
                Exit Function
              End If
              GoTo h
        Case 2 ' Numero de Referencia
            If lenCriterio = LenRefCobroSedapal Then
                DevuelveSuministro = left(Criterio, LenNumSuminSedapal)
            Else
                MsgBox "El Número de Referencia no esta completo", vbCritical, App.ProductName
                Exit Function
            End If
            GoTo h
    End Select
    DevuelveSuministro = ""
h:
    
End Function
'************************************************************************
' Devuelve el digito de verificacion
'************************************************************************
Private Function DigitoVerificacion(strCadena As String) As String
   Dim Dig As String
   Dim Cadena1$, cadena2$
   Dim Proceso As Integer
   Dim DIG1, DIG2, DIG3 As Integer
   Dim suma$
   Dim i%
   
   If Not IsNumeric(strCadena) Then
      MsgBox "Verifique Parametro de ingreso", vbOKOnly, App.ProductName
      Exit Function
   End If
   If Len(strCadena) = 7 Then
      Cadena1 = "000" & strCadena
      Proceso = 10
   End If
   If Len(strCadena) = 10 Then
      Cadena1 = strCadena
      Proceso = 10
   End If
   If Len(strCadena) = 25 Then
      Cadena1 = strCadena
      Proceso = 25
   End If
   
   Select Case Proceso
      Case 10
           cadena2 = ""
           suma = 0
           For i = 1 To Len(Cadena1)
               If i Mod 2 <> 0 Then
                  DIG1 = Int(Val(Mid(Cadena1, i, 1)) / 10)
                  DIG2 = Val(Mid(Cadena1, i, 1)) Mod 10
               Else
                  DIG1 = Int(2 * Val(Mid(Cadena1, i, 1)) / 10)
                  DIG2 = 2 * Val(Mid(Cadena1, i, 1)) Mod 10
               End If
               DIG3 = DIG1 + DIG2
               suma = suma + DIG3
           Next i
           Dig = suma Mod 10
           DigitoVerificacion = (10 - Dig) Mod 10
      Case 25
          DigitoVerificacion = "99"
  End Select
End Function
Function ConectaLuzDelSur(ByVal Tipo As String, ByVal Criterio As String) As XArrayDB
    Dim rs As oraDynaset
    Dim xarrTemp As New XArrayDB
    xarrTemp.ReDim 0, -1, 0, 9
    Dim Suministro As String
    Suministro = Criterio
    Set rs = gclsOracle.FN_Cursor("BTLPROD.PKG_SERVICIO.FN_LISTA_LUZ_SUR", 0, Suministro)
    Dim i As Integer
    While Not rs.EOF
        xarrTemp.AppendRows
        xarrTemp(i, 0) = "" & rs("NRO_BOLETA").Value 'Número de Boleta
        xarrTemp(i, 1) = "" & rs("FCH_FAC").Value 'Fecha Factura
        xarrTemp(i, 2) = "" & rs("MTO_FACTURA").Value 'MontoFactura
        xarrTemp(i, 3) = "" & rs("FLG_VERIF_BOL").Value 'Flag de Verificacion
        xarrTemp(i, 4) = "" & rs("FLG_VERIF_SUMI").Value 'Flag de Verificacion Suministro
        xarrTemp(i, 5) = "" & rs("NRO_SUMINISTRO").Value 'Suministro
        xarrTemp(i, 6) = "" & rs("FCH_VENCIMIENTO").Value 'FechaVencimiento
        xarrTemp(i, 7) = "" & rs("MTO_MINIMO").Value 'MontoMinimo
        xarrTemp(i, 8) = "" & rs("DSC_NOMBRE").Value 'Nombre
        xarrTemp(i, 9) = "" & rs("FLG_ANT_DEUDA").Value 'Flag deuda anterior
        i = i + 1
        rs.MoveNext
    Wend
    Set ConectaLuzDelSur = xarrTemp
End Function
Function ConectaLuzDelSurDB(ByVal Tipo As String, ByVal Criterio As String) As oraDynaset
On Error GoTo handle
    Set ConectaLuzDelSurDB = gclsOracle.FN_Cursor("BTLPROD.PKG_SERVICIO.FN_LISTA_LUZ_SUR", 0, Criterio)
    Exit Function
handle:
MsgBox Err.Description, vbCritical, App.ProductName

End Function

Function ConectaSedapal(Tipo As String, Criterio As String)
Dim objHComVPos As Object
Dim IngResult  As Long
'Dim strCodResp As String
Dim Suministro As String
Dim strNomb$, strDir$, strDist$, strCodResp$, strRec$
Suministro = DevuelveSuministro(Tipo, Criterio)
If Suministro = "" Then Exit Function
Suministro = Suministro & DigitoVerificacion(Suministro)
strSuministro = Suministro
Dim objSedapal As New CC_Transaction



Set objHComVPos = CreateObject("HComVPos.CC_Transaction")

'IngResult = objHComVPos.b_consultasedapal("28585412", _
'                             "090", _
'                             "BTL090C1", _
'                             "20070301", _
'                             "000000", _
'                             "0", _
'                             strNombre, _
'                             strDireccion, _
'                             strDistrito, _
'                             strCodResp, _
'                             strRec)



IngResult = objSedapal.b_consultasedapal(Suministro, _
                             objUsuario.CodigoLocal, _
                             objUsuario.NombrePC, _
                             Format(objUsuario.sysdate, "yyyymmdd"), _
                             Format(objUsuario.sysdate, "hhmmss"), _
                             "0", _
                             strNombre, _
                             strDireccion, _
                             strDistrito, _
                             strCodResp, _
                             strRec)
If IngResult = -1 Then
            Select Case strCodResp
                Case "00"
                    'MsgBox "'" & strNomb & Chr(13) & strDir & Chr(13) & strRec & "'", vbInformation, "Transacción Correcta"
                    Call psub_Add_Sumi(strRec)
                Case "01"
                    MsgBox "Mensaje invalido", vbCritical, App.ProductName
                Case "10"
                    MsgBox "TimeOut en respuesta", vbInformation, App.ProductName
                Case "15"
                    MsgBox "Error en el envio", vbCritical, App.ProductName
                Case "45"
                    MsgBox "Error en la recepción", vbCritical, App.ProductName
                Case "81"
                    MsgBox "Error genérico", vbCritical, App.ProductName
                Case "82"
                    MsgBox "Fuera de servicio", vbCritical, App.ProductName
                Case "83"
                    MsgBox "Archivos cerrados o inhabilitados", vbCritical, App.ProductName
                Case "84"
                    MsgBox "Archivo no existe", vbInformation, App.ProductName
                Case "85"
                    MsgBox "Codigo del banco errado", vbCritical, App.ProductName
                Case "86"
                    MsgBox "Fecha invalida", vbCritical, App.ProductName
                Case "87"
                    MsgBox "Codigo de compañia errado", vbCritical, App.ProductName
                Case "88"
                    MsgBox "Contrato errado", vbCritical, App.ProductName
                Case "89"
                    MsgBox "Cliente sin deuda", vbInformation, App.ProductName
                Case "90"
                    MsgBox "Saldo a favor del cliente", vbInformation, App.ProductName
                Case "91"
                    MsgBox "Registro no existe en archivos loggins", vbCritical, App.ProductName
                Case "92"
                    MsgBox "Error Oracle", vbCritical, App.ProductName
            End Select
           Else
             MsgBox "No hay Conexiòn con SEDAPAL o" & Chr(13) & "Codigo de Barra del Recibo no existe", vbCritical, App.ProductName
             Exit Function
        End If
Set objSedapal = Nothing
End Function
Property Get Cliente() As String
    Cliente = strNombre
End Property
Property Get Direccion() As String
     Direccion = strDireccion
End Property
Property Get Distrito() As String
    Distrito = strDistrito
End Property
Property Get Suministro() As String
    Suministro = strSuministro
End Property

Property Get Recibos() As XArrayDB
Dim xarrTemp As New XArrayDB
    If strRecibos = "" Then Set Recibos = xarrTemp
    
    Dim strRec$, varImp$, varFecha$
    Dim varLong#
    Dim iFila As Integer
    iFila = 1
    xarrTemp.ReDim 0, -1, 0, 5
    strRec = strRecibos
    varLong = (Len(strRec) / 35)
    While varLong > 0
        xarrTemp.AppendRows 1
        xarrTemp(xarrTemp.UpperBound(1), 0) = Mid(strRec, iFila, 10) 'Suministro
        iFila = iFila + 10
        varFecha = Mid(strRec, iFila, 8)
        xarrTemp(xarrTemp.UpperBound(1), 1) = right(varFecha, 2) & "/" & Mid(varFecha, 5, 2) & "/" & Mid(varFecha, 1, 4) 'Fecha
        iFila = iFila + 8
        varImp = Mid(strRec, iFila, 15)
        xarrTemp(xarrTemp.UpperBound(1), 2) = CDbl(Mid(varImp, 1, 13)) & "." & right(varImp, 2)
        iFila = iFila + 15
        xarrTemp(xarrTemp.UpperBound(1), 3) = Mid(strRec, iFila, 1)
          iFila = iFila + 1
        xarrTemp(xarrTemp.UpperBound(1), 4) = Mid(strRec, iFila, 1)
        varLong = -1
    Wend
Set Recibos = xarrTemp

End Property

 Public Function GrabaAsignaSoat(ByVal vstrCodBtl As String, _
                                   ByVal vstrNumIni As String, _
                                   ByVal vstrNumFin As String, _
                                   ByVal vstrCodUsu As String) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant
    gvarValores = Array(vstrCodBtl, _
                        vstrNumIni, _
                        vstrNumFin, _
                        vstrCodUsu)
    
    gvarIO = Array(entrada, _
                   entrada, _
                   entrada, _
                   entrada)
                   
    GrabaAsignaSoat = gclsOracle.SP("BTLPROD.PKG_SERVICIO.SP_ASIGNA_FORMS_SOAT", _
                              gvarValores, _
                              gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsServicio", Err.Description
End Function

Public Function ListaSrvCob(ByVal vstrCia As String, _
                            ByVal vstrCodLocal As String, _
                            ByVal vstrFchIni As String, _
                            ByVal vstrFchFin As String, _
                            Optional vstrCodServ As String, _
                            Optional vstrCodUsuario As String, _
                            Optional vstrCodMoneda As String) As oraDynaset
On Error GoTo CtrlErr
    Set ListaSrvCob = gclsOracle.FN_Cursor("BTLPROD.PKG_SERVICIO.FN_LST_SRV_COB", 0, vstrCia, _
                                                                                     vstrCodLocal, _
                                                                                     vstrFchIni, _
                                                                                     vstrFchFin, _
                                                                                     vstrCodServ, _
                                                                                     vstrCodUsuario, _
                                                                                     vstrCodMoneda)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsServicio.ListaSrvCob", Err.Description

End Function


Private Sub psub_Add_Sumi(ByVal strRec$)
    Dim varIdRec$, varFecha$, varImp$, varSitSumi$, varTipRec$
    Dim varLong#
    
    varIdRec = "": varFecha = "": varImp = "": varSitSumi = "": varTipRec = ""
    iFila = 1
    
    ' Hace la reconeccion del array con la grilla '
      Call psub_Inicia_Array(xdbVPos)
    ' ******************************************* '
    
    varLong = (Len(strRec) / 35)
    While varLong > 0
        varIdRec = Mid(strRec, iFila, 10)
          iFila = iFila + 10
        varFecha = Mid(strRec, iFila, 8)
          iFila = iFila + 8
        varImp = Mid(strRec, iFila, 15)
          iFila = iFila + 15
        varSitSumi = Mid(strRec, iFila, 1)
          iFila = iFila + 1
        varTipRec = Mid(strRec, iFila, 1)
          iFila = iFila + 1
        
        '** Longitud de trama -1 **'
        varLong = varLong - 1
        '**************************'
        
        xdbVPos.AppendRows 1
        xdbVPos(xdbVPos.UpperBound(1), 0) = Trim(varIdRec)
        xdbVPos(xdbVPos.UpperBound(1), 1) = right(varFecha, 2) & "-" & Mid(varFecha, 5, 2) & "-" & Mid(varFecha, 1, 4)
        xdbVPos(xdbVPos.UpperBound(1), 2) = CDbl(Mid(varImp, 1, 13)) & "." & right(varImp, 2)
        xdbVPos(xdbVPos.UpperBound(1), 3) = Trim(varSitSumi)
        xdbVPos(xdbVPos.UpperBound(1), 4) = Trim(varTipRec)
    Wend
    frm_VTA_ServicioSedapal.GrdSuministros.Rebind
End Sub

Private Sub psub_Inicia_Array(vxdbArray As XArrayDB)
    vxdbArray.ReDim 0, -1, 0, 4
    frm_VTA_ServicioSedapal.GrdSuministros.Array1 = vxdbArray
    frm_VTA_ServicioSedapal.GrdSuministros.Rebind
End Sub




Public Function LenCodBarraSedapal() As Integer
    On Error GoTo CtrlErr
    LenCodBarraSedapal = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_SEDAPAL_LONGCOD_BARRA")
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsServicio.LenCodBarraSedapal", Err.Description
End Function

Public Function LenRefCobroSedapal() As Integer
    On Error GoTo CtrlErr
    LenRefCobroSedapal = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_SEDAPAL_LONGREF_COBRO")
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsServicio.LenRefCobroSedapal", Err.Description
End Function

Public Function LenNumSuminSedapal() As Integer
    On Error GoTo CtrlErr
    LenNumSuminSedapal = gclsOracle.FN_Valor("BTLPROD.PKG_CONSTANTES.CONS_SEDAPAL_LONGNUM_SUMIN")
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsServicio.LenNumSuminSedapal", Err.Description
End Function




 Public Function Modifica(ByVal CodigoServicio As String, _
                                   ByVal NumeroVoucher As String, _
                                   ByVal Cia As String, _
                                   ByVal TipoDocumento As String, _
                                   ByVal NumeroDocumento As String, _
                                   ByVal CodigoUsuario As String) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant
    gvarValores = Array(CodigoServicio, _
                        NumeroVoucher, _
                        Cia, _
                        TipoDocumento, _
                        NumeroDocumento, _
                        CodigoUsuario)
    
    gvarIO = Array(entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada)
                   
    Modifica = gclsOracle.SP("BTLPROD.PKG_SERVICIO.SP_MODIFICA", _
                              gvarValores, _
                              gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsServicio", Err.Description
End Function


Public Function ListaRecarga(Optional ByVal flgRecarga As Boolean = False) As oraDynaset
Dim strRecarga As String
On Error GoTo CtrlErr
    If flgRecarga = True Then strRecarga = flgRecarga
    Set ListaRecarga = gclsOracle.FN_Cursor("BTLPROD.PKG_SERVICIO.FN_LISTA_PIN", 0, strRecarga)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsServicio.ListaRecarga", Err.Description
End Function


Public Function ListaRecargaX(Optional ByVal CodigoProducto As String = "") As oraDynaset
On Error GoTo CtrlErr
    
    godbVentas.Parameters.Remove "X_P_COD_PRODUCTO"
    godbVentas.Parameters.Add "X_P_COD_PRODUCTO", CodigoProducto, ORAPARM_INPUT
    Set ListaRecargaX = godbVentas.CreatePlsqlDynaset("BEGIN :SALIDA := BTLPROD.PKG_SERVICIO.FN_LISTA_PIN_X(:X_P_COD_PRODUCTO);END;", "SALIDA", ORADYN_READONLY)
    'Set ListaRecargaX = gclsOracle.FN_Cursor("BTLPROD.PKG_SERVICIO.FN_LISTA_PIN_X", 0, CodigoProducto)
    godbVentas.Parameters.Remove "X_P_COD_PRODUCTO"
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsServicio.ListaRecargaX", Err.Description
End Function


Public Function TotListaSrvCob(ByVal vstrCia As String, _
                            ByVal vstrCodLocal As String, _
                            ByVal vstrFchIni As String, _
                            ByVal vstrFchFin As String, _
                            Optional vstrCodServ As String, _
                            Optional vstrCodUsuario As String, _
                            Optional vstrCodMoneda As String) As oraDynaset
On Error GoTo CtrlErr
    Set TotListaSrvCob = gclsOracle.FN_Cursor("BTLPROD.PKG_SERVICIO.FN_TOT_LST_SRV_COB", 0, vstrCia, _
                                                                                     vstrCodLocal, _
                                                                                     vstrFchIni, _
                                                                                     vstrFchFin, _
                                                                                     vstrCodServ, _
                                                                                     vstrCodUsuario, _
                                                                                     vstrCodMoneda)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsServicio.TotListaSrvCob", Err.Description

End Function


Public Function logRecarga(ByVal Respuesta As String, _
                           ByVal CodigoUsuario As String, _
                           ByVal TramaEnvio As String, _
                           ByVal CodigoLiquidacion As String, _
                           ByVal Monto As String, _
                           ByVal Telefono As String, _
                           ByVal Servicio As String, _
                           ByVal Cia As String, _
                           ByVal TipoDocumento As String, _
                           ByVal NumeroDocumento As String, _
                           ByVal log_Recarga As String, _
                           ByVal Secuencia As String _
                           ) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant
    gvarValores = Array(Respuesta, _
                        CodigoUsuario, _
                        TramaEnvio, _
                        CodigoLiquidacion, _
                        Monto, _
                        Telefono, _
                        Servicio, _
                        Cia, _
                        TipoDocumento, _
                        NumeroDocumento, _
                        log_Recarga, _
                        Secuencia _
                        )
    
    gvarIO = Array(entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada, _
                   entrada _
                   )
                   
    logRecarga = gclsOracle.SP("BTLPROD.PKG_SERVICIO.SP_GRABA_LOG_RECARGA", _
                              gvarValores, _
                              gvarIO)

    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsServicio", Err.Description
End Function


Public Function ListaServiciosAsoc(ByVal Cia As String, _
                            ByVal CodigoTipoDocumento As String, _
                            ByVal NumeroDocumento As String) As oraDynaset
On Error GoTo CtrlErr
    Set ListaServiciosAsoc = gclsOracle.FN_Cursor("BTLPROD.PKG_SERVICIO.FN_SERVICIO_ASOC", 0, Cia, _
                                                                                     CodigoTipoDocumento, _
                                                                                     NumeroDocumento)
    
    Exit Function
CtrlErr:
    Err.Raise Err.Number, "clsServicio.ListaServiciosAsoc", Err.Description
End Function


Public Function ErrorNAVSAT(ByVal vstrCodigo As String) As String
    ErrorNAVSAT = "" & gclsOracle.FN_Valor("BTLPROD.PKG_SERVICIO.FN_ERROR_NAVSAT", vstrCodigo)
End Function

'Public Function DevProducto_SOAT() As oraDynaset
'    On Error GoTo CtrlErr
'    DevProducto_SOAT = gclsOracle.FN_Valor("BTLPROD.PKG_SOAT.FN_CERTIF_ACTIVO")
'    Exit Function
'CtrlErr:
'    Err.Raise Err.Number, "clsServicio.DevProducto_SOAT", Err.Description
'End Function



Public Property Get Codigo_SOAT() As String
    Codigo_SOAT = gclsOracle.FN_Valor("CMR.PKG_SOAT.FN_CODIGO_SOAT")
End Property

Public Property Get Ctd_Fraccion() As String
    Ctd_Fraccion = gclsOracle.FN_Valor("CMR.PKG_SOAT.FN_CANT_FRACCIONA")
End Property






 Public Function ConfirmaVenta(ByVal CodigoEmpresa As String, _
                               ByVal TipoDocumento As String, _
                               ByVal CodigoLocal As String, _
                               ByVal NumeroDocumento As String, _
                               ByVal NumeroAprobacion As String, _
                               ByVal NumeroVoucher As String, _
                               ByVal Suministro As String, _
                               ByVal Respuesta As String, _
                               ByVal CodigoUsuario As String, _
                               ByVal TramaEnvio As String, _
                               ByVal CodigoLiquidacion As String, _
                               ByVal MontoCarga As String, _
                               ByVal NumeroTelefono As String, _
                               ByVal CodigoProducto As String, _
                               ByVal Secuencia As String, _
                               ByVal Confirmacion As String, _
                               ByVal CodigoTipoRecibo As String, _
                               ByVal NumRecibo As String, _
                               ByRef DocAnulados As String _
                               ) As String
On Error GoTo CtrlErr
Dim gvarValores As Variant
Dim gvarIO  As Variant
    gvarValores = Array(CodigoEmpresa, _
                         TipoDocumento, _
                         CodigoLocal, _
                         NumeroDocumento, _
                         NumeroAprobacion, _
                         NumeroVoucher, _
                         Suministro, _
                         Respuesta, _
                         CodigoUsuario, _
                         TramaEnvio, _
                         CodigoLiquidacion, _
                         MontoCarga, _
                         NumeroTelefono, _
                         CodigoProducto, _
                         Secuencia, _
                         Confirmacion, _
                         CodigoTipoRecibo, _
                         NumRecibo, _
                         DocAnulados _
                         )
    
    gvarIO = Array(entrada, _
                         entrada, _
                         entrada, _
                         entrada, _
                         entrada, _
                         entrada, _
                         entrada, _
                         entrada, _
                         entrada, _
                         entrada, _
                         entrada, _
                         entrada, _
                         entrada, _
                         entrada, _
                         entrada, _
                         entrada, _
                         entrada, _
                         entrada, _
                         Salida)
                   
    ConfirmaVenta = gclsOracle.SP("BTLPROD.PKG_SERVICIO.SP_CONF_RECARGA", _
                              gvarValores, _
                              gvarIO)
    DocAnulados = gvarValores(18)
    Exit Function
CtrlErr:
    Err.Raise Err.Description, "clsServicio", Err.Description
End Function

